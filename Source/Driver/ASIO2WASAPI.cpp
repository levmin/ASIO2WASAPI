/*  ASIO2WASAPI Universal ASIO Driver
    Copyright (C) Lev Minkovsky
    Copyright (C) Zoltan Bacsko (Falcosoft) 2023-2026
    
    This file is part of ASIO2WASAPI.

    ASIO2WASAPI is free software; you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation; either version 2 of the License, or
    (at your option) any later version.

    ASIO2WASAPI is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with ASIO2WASAPI; if not, write to the Free Software
    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA
*/

#include "stdafx.h"
#include <math.h>
#include <mmsystem.h>
#include <stdio.h>
#include <string.h>
#include <new>
#include <Mmdeviceapi.h>
#include <Audioclient.h>
#include <versionhelpers.h>
#include <process.h>
#include "Avrt.h" //used for AvSetMmThreadCharacteristics
#include <Functiondiscoverykeys_devpkey.h>
#include "ASIO2WASAPI.h"
#include "resource.h"
#include "version.h"
#include "StreamFormats.h"

extern HINSTANCE g_hinstDLL;

CLSID CLSID_ASIO2WASAPI_DRIVER = { 0x3981c4c8, 0xfe12, 0x4b0f, { 0x98, 0xa0, 0xd1, 0xb6, 0x67, 0xbd, 0xa6, 0x15 } };

const CLSID CLSID_MMDeviceEnumerator = __uuidof(MMDeviceEnumerator);
const IID IID_IMMDeviceEnumerator = __uuidof(IMMDeviceEnumerator);
const IID IID_IAudioClient = __uuidof(IAudioClient);
const IID IID_IAudioRenderClient = __uuidof(IAudioRenderClient);

const char * szPrefsRegKey = "Software\\ASIO2WASAPI";

wchar_t* GetFileVersion(wchar_t* result, unsigned int buffSize)
{
    wcscat_s(result, buffSize, L"version: " stringify(VERSION_MAJOR) L"." stringify(VERSION_MINOR) L"." stringify(VERSION_PATCH));
    return result;
}

class CReleaser 
{
    IUnknown * m_pUnknown;
public:
    explicit CReleaser(IUnknown * pUnk) : m_pUnknown(pUnk) {}
    void deactivate() {m_pUnknown = NULL;}
    void reset(IUnknown * pUnk) 
    {
        SAFE_RELEASE(m_pUnknown)
        m_pUnknown = pUnk;
    }
    ~CReleaser() 
    {
        SAFE_RELEASE(m_pUnknown)
    }
};

class CHandleCloser
{
    HANDLE m_h;
public:
    explicit CHandleCloser(HANDLE h) : m_h(h) {}
    ~CHandleCloser() 
    {
        if (m_h != NULL)
            CloseHandle(m_h);
    }
};

class CAvReverter
{
    HANDLE m_handle;
public:
    explicit CAvReverter(HANDLE handle) : m_handle(handle) {}
    ~CAvReverter() { if (m_handle) AvRevertMmThreadCharacteristics(m_handle); }
};

class CAudioStopper
{
    IAudioClient* m_client;
public:
    explicit CAudioStopper(IAudioClient* client) : m_client(client) {}
    ~CAudioStopper()
    {
        if (m_client) { m_client->Stop(); m_client->Reset(); }
    }
};

/// CMMNotificationClient

CMMNotificationClient::CMMNotificationClient(ASIO2WASAPI* asio2Wasapi) :
    _cRef(1),
    _pEnumerator(NULL)
{
    _asio2Wasapi = asio2Wasapi;
}

CMMNotificationClient::~CMMNotificationClient()
{
    SAFE_RELEASE(_pEnumerator)
}

ULONG __stdcall CMMNotificationClient::AddRef()
{
    return InterlockedIncrement(&_cRef);
}

ULONG __stdcall CMMNotificationClient::Release()
{
    ULONG ulRef = InterlockedDecrement(&_cRef);
    if (0 == ulRef)
    {
        delete this;
    }
    return ulRef;
}

HRESULT __stdcall CMMNotificationClient::QueryInterface(REFIID riid, VOID** ppvInterface)
{
    if (IID_IUnknown == riid)
    {
        AddRef();
        *ppvInterface = (IUnknown*)this;
    }
    else if (__uuidof(IMMNotificationClient) == riid)
    {
        AddRef();
        *ppvInterface = (IMMNotificationClient*)this;
    }
    else
    {
        *ppvInterface = NULL;
        return E_NOINTERFACE;
    }
    return S_OK;
}

HRESULT __stdcall CMMNotificationClient::OnDefaultDeviceChanged(EDataFlow flow, ERole role, LPCWSTR pwstrDeviceId)
{
    if (role == eConsole &&
        ((flow == eRender && _asio2Wasapi->getUseDefaultDevice()) ||
         (flow == eCapture && _asio2Wasapi->getUseDefaultInputDevice())))
    {
        ASIOCallbacks* callbacks = _asio2Wasapi->getCallbacks();
        if (callbacks && callbacks->asioMessage)
            callbacks->asioMessage(kAsioResetRequest, 0, NULL, NULL);
    }
    return S_OK;
}

HRESULT __stdcall CMMNotificationClient::OnDeviceAdded(LPCWSTR pwstrDeviceId)
{
    return S_OK;
}

HRESULT __stdcall CMMNotificationClient::OnDeviceRemoved(LPCWSTR pwstrDeviceId)
{
    return S_OK;
}

HRESULT __stdcall CMMNotificationClient::OnDeviceStateChanged(LPCWSTR pwstrDeviceId, DWORD dwNewState)
{
    return S_OK;
}

inline HRESULT __stdcall CMMNotificationClient::OnPropertyValueChanged(LPCWSTR pwstrDeviceId, const PROPERTYKEY key)
{
    return S_OK;
}

/// ASIO2WASAPI

inline long ASIO2WASAPI::refTimeToBufferSize(REFERENCE_TIME time) const
{
    const double REFTIME_UNITS_PER_SECOND = 10000000.;
    return static_cast<long>(ceil(m_nSampleRate * ( time / REFTIME_UNITS_PER_SECOND ) ));
}

inline REFERENCE_TIME ASIO2WASAPI::bufferSizeToRefTime(long bufferSize) const
{
    const double REFTIME_UNITS_PER_SECOND = 10000000.;
    return static_cast<REFERENCE_TIME>(ceil(bufferSize / (m_nSampleRate / REFTIME_UNITS_PER_SECOND) ));
}

const double twoRaisedTo32 = 4294967296.;
const double twoRaisedTo32Reciprocal = 1. / twoRaisedTo32;

inline void getNanoSeconds (ASIOTimeStamp* ts)
{
    double nanoSeconds = (double)((unsigned long)timeGetTime ()) * 1000000.;
	ts->hi = (unsigned long)(nanoSeconds / twoRaisedTo32);
	ts->lo = (unsigned long)(nanoSeconds - (ts->hi * twoRaisedTo32));
}

inline vector<wchar_t> getDeviceId(IMMDevice * pDevice)
{
    vector<wchar_t> id;
    LPWSTR pDeviceId = NULL;
    HRESULT hr = pDevice->GetId(&pDeviceId);
    if (FAILED(hr)) return id;
    size_t nDeviceIdLength=wcslen(pDeviceId);
    if (nDeviceIdLength == 0) { CoTaskMemFree(pDeviceId); return id; }
    id.resize(nDeviceIdLength+1);
    wcscpy_s(&id[0],nDeviceIdLength+1,pDeviceId);
    CoTaskMemFree(pDeviceId);
    pDeviceId = NULL;
    return id;
}

BOOL IsFormatSupported (IMMDevice* pDevice, WORD nChannels, DWORD nSampleRate, AUDCLNT_SHAREMODE shareMode, BOOL doResampling)
{
    if (!pDevice)
        return NULL;

    IAudioClient* pAudioClient = NULL;
    HRESULT hr = pDevice->Activate(
        IID_IAudioClient, CLSCTX_ALL,
        NULL, (void**)&pAudioClient);
    if (FAILED(hr) || !pAudioClient)
        return NULL;
    CReleaser r(pAudioClient);

    WAVEFORMATEX* devFormat;
    hr = pAudioClient->GetMixFormat(&devFormat);

    if (SUCCEEDED(hr) && !shareMode && doResampling && (nChannels <= devFormat->nChannels))
    {
        CoTaskMemFree(devFormat);
        return TRUE;
    }

    DWORD dwChannelMask = 0;
    if (SUCCEEDED(hr) && devFormat->nChannels == nChannels &&
        devFormat->wFormatTag == WAVE_FORMAT_EXTENSIBLE &&
        devFormat->cbSize >= sizeof(WAVEFORMATEXTENSIBLE) - sizeof(WAVEFORMATEX))
    {
        WAVEFORMATEXTENSIBLE* tmpDevFormat = (WAVEFORMATEXTENSIBLE*)devFormat;
        dwChannelMask = tmpDevFormat->dwChannelMask;
    }
    else
    {
        // create a reasonable channel mask
        DWORD bit = 1;
        for (int i = 0; i < nChannels; i++)
        {
            dwChannelMask |= bit;
            bit <<= 1;
        }
    }
    if (SUCCEEDED(hr)) CoTaskMemFree(devFormat);
    

    WAVEFORMATEXTENSIBLE waveFormat = {0};
    //try 32-bit first
    waveFormat.Format.wFormatTag = WAVE_FORMAT_EXTENSIBLE;
    waveFormat.Format.nChannels = nChannels;
    waveFormat.Format.nSamplesPerSec = nSampleRate;
    waveFormat.Format.wBitsPerSample = 32;
    waveFormat.Format.nBlockAlign = waveFormat.Format.wBitsPerSample * waveFormat.Format.nChannels / 8;
    waveFormat.Format.nAvgBytesPerSec = waveFormat.Format.nSamplesPerSec * waveFormat.Format.nBlockAlign;
    waveFormat.Format.cbSize = sizeof(WAVEFORMATEXTENSIBLE) - sizeof(WAVEFORMATEX);
    waveFormat.Samples.wValidBitsPerSample = waveFormat.Format.wBitsPerSample;
    waveFormat.dwChannelMask = dwChannelMask;
    //test native support for 32-bit float first
    waveFormat.SubFormat = KSDATAFORMAT_SUBTYPE_IEEE_FLOAT;
    
    WAVEFORMATEX* closeMatch = NULL;
    hr = pAudioClient->IsFormatSupported(shareMode, (WAVEFORMATEX*)&waveFormat, shareMode ? NULL : &closeMatch);
    if(SUCCEEDED(hr) && !shareMode) CoTaskMemFree(closeMatch);
    if (hr == S_OK)
        return TRUE;
   
    //then try integer formats, first 32-bit int
    waveFormat.SubFormat = KSDATAFORMAT_SUBTYPE_PCM;
   
    closeMatch = NULL;
    hr = pAudioClient->IsFormatSupported(shareMode, (WAVEFORMATEX*)&waveFormat, shareMode ? NULL : &closeMatch);
    if (SUCCEEDED(hr) && !shareMode) CoTaskMemFree(closeMatch);
    if (hr == S_OK)
        return TRUE;

    //try 24-bit containered next
    waveFormat.Samples.wValidBitsPerSample = 24;

    closeMatch = NULL;
    hr = pAudioClient->IsFormatSupported(shareMode, (WAVEFORMATEX*)&waveFormat, shareMode ? NULL : &closeMatch);
    if (SUCCEEDED(hr) && !shareMode) CoTaskMemFree(closeMatch);
    if (hr == S_OK)
        return TRUE;

    //try 24-bit packed next
    waveFormat.Format.wBitsPerSample = 24;
    waveFormat.Format.nBlockAlign = waveFormat.Format.wBitsPerSample * waveFormat.Format.nChannels / 8;
    waveFormat.Format.nAvgBytesPerSec = waveFormat.Format.nSamplesPerSec * waveFormat.Format.nBlockAlign;
    waveFormat.Samples.wValidBitsPerSample = waveFormat.Format.wBitsPerSample;

    closeMatch = NULL;
    hr = pAudioClient->IsFormatSupported(shareMode, (WAVEFORMATEX*)&waveFormat, shareMode ? NULL : &closeMatch);
    if (SUCCEEDED(hr) && !shareMode) CoTaskMemFree(closeMatch);
    if (hr == S_OK)
        return TRUE;

    //finally, try 16-bit   
    waveFormat.Format.wBitsPerSample = 16;
    waveFormat.Format.nBlockAlign = waveFormat.Format.wBitsPerSample * waveFormat.Format.nChannels / 8;
    waveFormat.Format.nAvgBytesPerSec = waveFormat.Format.nSamplesPerSec * waveFormat.Format.nBlockAlign;
    waveFormat.Samples.wValidBitsPerSample = waveFormat.Format.wBitsPerSample;

    closeMatch = NULL;
    hr = pAudioClient->IsFormatSupported(shareMode, (WAVEFORMATEX*)&waveFormat, shareMode ? NULL : &closeMatch);
    if (SUCCEEDED(hr) && !shareMode) CoTaskMemFree(closeMatch);
    if (hr == S_OK)
        return TRUE;

    return FALSE;  

}

IAudioClient* getAudioClient(IMMDevice* pDevice, WAVEFORMATEX* pWaveFormat, int& bufferSize, AUDCLNT_SHAREMODE shareMode, BOOL doResampling, BOOL doLowLatency)

{
    if (!pDevice || !pWaveFormat)
        return NULL;

    UINT tmpBuffSize = 0;
    IAudioClient * pAudioClient = NULL;
    HRESULT hr = pDevice->Activate(
                    IID_IAudioClient, CLSCTX_ALL,
                    NULL, (void**)&pAudioClient);
    if (FAILED(hr) || !pAudioClient)
        return NULL;
    CReleaser r(pAudioClient);
   
    WAVEFORMATEX* closeMatch = NULL;
    hr = pAudioClient->IsFormatSupported(shareMode, pWaveFormat, shareMode ? NULL : &closeMatch);
    if (SUCCEEDED(hr) && !shareMode) CoTaskMemFree(closeMatch);

    if (hr == AUDCLNT_E_UNSUPPORTED_FORMAT) return NULL;

    if (!shareMode && doLowLatency) //Win10+ special low latency shared mode 
    {
        IAudioClient3* pAudioClient3 = NULL;
        hr = pAudioClient->QueryInterface(&pAudioClient3);
        if (FAILED(hr) || !pAudioClient3) return NULL;

        CReleaser ra3(pAudioClient3);
        UINT32 minPeriod, maxPeriod, fundamentalPeriod, defaultPeriod;
        hr = pAudioClient3->GetSharedModeEnginePeriod(pWaveFormat, &defaultPeriod, &fundamentalPeriod, &minPeriod, &maxPeriod);
        if (FAILED(hr) || !fundamentalPeriod) return NULL;
        UINT32 buffsizeInSamples = (UINT32)(bufferSize * (pWaveFormat->nSamplesPerSec * 0.001));
        while (minPeriod < (UINT32)(buffsizeInSamples * 0.44)) minPeriod += fundamentalPeriod;
        if (minPeriod > maxPeriod) minPeriod = maxPeriod;
        hr = pAudioClient3->InitializeSharedAudioStream(AUDCLNT_STREAMFLAGS_EVENTCALLBACK, minPeriod, pWaveFormat, NULL);
        if (hr == AUDCLNT_E_ENGINE_PERIODICITY_LOCKED)
        {
            UINT32 currentPeriod;
            WAVEFORMATEX* format = NULL;
            HRESULT periodResult = pAudioClient3->GetCurrentSharedModeEnginePeriod(&format, &currentPeriod);
            CoTaskMemFree(format);
            if (SUCCEEDED(periodResult))
                hr = pAudioClient3->InitializeSharedAudioStream(AUDCLNT_STREAMFLAGS_EVENTCALLBACK, currentPeriod, pWaveFormat, NULL);
        }
        if (FAILED(hr)) return NULL;
    }
    else
    {
        //calculate buffer size and duration
        REFERENCE_TIME hnsMinimumDuration = 0;
        if (shareMode)
            hr = pAudioClient->GetDevicePeriod(NULL, &hnsMinimumDuration); //Actually 2nd parameter should be used for exclusive mode streams...
        else
            hr = pAudioClient->GetDevicePeriod(&hnsMinimumDuration, NULL);

        if (FAILED(hr)) return NULL;

        if (shareMode)
            hnsMinimumDuration = max(hnsMinimumDuration, (REFERENCE_TIME)bufferSize * 10000);
        else
            hnsMinimumDuration = max(hnsMinimumDuration * 2, (REFERENCE_TIME)bufferSize * 10000);


        hr = pAudioClient->Initialize(
            shareMode,
            AUDCLNT_STREAMFLAGS_EVENTCALLBACK | (!shareMode && doResampling ? (AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM | AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY) : 0),
            hnsMinimumDuration,
            shareMode ? hnsMinimumDuration : 0,
            pWaveFormat,
            NULL);

        if (hr == AUDCLNT_E_BUFFER_SIZE_NOT_ALIGNED)
        {
            hr = pAudioClient->GetBufferSize(&tmpBuffSize);
            if (FAILED(hr)) return NULL;

            const double REFTIME_UNITS_PER_SECOND = 10000000.;
            REFERENCE_TIME hnsAlignedDuration = static_cast<REFERENCE_TIME>((REFTIME_UNITS_PER_SECOND / pWaveFormat->nSamplesPerSec * tmpBuffSize) + 0.5);
            r.deactivate();
            SAFE_RELEASE(pAudioClient);
            hr = pDevice->Activate(IID_IAudioClient, CLSCTX_ALL, NULL, (void**)&pAudioClient);
            if (FAILED(hr) || !pAudioClient)
                return NULL;
            r.reset(pAudioClient);
            hr = pAudioClient->Initialize(
                shareMode,
                AUDCLNT_STREAMFLAGS_EVENTCALLBACK | (!shareMode && doResampling ? (AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM | AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY) : 0),
                hnsAlignedDuration,
                shareMode ? hnsAlignedDuration : 0,
                pWaveFormat,
                NULL);
        }

        if (FAILED(hr)) return NULL;
    }
    
    pAudioClient->GetBufferSize(&tmpBuffSize); //calculate real latency/buffer size
    bufferSize = (int)round(tmpBuffSize / (pWaveFormat->nSamplesPerSec * 0.001));

    r.deactivate();
    return pAudioClient;
}

static bool MakeStreamFormat(IMMDevice* device, int channels, int rate, WAVEFORMATEXTENSIBLE& format)
{
    IAudioClient* client = NULL;
    if (FAILED(device->Activate(IID_IAudioClient, CLSCTX_ALL, NULL, (void**)&client))) return false;
    CReleaser release(client);
    WAVEFORMATEX* mix = NULL;
    HRESULT hr = client->GetMixFormat(&mix);
    memset(&format, 0, sizeof(format));
    format.Format.wFormatTag = WAVE_FORMAT_EXTENSIBLE;
    format.Format.nChannels = (WORD)channels;
    format.Format.nSamplesPerSec = rate;
    format.Format.cbSize = sizeof(WAVEFORMATEXTENSIBLE) - sizeof(WAVEFORMATEX);
    if (SUCCEEDED(hr) && mix->nChannels == channels && mix->wFormatTag == WAVE_FORMAT_EXTENSIBLE &&
        mix->cbSize >= sizeof(WAVEFORMATEXTENSIBLE) - sizeof(WAVEFORMATEX))
        format.dwChannelMask = ((WAVEFORMATEXTENSIBLE*)mix)->dwChannelMask;
    else
        for (int i = 0; i < channels && i < 32; ++i) format.dwChannelMask |= (DWORD)1 << i;
    CoTaskMemFree(mix);
    return true;
}

static void SetStreamSampleFormat(WAVEFORMATEXTENSIBLE& format, const StreamFormatCandidate& candidate)
{
    format.Format.wBitsPerSample = candidate.containerBits;
    format.Samples.wValidBitsPerSample = candidate.validBits;
    format.Format.nBlockAlign = candidate.containerBits / 8 * format.Format.nChannels;
    format.Format.nAvgBytesPerSec = format.Format.nSamplesPerSec * format.Format.nBlockAlign;
    format.SubFormat = candidate.floatingPoint ? KSDATAFORMAT_SUBTYPE_IEEE_FLOAT : KSDATAFORMAT_SUBTYPE_PCM;
}

static bool SupportsExactStreamFormat(IMMDevice* device, WAVEFORMATEXTENSIBLE& format, AUDCLNT_SHAREMODE mode)
{
    IAudioClient* client = NULL;
    if (FAILED(device->Activate(IID_IAudioClient, CLSCTX_ALL, NULL, (void**)&client))) return false;
    CReleaser release(client);
    WAVEFORMATEX* closest = NULL;
    HRESULT hr = client->IsFormatSupported(mode, &format.Format, mode == AUDCLNT_SHAREMODE_SHARED ? &closest : NULL);
    CoTaskMemFree(closest);
    return hr == S_OK; // a closest match is not a common exact format
}

static void ReleaseInitializedAudioClient(IAudioClient*& client)
{
    if (!client) return;
    // Keep the event alive during Release (avoids slow teardown on Win7).
    HANDLE event = CreateEvent(NULL, FALSE, FALSE, NULL);
    CHandleCloser closeEvent(event);
    if (event) client->SetEventHandle(event);
    SAFE_RELEASE(client)
}

BOOL FindStreamFormat(IMMDevice* pDevice, int nChannels, int nSampleRate, int& nbufferSize,
    AUDCLNT_SHAREMODE shareMode, BOOL doResampling, BOOL doLowLatency,
    WAVEFORMATEXTENSIBLE* pwfxt = NULL, IAudioClient** ppAudioClient = NULL,
    IMMDevice* peerDevice = NULL, int peerChannels = 0, WAVEFORMATEXTENSIBLE* peerSelectedFormat = NULL)
{
    if (!pDevice) return FALSE;
    WAVEFORMATEXTENSIBLE waveFormat, peerFormat;
    if (!MakeStreamFormat(pDevice, nChannels, nSampleRate, waveFormat)) return FALSE;
    if (peerDevice && !MakeStreamFormat(peerDevice, peerChannels, nSampleRate, peerFormat)) return FALSE;
    IAudioClient* pAudioClient = NULL;
    int selectedDuration = nbufferSize;
    auto tryFormat = [&](const StreamFormatCandidate& candidate) -> bool
    {
        SetStreamSampleFormat(waveFormat, candidate);
        IAudioClient* peerClient = NULL;
        int duration = nbufferSize;
        if (peerDevice)
        {
            SetStreamSampleFormat(peerFormat, candidate);
            if (!SupportsExactStreamFormat(pDevice, waveFormat, shareMode) ||
                !SupportsExactStreamFormat(peerDevice, peerFormat, shareMode)) return false;
            int peerDuration = nbufferSize;
            peerClient = getAudioClient(peerDevice, &peerFormat.Format, peerDuration, shareMode, doResampling, doLowLatency);
            if (!peerClient) return false;
        }
        // Validate simultaneous initialization, then release the capture probe.
        pAudioClient = getAudioClient(pDevice, &waveFormat.Format, duration, shareMode, doResampling, doLowLatency);
        ReleaseInitializedAudioClient(peerClient);
        if (!pAudioClient) return false;
        selectedDuration = duration;
        return true;
    };
    if (!SelectStreamFormat(tryFormat)) return FALSE;
    nbufferSize = selectedDuration;
    if (pwfxt) *pwfxt = waveFormat;
    if (peerSelectedFormat && peerDevice) *peerSelectedFormat = peerFormat;
    if (ppAudioClient) *ppAudioClient = pAudioClient;
    else ReleaseInitializedAudioClient(pAudioClient);
    return TRUE;
}

CUnknown* ASIO2WASAPI::CreateInstance (LPUNKNOWN pUnk, HRESULT *phr)
{
	return static_cast<CUnknown*>(new ASIO2WASAPI (pUnk,phr));
};

STDMETHODIMP ASIO2WASAPI::NonDelegatingQueryInterface (REFIID riid, void ** ppv)
{
	if (riid == CLSID_ASIO2WASAPI_DRIVER)
	{
		return GetInterface (this, ppv);
	}
	return CUnknown::NonDelegatingQueryInterface (riid, ppv);
}

ASIOSampleType ASIO2WASAPI::getASIOSampleType(bool input) const
{
    // REAPER and other hosts require one ASIO sample type for both streams.
    // Input WASAPI packets are converted to the unchanged output representation.
    (void)input;
    const WAVEFORMATEXTENSIBLE& format = m_waveFormat;
    if (format.SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT) return ASIOSTFloat32LSB;

    switch (format.Format.wBitsPerSample)
    {
        case 16: return  ASIOSTInt16LSB;
        case 24: return  ASIOSTInt24LSB;
        case 32: return  ASIOSTInt32LSB;
           /* switch (m_waveFormat.Samples.wValidBitsPerSample)
            {
                case 32: return ASIOSTInt32LSB;
                case 24: return ASIOSTInt32LSB; //falco: In case of 24-bit data Windows simply chops the last 8 bits. No special alignment needed. ASIOSTInt32LSB24 is simply wrong. 
                default: return ASIOSTLastEntry ;
            }
           */
        default: return ASIOSTLastEntry;
    }
}

const char * szChannelRegValName = "Channels";
const char * szSampRateRegValName = "Sample Rate";
const char * szBufferSizeRegValName = "Buffer Size";
const char * szWasapiModeRegValName = "WASAPI Exclusive Mode";
const char * szEnableResamplingRegValName = "Enable Shared Mode Resampling";
const char * szEnableLowLatencygRegValName = "Enable Shared Mode Low Latency";

const char* szInputEnabled = "Input Enabled";
const char* szPreferCommonFormat = "Prefer Common Format";
const wchar_t* szInputDeviceId = L"Input Device Id";

const wchar_t * szDeviceId = L"Device Id";

void ASIO2WASAPI::readFromRegistry()
{
    HKEY key = 0;
    LONG lResult = RegOpenKeyEx(HKEY_CURRENT_USER, szPrefsRegKey, 0, KEY_READ,&key);
    if (ERROR_SUCCESS == lResult)
    {
        DWORD size = sizeof (m_nChannels);
        RegGetValue(key,NULL,szChannelRegValName,RRF_RT_REG_DWORD,NULL,&m_nChannels,&size);
        size = sizeof (m_nSampleRate);
        RegGetValue(key,NULL,szSampRateRegValName,RRF_RT_REG_DWORD,NULL,&m_nSampleRate,&size);
        size = sizeof(m_nBufferSize);
        RegGetValue(key, NULL, szBufferSizeRegValName, RRF_RT_REG_DWORD, NULL, &m_nBufferSize, &size);
        size = sizeof(m_wasapiExclusiveMode);
        RegGetValue(key, NULL, szWasapiModeRegValName, RRF_RT_REG_DWORD, NULL, &m_wasapiExclusiveMode, &size);
        size = sizeof(m_wasapiEnableResampling);
        RegGetValue(key, NULL, szEnableResamplingRegValName, RRF_RT_REG_DWORD, NULL, &m_wasapiEnableResampling, &size);
        size = sizeof(m_wasapiLowLatencySharedMode);
        RegGetValue(key, NULL, szEnableLowLatencygRegValName, RRF_RT_REG_DWORD, NULL, &m_wasapiLowLatencySharedMode, &size);
        
        RegGetValueW(key,NULL,szDeviceId,RRF_RT_REG_SZ,NULL,NULL,&size);
        m_deviceId.resize(size/sizeof(m_deviceId[0]));
        if (size)
            RegGetValueW(key,NULL,szDeviceId,RRF_RT_REG_SZ,NULL,&m_deviceId[0],&size);
        DWORD inputEnabled = 0;
        size = sizeof(inputEnabled);
        RegGetValue(key, NULL, szInputEnabled, RRF_RT_REG_DWORD, NULL, &inputEnabled, &size);
        m_inputEnabled = inputEnabled != 0;
        size = sizeof(m_preferCommonFormat);
        RegGetValue(key, NULL, szPreferCommonFormat, RRF_RT_REG_DWORD, NULL, &m_preferCommonFormat, &size);
        DWORD inputSize = 0;
        if (RegGetValueW(key, NULL, szInputDeviceId, RRF_RT_REG_SZ, NULL, NULL, &inputSize) == ERROR_SUCCESS && inputSize)
        {
            m_inputDeviceId.resize(inputSize / sizeof(wchar_t));
            if (RegGetValueW(key, NULL, szInputDeviceId, RRF_RT_REG_SZ, NULL, &m_inputDeviceId[0], &inputSize) != ERROR_SUCCESS)
                m_inputDeviceId.clear();
        }
        // The empty string denotes the default capture endpoint.
        if (!m_inputDeviceId.empty() && m_inputDeviceId[0] == 0) m_inputDeviceId.clear();
        RegCloseKey(key);
        size = (DWORD)(m_deviceId.size() * sizeof(wchar_t));
        m_useDefaultDevice = (size < 16);

    }
}

void ASIO2WASAPI::writeToRegistry()
{
    HKEY key = 0;
    LONG lResult = RegCreateKeyEx(HKEY_CURRENT_USER, szPrefsRegKey, 0, NULL,0, KEY_WRITE,NULL,&key,NULL);
    if (ERROR_SUCCESS == lResult)
    {
        DWORD size = sizeof (m_nChannels);
        RegSetValueEx (key,szChannelRegValName,NULL,REG_DWORD,(const BYTE *)&m_nChannels,size);
        size = sizeof (m_nSampleRate);
        RegSetValueEx(key,szSampRateRegValName,NULL,REG_DWORD,(const BYTE *)&m_nSampleRate,size);
        size = sizeof(m_nBufferSize);
        RegSetValueEx(key, szBufferSizeRegValName, NULL, REG_DWORD, (const BYTE*)&m_nBufferSize, size);
        size = sizeof(m_wasapiExclusiveMode);
        RegSetValueEx(key, szWasapiModeRegValName, NULL, REG_DWORD, (const BYTE*)&m_wasapiExclusiveMode, size);
        size = sizeof(m_wasapiEnableResampling);
        RegSetValueEx(key, szEnableResamplingRegValName, NULL, REG_DWORD, (const BYTE*)&m_wasapiEnableResampling, size);
        size = sizeof(m_wasapiLowLatencySharedMode);
        RegSetValueEx(key, szEnableLowLatencygRegValName, NULL, REG_DWORD, (const BYTE*)&m_wasapiLowLatencySharedMode, size);
        
        size = (DWORD)(m_deviceId.size()) * sizeof(m_deviceId[0]);        
        RegSetValueExW(key,szDeviceId,NULL,REG_SZ,(const BYTE *)(m_deviceId.empty() ? L"" : &m_deviceId[0]),size ? size : sizeof(wchar_t));
        RegSetValueEx(key, szPreferCommonFormat, 0, REG_DWORD, (const BYTE*)&m_preferCommonFormat, sizeof(m_preferCommonFormat));
        DWORD inputEnabled = m_inputEnabled ? 1 : 0;
        RegSetValueEx(key, szInputEnabled, 0, REG_DWORD, (const BYTE*)&inputEnabled, sizeof(inputEnabled));
        const wchar_t* inputId = m_inputDeviceId.empty() ? L"" : &m_inputDeviceId[0];
        RegSetValueExW(key, szInputDeviceId, 0, REG_SZ, (const BYTE*)inputId, (DWORD)((wcslen(inputId) + 1) * sizeof(wchar_t)));
        RegCloseKey(key);
    }
}

void ASIO2WASAPI::clearState()
{
    //fields valid before initialization
    m_nChannels = 2;
    m_nSampleRate = 48000;    
    m_nBufferSize = m_wasapiExclusiveMode ? 10 : 30;

    memset(m_errorMessage,0,sizeof(m_errorMessage));
    m_deviceId.clear();
    m_hStopPlayThreadEvent = NULL; 

    //fields filled by init()/cleaned by shutdown()
    m_active = false;
    m_pDevice = NULL;
    m_pAudioClient = NULL;
    m_pInputDevice = NULL;
    m_pCaptureAudioClient = NULL;
    m_nInputChannels = 0;
    m_hasInputBuffers = m_hasOutputBuffers = false;
    m_capturePrimed = false;
    memset(&m_inputWaveFormat, 0, sizeof(m_inputWaveFormat));
    memset(&m_waveFormat,0,sizeof(m_waveFormat));
    m_bufferIndex = 0;
    m_hAppWindowHandle = NULL;   

    //fields filled by createBuffers()/cleaned by disposeBuffers()
    m_buffers[0].clear();
    m_buffers[1].clear();
    m_inputBuffers[0].clear();
    m_inputBuffers[1].clear();
    m_captureBuffer.clear();
    m_callbacks = NULL;

    //fields filled by start()/cleaned by stop()
    m_hPlayThread = NULL;
    m_playThreadId = 0;
    m_bufferSize = 0;
    m_theSystemTime.hi = 0;
    m_theSystemTime.lo = 0;
    m_samplePosition = 0;
}

ASIO2WASAPI::ASIO2WASAPI (LPUNKNOWN pUnk, HRESULT *phr)
	: CUnknown((TCHAR*)"ASIO2WASAPI", pUnk, phr)
{
    clearState();
    readFromRegistry();   
}

ASIO2WASAPI::~ASIO2WASAPI ()
{
    shutdown();    
}

void ASIO2WASAPI::shutdown()
{   
    IMMDeviceEnumerator* pEnumerator = NULL;    
    HRESULT hr = S_OK;
    
    //stop(); Redundant since disposeuffers calls stop()
	if (disposeBuffers() != ASE_OK) return;
    
    if (!m_hCallbackEvent)
    {
        HANDLE hEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
        CHandleCloser cl(hEvent);
        if (m_pAudioClient) m_pAudioClient->SetEventHandle(hEvent); //Without this you have to wait long seconds in Win7...
    }
    SAFE_RELEASE(m_pAudioClient)
    
    SAFE_RELEASE(m_pDevice)
    SAFE_RELEASE(m_pInputDevice)
    clearState();
    if (m_hCallbackEvent)
    {
        CloseHandle(m_hCallbackEvent);
        m_hCallbackEvent = NULL;
    }
    
    hr = CoCreateInstance(
        CLSID_MMDeviceEnumerator, NULL,
        CLSCTX_ALL, IID_IMMDeviceEnumerator,
        (void**)&pEnumerator);
    if (FAILED(hr)) return;
    CReleaser r2(pEnumerator);
    
    if (pNotificationClient)
    {
        pEnumerator->UnregisterEndpointNotificationCallback(pNotificationClient);
        pNotificationClient->Release();
        pNotificationClient = NULL;
    }

}

void ASIO2WASAPI::initInputFields(IMMDevice* pDevice, ASIO2WASAPI* pDriver, const HWND hwndDlg)
{
    static const wchar_t* const sampleRates[6] = { L"22050", L"32000", L"44100", L"48000", L"96000", L"192000" };
    static int sampleRatesLength = sizeof(sampleRates) / sizeof(sampleRates[0]);
    
    DWORD devMixSampleRate = 48000;
    WORD devMixChannels = 2;
    HRESULT hr;

    IAudioClient* pAudioClient = NULL;
    hr = pDevice->Activate(
        IID_IAudioClient, CLSCTX_ALL,
        NULL, (void**)&pAudioClient);
    CReleaser r(pAudioClient);

    if (SUCCEEDED(hr) && pAudioClient)
    {
        WAVEFORMATEX* devFormat;
        hr = pAudioClient->GetMixFormat(&devFormat);
        if (SUCCEEDED(hr))
        {
            devMixSampleRate = devFormat->nSamplesPerSec;
            devMixChannels = devFormat->nChannels;
            CoTaskMemFree(devFormat);
        }
    }

    //channels
    wchar_t tmpBuff[8] = { 0 };
    for (UINT i = 2; i <= 20; i += 2)
    {
        if (IsFormatSupported(pDevice, i, devMixSampleRate, pDriver->m_wasapiExclusiveMode, pDriver->m_wasapiEnableResampling))
        {
            SendDlgItemMessageW(hwndDlg, IDC_CHANNELS, CB_ADDSTRING, 0, (LPARAM)_itow(i, tmpBuff, 10));
        }

    }
    LRESULT nItemIdIndex = SendDlgItemMessageW(hwndDlg, IDC_CHANNELS, CB_FINDSTRING, -1, (LPARAM)_itow(pDriver->m_nChannels, tmpBuff, 10));
    if (nItemIdIndex > 0)
        SendDlgItemMessage(hwndDlg, IDC_CHANNELS, CB_SETCURSEL, nItemIdIndex, 0);
    else
        SendDlgItemMessage(hwndDlg, IDC_CHANNELS, CB_SETCURSEL, 0, 0);

    //sample rates
    for (int i = 0; i < sampleRatesLength; i++)
    {
        if (IsFormatSupported(pDevice, pDriver->m_wasapiExclusiveMode ? 2 : devMixChannels, _wtoi(sampleRates[i]), pDriver->m_wasapiExclusiveMode, pDriver->m_wasapiEnableResampling))
        {
            SendDlgItemMessageW(hwndDlg, IDC_SAMPLE_RATE, CB_ADDSTRING, 0, (LPARAM)sampleRates[i]);
        }

    }
    nItemIdIndex = SendDlgItemMessageW(hwndDlg, IDC_SAMPLE_RATE, CB_FINDSTRING, -1, (LPARAM)_itow(pDriver->m_nSampleRate, tmpBuff, 10));
    if (nItemIdIndex > 0)
        SendDlgItemMessage(hwndDlg, IDC_SAMPLE_RATE, CB_SETCURSEL, nItemIdIndex, 0);
    else
        SendDlgItemMessage(hwndDlg, IDC_SAMPLE_RATE, CB_SETCURSEL, 0, 0);

}

static bool SupportsLowLatencySharedMode(IMMDevice* device)
{
    if (!device) return false;
    DWORD state = 0;
    if (FAILED(device->GetState(&state)) || !(state & DEVICE_STATE_ACTIVE)) return false;

    IAudioClient* client = NULL;
    HRESULT hr = device->Activate(IID_IAudioClient, CLSCTX_ALL, NULL, (void**)&client);
    CReleaser releaseClient(client);
    if (FAILED(hr) || !client) return false;
    IAudioClient3* client3 = NULL;
    hr = client->QueryInterface(&client3);
    CReleaser releaseClient3(client3);
    if (FAILED(hr) || !client3) return false;

    UINT32 currentPeriod = 0, minPeriod = 0, maxPeriod = 0;
    UINT32 fundamentalPeriod = 0, defaultPeriod = 0;
    WAVEFORMATEX* format = NULL;
    hr = client3->GetCurrentSharedModeEnginePeriod(&format, &currentPeriod);
    if (SUCCEEDED(hr) && format)
        hr = client3->GetSharedModeEnginePeriod(format, &defaultPeriod,
            &fundamentalPeriod, &minPeriod, &maxPeriod);
    else if (SUCCEEDED(hr))
        hr = E_FAIL;
    CoTaskMemFree(format);
    // Preserve the playback-only criterion: a shorter-than-default period
    // must be available. Merely exposing IAudioClient3 is not enough.
    return SUCCEEDED(hr) && fundamentalPeriod && minPeriod &&
        minPeriod < defaultPeriod && maxPeriod >= defaultPeriod;
}

static bool SelectedEndpointSupportsLowLatency(IMMDeviceEnumerator* enumerator,
    HWND dialog, int control, const vector< vector<wchar_t> >& ids,
    EDataFlow flow, LRESULT defaultIndex)
{
    LRESULT index = SendDlgItemMessage(dialog, control, CB_GETCURSEL, 0, 0);
    if (index == CB_ERR || index < 0 || (size_t)index >= ids.size()) return false;
    IMMDevice* device = NULL;
    HRESULT hr = index == defaultIndex ?
        enumerator->GetDefaultAudioEndpoint(flow, eConsole, &device) :
        (ids[index].empty() ? E_FAIL : enumerator->GetDevice(&ids[index][0], &device));
    CReleaser releaseDevice(device);
    return SUCCEEDED(hr) && SupportsLowLatencySharedMode(device);
}

void ASIO2WASAPI::updateLowLatencyAvailability(ASIO2WASAPI* driver, HWND dialog,
    const vector< vector<wchar_t> >& deviceIds, const vector< vector<wchar_t> >& inputDeviceIds)
{
    bool available = false;
    if (!driver->m_wasapiExclusiveMode)
    {
        IMMDeviceEnumerator* enumerator = NULL;
        HRESULT hr = CoCreateInstance(CLSID_MMDeviceEnumerator, NULL, CLSCTX_ALL,
            IID_IMMDeviceEnumerator, (void**)&enumerator);
        CReleaser releaseEnumerator(enumerator);
        if (SUCCEEDED(hr) && enumerator)
        {
            available = SelectedEndpointSupportsLowLatency(enumerator, dialog,
                IDC_DEVICE, deviceIds, eRender, 0);
            // Consult the current dialog selection, which has not yet been
            // saved into m_inputEnabled/m_inputDeviceId by the OK handler.
            if (available && SendDlgItemMessage(dialog, IDC_INPUT_DEVICE, CB_GETCURSEL, 0, 0) != 0)
                available = SelectedEndpointSupportsLowLatency(enumerator, dialog,
                    IDC_INPUT_DEVICE, inputDeviceIds, eCapture, 1);
        }
    }
    EnableWindow(GetDlgItem(dialog, IDC_LOWLATENCY), available);
    if (!available)
    {
        SendDlgItemMessage(dialog, IDC_LOWLATENCY, BM_SETCHECK, BST_UNCHECKED, 0);
        driver->m_wasapiLowLatencySharedMode = FALSE;
    }
}

static void FormatStreamDiagnostic(char* text, size_t size, const char* name,
    const WAVEFORMATEXTENSIBLE& format)
{
    const WAVEFORMATEX& wave = format.Format;
    bool floating = format.SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT ||
        wave.wFormatTag == WAVE_FORMAT_IEEE_FLOAT;
    if (floating)
        sprintf_s(text, size, "%s: IEEE float %u-bit\r\n%lu Hz, %u channels", name,
            (unsigned)wave.wBitsPerSample, (unsigned long)wave.nSamplesPerSec, (unsigned)wave.nChannels);
    else
    {
        unsigned valid = wave.wFormatTag == WAVE_FORMAT_EXTENSIBLE && format.Samples.wValidBitsPerSample ?
            format.Samples.wValidBitsPerSample : wave.wBitsPerSample;
        sprintf_s(text, size, "%s: PCM %u-bit / %u valid\r\n%lu Hz, %u channels", name,
            (unsigned)wave.wBitsPerSample, valid, (unsigned long)wave.nSamplesPerSec, (unsigned)wave.nChannels);
    }
}

void ASIO2WASAPI::updateFormatDiagnostics(HWND dialog) const
{
    char text[160];
    if (m_active) FormatStreamDiagnostic(text, sizeof(text), "Playback", m_waveFormat);
    else strcpy_s(text, "Playback: Not initialized");
    SetDlgItemTextA(dialog, IDC_OUTPUT_FORMAT_INFO, text);

    if (!m_inputEnabled) strcpy_s(text, "Recording: Disabled");
    else if (m_active && m_nInputChannels > 0)
        FormatStreamDiagnostic(text, sizeof(text), "Recording", m_inputWaveFormat);
    else strcpy_s(text, "Recording: Not initialized");
    SetDlgItemTextA(dialog, IDC_INPUT_FORMAT_INFO, text);

    // This reports whether the negotiated stream representation requires our
    // capture conversion, even if the host pauses audio to open this dialog.
    // Valid-bit counts can differ in identical PCM containers without conversion.
    if (!m_inputEnabled) strcpy_s(text, "Capture conversion: Not applicable");
    else if (!m_active || m_nInputChannels <= 0) strcpy_s(text, "Capture conversion: Unavailable");
    else
    {
        bool inputFloat = m_inputWaveFormat.SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT;
        bool outputFloat = m_waveFormat.SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT;
        if (inputFloat == outputFloat && m_inputWaveFormat.Format.wBitsPerSample == m_waveFormat.Format.wBitsPerSample)
            strcpy_s(text, "Capture conversion: None (direct copy)");
        else
            sprintf_s(text, sizeof(text), "Capture conversion: %s %u -> %s %u", inputFloat ? "Float" : "PCM",
                (unsigned)m_inputWaveFormat.Format.wBitsPerSample, outputFloat ? "Float" : "PCM",
                (unsigned)m_waveFormat.Format.wBitsPerSample);
    }
    SetDlgItemTextA(dialog, IDC_CAPTURE_CONVERSION_INFO, text);
}

BOOL CALLBACK ASIO2WASAPI::ControlPanelProc(HWND hwndDlg, 
        UINT message, WPARAM wParam, LPARAM lParam)
{ 
    static ASIO2WASAPI * pDriver = NULL;
    static vector< vector<wchar_t> > deviceStringIds;
    static vector< vector<wchar_t> > inputDeviceStringIds;
    
    switch (message) 
    { 
         case WM_DESTROY:
            if (pDriver) pDriver->m_hControlPanelHandle = 0;
            pDriver = NULL;
            deviceStringIds.clear();
            inputDeviceStringIds.clear();
            return 0;
         case WM_COMMAND: 
            {
            switch (LOWORD(wParam)) 
            { 

            case IDC_LOWLATENCY:
            {
                if (HIWORD(wParam) == BN_CLICKED)
                {
                    LRESULT lr = SendDlgItemMessage(hwndDlg, IDC_LOWLATENCY, BM_GETCHECK, 0, 0);
                    pDriver->m_wasapiLowLatencySharedMode = lr == BST_CHECKED;
                    if (pDriver->m_wasapiLowLatencySharedMode)
                    {
                        pDriver->m_wasapiEnableResampling = FALSE;
                        SendDlgItemMessage(hwndDlg, IDC_RESAMPLING, BM_SETCHECK, BST_UNCHECKED, 0);
                    }
                    PostMessage(hwndDlg, WM_COMMAND, IDC_DEVICE | (CBN_SELCHANGE << 16), 0);
                }

                break;
            }
            case IDC_RESAMPLING:
            {
                if (HIWORD(wParam) == BN_CLICKED)
                {
                    LRESULT lr = SendDlgItemMessage(hwndDlg, IDC_RESAMPLING, BM_GETCHECK, 0, 0);
                    pDriver->m_wasapiEnableResampling = lr == BST_CHECKED;
                    if (pDriver->m_wasapiEnableResampling)
                    {
                        pDriver->m_wasapiLowLatencySharedMode = FALSE;
                        SendDlgItemMessage(hwndDlg, IDC_LOWLATENCY, BM_SETCHECK, BST_UNCHECKED, 0);
                    }
                    PostMessage(hwndDlg, WM_COMMAND, IDC_DEVICE | (CBN_SELCHANGE << 16), 0);
                }

                break;
            }
            case IDC_SHAREMODE:
            {
                
                if (HIWORD(wParam) == CBN_SELCHANGE) 
                {
                    LRESULT lr = SendDlgItemMessage(hwndDlg, IDC_SHAREMODE, CB_GETCURSEL, 0, 0);
                    pDriver->m_wasapiExclusiveMode = (AUDCLNT_SHAREMODE)lr;
                    EnableWindow(GetDlgItem(hwndDlg, IDC_RESAMPLING), !pDriver->m_wasapiExclusiveMode && IsWindows7OrGreater());
                    if (pDriver->m_wasapiExclusiveMode)
                    {
                        pDriver->m_wasapiEnableResampling = FALSE;
                        SendDlgItemMessage(hwndDlg, IDC_RESAMPLING, BM_SETCHECK, BST_UNCHECKED, 0);
                    }


                    PostMessage(hwndDlg, WM_COMMAND, IDC_DEVICE | (CBN_SELCHANGE << 16), 0);
                }
                break;
            }

            case IDC_INPUT_DEVICE:
                if (HIWORD(wParam) == CBN_SELCHANGE)
                {
                    EnableWindow(GetDlgItem(hwndDlg, IDC_COMMON_FORMAT),
                        SendDlgItemMessage(hwndDlg, IDC_INPUT_DEVICE, CB_GETCURSEL, 0, 0) > 0);
                    updateLowLatencyAvailability(pDriver, hwndDlg, deviceStringIds, inputDeviceStringIds);
                }
                break;

            case IDC_DEVICE:
            {
                if (HIWORD(wParam) == CBN_SELCHANGE)
                {
                    updateLowLatencyAvailability(pDriver, hwndDlg, deviceStringIds, inputDeviceStringIds);
                    SendDlgItemMessage(hwndDlg, IDC_CHANNELS, CB_RESETCONTENT, 0, 0);
                    SendDlgItemMessage(hwndDlg, IDC_SAMPLE_RATE, CB_RESETCONTENT, 0, 0);

                    //get the selected device's index from the dialog
                    LRESULT lr = SendDlgItemMessage(hwndDlg, IDC_DEVICE, CB_GETCURSEL, 0, 0);
                    pDriver->setUseDefaultDevice(!lr);

                    vector<wchar_t>& selectedDeviceId = deviceStringIds[lr];
                                        
                    //find this device
                    IMMDevice* pDevice = NULL;
                    IMMDeviceEnumerator* pEnumerator = NULL;
                    HRESULT hr = CoCreateInstance(
                        CLSID_MMDeviceEnumerator, NULL,
                        CLSCTX_ALL, IID_IMMDeviceEnumerator,
                        (void**)&pEnumerator);
                    if (FAILED(hr)) return 0;
                    CReleaser r1(pEnumerator);

                    if (!pDriver->getUseDefaultDevice())
                    {                      
                        IMMDeviceCollection* pMMDeviceCollection = NULL;
                        hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pMMDeviceCollection);
                        if (FAILED(hr)) return 0;
                        CReleaser r2(pMMDeviceCollection);

                        UINT nDevices = 0;
                        hr = pMMDeviceCollection->GetCount(&nDevices);
                        if (FAILED(hr)) return 0;

                        for (UINT i = 0; i < nDevices; i++)
                        {
                            IMMDevice* pMMDevice = NULL;
                            hr = pMMDeviceCollection->Item(i, &pMMDevice);
                            if (FAILED(hr)) continue;
                            CReleaser r(pMMDevice);
                            vector<wchar_t> deviceId = getDeviceId(pMMDevice);
                            if (deviceId.size() == 0)
                                continue;
                            if (wcscmp(&deviceId[0], &selectedDeviceId[0]) == 0)
                            {
                                pDevice = pMMDevice;
                                r.deactivate();
                                break;
                            }
                        }
                    }
                    else
                    {
                        pEnumerator->GetDefaultAudioEndpoint(
                            eRender, eConsole, &pDevice);
                    }
                    
                    if (!pDevice)
                    {
                        MessageBox(hwndDlg, "Invalid audio device", szDescription, MB_OK);
                        return 0;
                    }

                    initInputFields(pDevice, pDriver, hwndDlg);

                    
                    CReleaser r2(pDevice);

                }
                break;
            }
            case IDOK: 
                    if (pDriver)
                    {
                        int nChannels = 2;
                        int nSampleRate = 48000;
                        int nBufferSize = pDriver->m_wasapiExclusiveMode ? 10 : 30;
                        //get nChannels and nSampleRate from the dialog
                        {
                            BOOL bSuccess = FALSE;
                            int tmp = (int)GetDlgItemInt(hwndDlg,IDC_CHANNELS,&bSuccess,TRUE);
                            if (bSuccess && tmp >= 0)
                                nChannels = tmp;
                            else {
                                MessageBox(hwndDlg,"Invalid number of channels",szDescription,MB_OK);
                                return 0;                        
                            }

                            tmp = (int)GetDlgItemInt(hwndDlg, IDC_BUFFERSIZE, &bSuccess, TRUE);
                            if (bSuccess && tmp >= 0)
                                nBufferSize = tmp;
                            else {
                                MessageBox(hwndDlg, "Invalid buffer size", szDescription, MB_OK);
                                return 0;
                            }


                            tmp = (int)GetDlgItemInt(hwndDlg,IDC_SAMPLE_RATE,&bSuccess,TRUE);
                            if (bSuccess && tmp >= 0)
                                nSampleRate = tmp;
                            else {
                                MessageBox(hwndDlg,"Invalid sample rate",szDescription,MB_OK);
                                return 0;                        
                            }
                        }
                        //get the selected device's index from the dialog
                        LRESULT lr = SendDlgItemMessage(hwndDlg,IDC_DEVICE,CB_GETCURSEL,0,0);
                        if (lr == CB_ERR || lr < 0 || (size_t)lr >= deviceStringIds.size()) {
                            MessageBox(hwndDlg,"No audio device selected",szDescription,MB_OK);
                            return 0;
                        }
                        vector<wchar_t>& selectedDeviceId = deviceStringIds[lr];
                        LRESULT inputIndex = SendDlgItemMessage(hwndDlg, IDC_INPUT_DEVICE, CB_GETCURSEL, 0, 0);
                        if (inputIndex == CB_ERR || (size_t)inputIndex >= inputDeviceStringIds.size())
                        {
                            MessageBox(hwndDlg, "No recording device selected", szDescription, MB_OK);
                            return 0;
                        }
                        bool inputEnabled = inputIndex != 0;
                        BOOL preferCommon = SendDlgItemMessage(hwndDlg, IDC_COMMON_FORMAT, BM_GETCHECK, 0, 0) == BST_CHECKED;
                        vector<wchar_t> inputId = inputDeviceStringIds[inputIndex];
                        //find this device
                        IMMDevice * pDevice = NULL;
                        {
                            IMMDeviceEnumerator *pEnumerator = NULL;
                            HRESULT hr = CoCreateInstance(
                                   CLSID_MMDeviceEnumerator, NULL,
                                   CLSCTX_ALL, IID_IMMDeviceEnumerator,
                                   (void**)&pEnumerator);
                            if (FAILED(hr)) return 0;
                            CReleaser r1(pEnumerator);

                            IMMDeviceCollection *pMMDeviceCollection = NULL;
                            hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pMMDeviceCollection);
                            if (FAILED(hr)) return 0;
                            CReleaser r2(pMMDeviceCollection);
    
                            UINT nDevices = 0;
                            hr = pMMDeviceCollection->GetCount(&nDevices);
                            if (FAILED(hr)) return 0;
    
							if (!pDriver->getUseDefaultDevice())
								for (UINT i = 0; i < nDevices; i++)
								{
									IMMDevice* pMMDevice = NULL;
									hr = pMMDeviceCollection->Item(i, &pMMDevice);
									if (FAILED(hr))	continue;
									CReleaser r(pMMDevice);
									vector<wchar_t> deviceId = getDeviceId(pMMDevice);
									if (deviceId.size() == 0)
										continue;
									if (wcscmp(&deviceId[0], &selectedDeviceId[0]) == 0)
									{
										pDevice = pMMDevice;
										r.deactivate();
										break;
									}
								}
							else
							{
								pEnumerator->GetDefaultAudioEndpoint(
									eRender, eConsole, &pDevice);
							}
                        }
                        if (!pDevice)
                        {
                            MessageBox(hwndDlg,"Invalid audio device",szDescription,MB_OK);
                            return 0;
                        }
                        
                        CReleaser r2(pDevice);

                        //make sure the reset request is issued no matter how we proceed
                        class CCallbackCaller
                        {
                            ASIOCallbacks * m_pCallbacks;
                        public:
                            explicit CCallbackCaller(ASIOCallbacks * pCallbacks) : m_pCallbacks(pCallbacks) {}
                            ~CCallbackCaller() 
                            {
                                if (m_pCallbacks)
                                    m_pCallbacks->asioMessage(kAsioResetRequest,0,NULL,NULL);
                            }
                        } caller(pDriver->m_callbacks);
                        
                        //shut down the driver so no exclusive WASAPI connection would stand in our way
                        HWND hAppWindowHandle = pDriver->m_hAppWindowHandle;
                        pDriver->shutdown();
                        pDriver->updateFormatDiagnostics(hwndDlg);

                        //make sure the device supports this combination of nChannels and nSampleRate
                        int tmpBuffSize = nBufferSize;
                        BOOL rc = FindStreamFormat(pDevice, nChannels, nSampleRate, nBufferSize, pDriver->m_wasapiExclusiveMode, pDriver->m_wasapiEnableResampling, pDriver->m_wasapiLowLatencySharedMode);
                        if (!rc)
                        {
                            if(pDriver->m_wasapiExclusiveMode)
                                MessageBox(hwndDlg, "Format is not supported in WASAPI exclusive mode.",szDescription,MB_OK | MB_ICONWARNING);
                            else
                                MessageBox(hwndDlg, "Format is not supported in WASAPI shared mode.", szDescription, MB_OK | MB_ICONWARNING);
                            
                            return 0;
                        }
                        else if (tmpBuffSize != nBufferSize) 
                        {
                            char msgTxt[64] = "Closest valid buffer size seems to be ";
                            char convTxt[10] = { 0 };
                            strcat_s(msgTxt, _itoa(nBufferSize, convTxt, 10));
                            strcat_s(msgTxt, " ms.");
                            MessageBox(hwndDlg, msgTxt, szDescription, MB_OK | MB_ICONINFORMATION);
                        }
                        
                        //copy selected device/sample rate/channel combination into the driver
                        pDriver->m_nSampleRate = nSampleRate;
                        pDriver->m_nChannels = nChannels;
                        pDriver->m_nBufferSize = nBufferSize;
                        pDriver->m_deviceId.resize(selectedDeviceId.size());
                        wcscpy_s(&pDriver->m_deviceId[0],selectedDeviceId.size(),&selectedDeviceId.at(0));
                        pDriver->m_inputEnabled = inputEnabled;
                        pDriver->m_preferCommonFormat = preferCommon;
                        pDriver->m_inputDeviceId = inputId;
                        //try to init the driver
                        if (pDriver->init(hAppWindowHandle) == ASIOFalse)
                        {
                            pDriver->updateFormatDiagnostics(hwndDlg);
                            MessageBox(hwndDlg, pDriver->m_errorMessage[0] ? pDriver->m_errorMessage : "ASIO driver failed to initialize", szDescription, MB_OK);
                            return 0;
                        }
                        pDriver->writeToRegistry();
                    }
                    EndDialog(hwndDlg, wParam); 
                    return 0; 
                case IDCANCEL: 
                    EndDialog(hwndDlg, wParam); 
                    return 0; 
            }
            }
            break;
        case WM_INITDIALOG: 
            {
            pDriver = reinterpret_cast<ASIO2WASAPI*>(lParam);
            if (!pDriver)
                return FALSE;
            pDriver->m_hControlPanelHandle = hwndDlg;
            pDriver->updateFormatDiagnostics(hwndDlg);
#ifdef _WIN64           
            SetWindowText(hwndDlg, "ASIO2WASAPI x64");
#else	
            SetWindowText(hwndDlg, "ASIO2WASAPI x86");
#endif	
            wchar_t fileversionBuff[32] = { 0 }; ;
            SetWindowTextW(GetDlgItem(hwndDlg, IDC_VERSIONINFO), GetFileVersion(fileversionBuff, _countof(fileversionBuff)));

           
            
            HWND hwndOwner = 0;
            RECT rcOwner, rcDlg, rc;

            if ((hwndOwner = GetParent(hwndDlg)) == NULL)
            {
                hwndOwner = GetDesktopWindow();
            }

            GetWindowRect(hwndOwner, &rcOwner);
            GetWindowRect(hwndDlg, &rcDlg);
            CopyRect(&rc, &rcOwner);

            OffsetRect(&rcDlg, -rcDlg.left, -rcDlg.top);
            OffsetRect(&rc, -rc.left, -rc.top);
            OffsetRect(&rc, -rcDlg.right, -rcDlg.bottom);

            SetWindowPos(hwndDlg,
                HWND_TOPMOST,
                max(0, rcOwner.left + (rc.right / 2)),
                max(0, rcOwner.top + (rc.bottom / 2)),
                0, 0,
                SWP_NOSIZE);

            if (GetDlgCtrlID((HWND)wParam) != IDC_DEVICE) SetFocus(GetDlgItem(hwndDlg, IDC_DEVICE));

            //SetDlgItemInt(hwndDlg,IDC_CHANNELS,(UINT)pDriver->m_nChannels,TRUE);
            //SetDlgItemInt(hwndDlg, IDC_SAMPLE_RATE, (UINT)pDriver->m_nSampleRate, TRUE);
            SetDlgItemInt(hwndDlg, IDC_BUFFERSIZE, (UINT)pDriver->m_nBufferSize, TRUE);

            IMMDeviceEnumerator *pEnumerator = NULL;
            //DWORD flags = 0;
            (void) CoInitialize(NULL);

            HRESULT hr = CoCreateInstance(
                   CLSID_MMDeviceEnumerator, NULL,
                   CLSCTX_ALL, IID_IMMDeviceEnumerator,
                   (void**)&pEnumerator);
            if (FAILED(hr)) return false;
            CReleaser r1(pEnumerator);

            IMMDeviceCollection *pMMDeviceCollection = NULL;
            hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pMMDeviceCollection);
            if (FAILED(hr)) return false;
            CReleaser r2(pMMDeviceCollection);
    
            UINT nDevices = 0;
            hr = pMMDeviceCollection->GetCount(&nDevices);
            if (FAILED(hr)) return false;
                        
    
            vector< vector<wchar_t> > deviceIds;
            
            //Add Default Device first
            SendDlgItemMessageW(hwndDlg, IDC_DEVICE, CB_ADDSTRING, 0, (LPARAM)L"Default Device");
            vector<wchar_t> defDevId;
            defDevId.resize(2);
            wcscpy(&defDevId[0], L"0");
            deviceIds.push_back(defDevId);

            for (UINT i = 0; i < nDevices; i++) 
            {
                IMMDevice *pMMDevice = NULL;
                hr = pMMDeviceCollection->Item(i, &pMMDevice);
                if (FAILED(hr)) return false;
                CReleaser r(pMMDevice);

                vector<wchar_t> deviceId = getDeviceId(pMMDevice);
                if (deviceId.size() == 0)
                    return false;
                deviceIds.push_back(deviceId);

                IPropertyStore *pPropertyStore;
                hr = pMMDevice->OpenPropertyStore(STGM_READ, &pPropertyStore);
                if (FAILED(hr)) return false;
                CReleaser rr(pPropertyStore);
                PROPVARIANT var; PropVariantInit(&var);
                hr = pPropertyStore->GetValue(PKEY_Device_FriendlyName,&var);
                if (FAILED(hr)) return false;
                LRESULT lr = 0;
                if (var.vt != VT_LPWSTR ||
                    (lr = SendDlgItemMessageW(hwndDlg,IDC_DEVICE,CB_ADDSTRING,0,(LPARAM)var.pwszVal)) == CB_ERR)
                {
                    PropVariantClear(&var);
                    return false;
                }
                PropVariantClear(&var);
            }
            deviceStringIds = deviceIds;

            //find current device id
            int nItemIdIndex = 0;
            if (pDriver->m_deviceId.size())
                for (unsigned i = 1; i<deviceStringIds.size(); i++)
                {
                    if (wcscmp(&deviceStringIds[i].at(0),&pDriver->m_deviceId[0]) == 0)
                    {    
                        nItemIdIndex = i;
                        break;
                    }
                }
            SendDlgItemMessage(hwndDlg,IDC_DEVICE,CB_SETCURSEL,nItemIdIndex,0);
            pDriver->setUseDefaultDevice(!nItemIdIndex);

            // Recording endpoint IDs are independent from playback IDs.
            inputDeviceStringIds.clear();
            inputDeviceStringIds.resize(2); // Disabled and Default Device use empty IDs
            SendDlgItemMessageW(hwndDlg, IDC_INPUT_DEVICE, CB_ADDSTRING, 0, (LPARAM)L"Disabled");
            SendDlgItemMessageW(hwndDlg, IDC_INPUT_DEVICE, CB_ADDSTRING, 0, (LPARAM)L"Default Device");
            IMMDeviceCollection* captureDevices = NULL;
            if (SUCCEEDED(pEnumerator->EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE, &captureDevices)))
            {
                CReleaser releaseDevices(captureDevices);
                UINT count = 0;
                captureDevices->GetCount(&count);
                for (UINT i = 0; i < count; ++i)
                {
                    IMMDevice* device = NULL;
                    if (FAILED(captureDevices->Item(i, &device))) continue;
                    CReleaser releaseDevice(device);
                    vector<wchar_t> id = getDeviceId(device);
                    if (id.empty()) continue;
                    // A failed friendly-name lookup does not make an active
                    // endpoint unavailable. Keep its ID as a fallback label.
                    IPropertyStore* properties = NULL;
                    PROPVARIANT name;
                    PropVariantInit(&name);
                    const wchar_t* label = &id[0];
                    if (SUCCEEDED(device->OpenPropertyStore(STGM_READ, &properties)))
                    {
                        if (SUCCEEDED(properties->GetValue(PKEY_Device_FriendlyName, &name)) &&
                            name.vt == VT_LPWSTR && name.pwszVal)
                            label = name.pwszVal;
                    }
                    LRESULT index = SendDlgItemMessageW(hwndDlg, IDC_INPUT_DEVICE, CB_ADDSTRING, 0, (LPARAM)label);
                    if (index != CB_ERR && index != CB_ERRSPACE) inputDeviceStringIds.push_back(id);
                    PropVariantClear(&name);
                    SAFE_RELEASE(properties)
                }
            }
            int inputIndex = pDriver->m_inputEnabled ? 1 : 0;
            if (pDriver->m_inputEnabled && !pDriver->m_inputDeviceId.empty())
            {
                bool found = false;
                for (size_t i = 2; i < inputDeviceStringIds.size(); ++i)
                    if (!inputDeviceStringIds[i].empty() &&
                        wcscmp(&inputDeviceStringIds[i][0], &pDriver->m_inputDeviceId[0]) == 0) { inputIndex = (int)i; found = true; break; }
                if (!found)
                {
                    // Resolve the selected endpoint directly before calling it
                    // unavailable; enumeration/property retrieval can be partial.
                    IMMDevice* selected = NULL;
                    DWORD state = 0;
                    bool available = SUCCEEDED(pEnumerator->GetDevice(&pDriver->m_inputDeviceId[0], &selected)) &&
                        SUCCEEDED(selected->GetState(&state)) && (state & DEVICE_STATE_ACTIVE);
                    IPropertyStore* properties = NULL;
                    PROPVARIANT name;
                    PropVariantInit(&name);
                    const wchar_t* label = available ? L"Selected recording device" : L"Selected device (unavailable)";
                    if (available && SUCCEEDED(selected->OpenPropertyStore(STGM_READ, &properties)) &&
                        SUCCEEDED(properties->GetValue(PKEY_Device_FriendlyName, &name)) &&
                        name.vt == VT_LPWSTR && name.pwszVal)
                        label = name.pwszVal;
                    SendDlgItemMessageW(hwndDlg, IDC_INPUT_DEVICE, CB_ADDSTRING, 0, (LPARAM)label);
                    PropVariantClear(&name);
                    SAFE_RELEASE(properties)
                    SAFE_RELEASE(selected)
                    inputDeviceStringIds.push_back(pDriver->m_inputDeviceId);
                    inputIndex = (int)inputDeviceStringIds.size() - 1;
                }
            }
            SendDlgItemMessage(hwndDlg, IDC_INPUT_DEVICE, CB_SETCURSEL, inputIndex, 0);
            SendDlgItemMessage(hwndDlg, IDC_COMMON_FORMAT, BM_SETCHECK,
                pDriver->m_preferCommonFormat ? BST_CHECKED : BST_UNCHECKED, 0);
            EnableWindow(GetDlgItem(hwndDlg, IDC_COMMON_FORMAT), inputIndex != 0);

            //Share mode
            SendDlgItemMessageW(hwndDlg, IDC_SHAREMODE, CB_ADDSTRING, 0, (LPARAM)L"Shared");
            SendDlgItemMessageW(hwndDlg, IDC_SHAREMODE, CB_ADDSTRING, 0, (LPARAM)L"Exclusive");
            if (pDriver->m_wasapiExclusiveMode)
            {
                SendDlgItemMessage(hwndDlg, IDC_SHAREMODE, CB_SETCURSEL, 1, 0);
            }
            else 
            {
                SendDlgItemMessage(hwndDlg, IDC_SHAREMODE, CB_SETCURSEL, 0, 0);
                if (IsWindows7OrGreater())
                {
                    EnableWindow(GetDlgItem(hwndDlg, IDC_RESAMPLING), true);
                    if (pDriver->m_wasapiEnableResampling)
                        SendDlgItemMessage(hwndDlg, IDC_RESAMPLING, BM_SETCHECK, BST_CHECKED, 0);
                    else if (pDriver->m_wasapiLowLatencySharedMode)
                        SendDlgItemMessage(hwndDlg, IDC_LOWLATENCY, BM_SETCHECK, BST_CHECKED, 0);
                }
            }

            initInputFields(pDriver->m_pDevice, pDriver, hwndDlg);
            updateLowLatencyAvailability(pDriver, hwndDlg, deviceStringIds, inputDeviceStringIds);


            SendDlgItemMessage(hwndDlg, IDC_BUFFERSIZE, EM_LIMITTEXT, 4, 0);

            return TRUE;
            }
            break;
    } 
    return FALSE; 
} 

#define RETURN_ON_ERROR(hres)  \
              if (FAILED(hres)) { if (WaitForSingleObject(pDriver->m_hThreadStartedEvent, 0) != WAIT_OBJECT_0) pDriver->m_threadStartResult = (hres); return; }

void ASIO2WASAPI::PlayThreadProcShared(LPVOID pThis)
{
    ASIO2WASAPI* pDriver = static_cast<ASIO2WASAPI*>(pThis);


    HRESULT hr = S_OK;

    IAudioClient* pAudioClient = pDriver->m_pAudioClient;
    CAudioStopper stopAudio(pAudioClient);
    IAudioRenderClient* pRenderClient = NULL;
    BYTE* pData = NULL;

    if (!pDriver->m_hCallbackEvent)// In Cubase 5 multiple start/stop cycles can occur without releasing AudioClient. And in shared mode AudioClient->SetEventHandle fails the 2nd time. So private m_hCallbackEvent added as a global event.
    {
        pDriver->m_hCallbackEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
        hr = pAudioClient->SetEventHandle(pDriver->m_hCallbackEvent);
        RETURN_ON_ERROR(hr)
    }
    else
        ResetEvent(pDriver->m_hCallbackEvent); //make sure event is not signaled when AudioClient start is called

    hr = pAudioClient->GetService(
            IID_IAudioRenderClient,
            (void**)&pRenderClient);

    RETURN_ON_ERROR(hr)
        CReleaser r(pRenderClient);

    // Ask MMCSS to temporarily boost the thread priority
    // to reduce the possibility of glitches while we play.
    DWORD taskIndex = 0;
    HANDLE hAv = AvSetMmThreadCharacteristics(TEXT("Pro Audio"), &taskIndex);
    
    CAvReverter avReverter(hAv);
    BOOL result = FALSE;
    if (hAv)
        result = AvSetMmThreadPriority(hAv, AVRT_PRIORITY_CRITICAL);

    if (!result)
        SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL); //fail save high priority.

    // Pre-load the first buffer with data  
   
    UINT32 bufferFrameCount;
    UINT32 numFramesPadding;
    hr = pAudioClient->GetBufferSize(&bufferFrameCount);
    RETURN_ON_ERROR(hr)
    hr = pAudioClient->GetCurrentPadding(&numFramesPadding);
    RETURN_ON_ERROR(hr)
    
    UINT32 startFrames;
    if (bufferFrameCount > (UINT32)(4 * pDriver->m_bufferSize))
        startFrames = bufferFrameCount - pDriver->m_bufferSize;
    else
        startFrames = pDriver->m_bufferSize;

    if (startFrames > (bufferFrameCount - numFramesPadding)) startFrames = bufferFrameCount - numFramesPadding;

    hr = pRenderClient->GetBuffer(startFrames, &pData);
    RETURN_ON_ERROR(hr)
    //memset(pData, 0, bufferFrameCount * pDriver->m_waveFormat.Format.nBlockAlign);
    hr = pRenderClient->ReleaseBuffer(startFrames, AUDCLNT_BUFFERFLAGS_SILENT);
    RETURN_ON_ERROR(hr)

    hr = pAudioClient->Start();  // Start playing.
    RETURN_ON_ERROR(hr)

    pDriver->m_threadStartResult = S_OK;
    SetEvent(pDriver->m_hThreadStartedEvent);
    getNanoSeconds(&pDriver->m_theSystemTime);
    pDriver->m_samplePosition = 0;

    if (pDriver->m_callbacks)
        pDriver->m_callbacks->bufferSwitch(1 - pDriver->m_bufferIndex, ASIOTrue);

    //char convTxt[11] = { 0 };

    DWORD retval = 0;
    HANDLE events[2] = { pDriver->m_hStopPlayThreadEvent, pDriver->m_hCallbackEvent };
    while ((retval = WaitForMultipleObjects(2, events, FALSE, INFINITE)) == (WAIT_OBJECT_0 + 1))
    {//the hEvent is signalled and m_hStopPlayThreadEvent is not
        // Grab the next empty buffer from the audio device.   
       
        hr = pAudioClient->GetCurrentPadding(&numFramesPadding);
        if (FAILED(hr))
        {
            if (pDriver->m_callbacks->asioMessage) pDriver->m_callbacks->asioMessage(kAsioResetRequest, 0, NULL, NULL);
            break;
        }
        if (pDriver->m_bufferSize > (int)(bufferFrameCount - numFramesPadding))
        {           
            continue; //we will get more space in next turn
        }

        hr = pDriver->LoadData(pRenderClient);
        if (hr == AUDCLNT_E_BUFFER_ERROR) continue;
        if (hr != S_OK)
        {
            //OutputDebugString(_itoa(hr, convTxt, 10));
            if (pDriver->m_callbacks->asioMessage) pDriver->m_callbacks->asioMessage(kAsioResetRequest, 0, NULL, NULL);
            break;
        }
        getNanoSeconds(&pDriver->m_theSystemTime);
        pDriver->m_samplePosition += pDriver->m_bufferSize;
        if (pDriver->m_callbacks)
            pDriver->m_callbacks->bufferSwitch(1 - pDriver->m_bufferIndex, ASIOTrue);
       
    }

    hr = pAudioClient->Stop();  // Stop playing.
    RETURN_ON_ERROR(hr)
        pDriver->m_samplePosition = 0;

    return;
}


void ASIO2WASAPI::PlayThreadProc(LPVOID pThis)
{
    ASIO2WASAPI * pDriver = static_cast<ASIO2WASAPI *>(pThis);


    HRESULT hr=S_OK;

    IAudioClient *pAudioClient = pDriver->m_pAudioClient;
    CAudioStopper stopAudio(pAudioClient);
    IAudioRenderClient *pRenderClient = NULL;
    BYTE *pData = NULL;
                                              


    if (!pDriver->m_hCallbackEvent) // In Cubase 5 multiple start/stop cycles can occur without releasing AudioClient. And in shared mode AudioClient->SetEventHandle fails the 2nd time. So private m_hCallbackEvent added as a global event.
    {
        pDriver->m_hCallbackEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
        hr = pAudioClient->SetEventHandle(pDriver->m_hCallbackEvent);
        RETURN_ON_ERROR(hr)
    }
    else
        ResetEvent(pDriver->m_hCallbackEvent); //make sure event is not signaled when AudioClient start is called 

    hr = pAudioClient->GetService(
                         IID_IAudioRenderClient,
                         (void**)&pRenderClient);

    RETURN_ON_ERROR(hr)
    CReleaser r(pRenderClient);
    
    // Ask MMCSS to temporarily boost the thread priority
    // to reduce the possibility of glitches while we play.
    DWORD taskIndex = 0;
    HANDLE hAv = AvSetMmThreadCharacteristics(TEXT("Pro Audio"), &taskIndex);
    
    CAvReverter avReverter(hAv);
    BOOL result = FALSE;
    if (hAv)
        result = AvSetMmThreadPriority(hAv, AVRT_PRIORITY_CRITICAL);

    if (!result)
        SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL); //fail save high priority.

    // Pre-load the first buffer with data
    // from the audio source before starting the stream.
    hr = pDriver->LoadData(pRenderClient);  //Does this actually add latency of the length of 1 buffer?     
    RETURN_ON_ERROR(hr)

    hr = pAudioClient->Start();  // Start playing.
    RETURN_ON_ERROR(hr)

    pDriver->m_threadStartResult = S_OK;
    SetEvent(pDriver->m_hThreadStartedEvent);
    getNanoSeconds(&pDriver->m_theSystemTime);
    pDriver->m_samplePosition = 0;    
    
    if (pDriver->m_callbacks)
        pDriver->m_callbacks->bufferSwitch(1-pDriver->m_bufferIndex,ASIOTrue);

    DWORD normalInterval = ((DWORD)round(pDriver->m_bufferSize / (pDriver->m_nSampleRate * 0.001))) + 1;
    DWORD counter = 0;
    LARGE_INTEGER startTime = { 0 };    
    double	queryPerformanceUnit = 0.0;

    LARGE_INTEGER tmpFreq;
    if (QueryPerformanceFrequency(&tmpFreq))
    {
        queryPerformanceUnit = 1.0 / (tmpFreq.QuadPart * 0.001);
        QueryPerformanceCounter(&startTime);
    }
    //char convTxt[11] = { 0 };
    
    DWORD retval = 0;
    HANDLE events[2] = {pDriver->m_hStopPlayThreadEvent, pDriver->m_hCallbackEvent };
    while ((retval  = WaitForMultipleObjects(2,events,FALSE, INFINITE)) == (WAIT_OBJECT_0 + 1))
    {//the hEvent is signalled and m_hStopPlayThreadEvent is not
        // Grab the next empty buffer from the audio device.
       
       //workaround for driver bug when suddenly event interval increases with a fixed amount.
       //timeGetTime is not reliable on Win11 so we use QPC functions. 
        LARGE_INTEGER tmpCounter = { 0 };
        QueryPerformanceCounter(&tmpCounter);
        DWORD endTime = (DWORD)round((tmpCounter.QuadPart - startTime.QuadPart) * queryPerformanceUnit);
        //OutputDebugString(_itoa(endTime, convTxt, 10));

        if (endTime > normalInterval)
            counter++;
        else
            counter = 0;

        if (counter > 10)
        {
            //pDriver->m_callbacks->asioMessage(kAsioResetRequest, 0, NULL, NULL);
            //break;
            pAudioClient->Stop();
            pAudioClient->Reset();
            pAudioClient->Start();
            counter = 0;
        }

        hr = pDriver->LoadData(pRenderClient);
        if (hr == AUDCLNT_E_BUFFER_ERROR) continue;
        if (hr != S_OK)
        {
            if (pDriver->m_callbacks->asioMessage) pDriver->m_callbacks->asioMessage(kAsioResetRequest, 0, NULL, NULL);
            break;
        }            
        getNanoSeconds(&pDriver->m_theSystemTime);
        pDriver->m_samplePosition += pDriver->m_bufferSize;
        if (pDriver->m_callbacks)
            pDriver->m_callbacks->bufferSwitch(1-pDriver->m_bufferIndex,ASIOTrue);

        QueryPerformanceCounter(&startTime);
    }

    hr = pAudioClient->Stop();  // Stop playing.
    RETURN_ON_ERROR(hr)
    pDriver->m_samplePosition = 0;    

    return;
}

HRESULT ASIO2WASAPI::DrainCapture(IAudioCaptureClient* capture)
{
    for (;;)
    {
        if (WaitForSingleObject(m_hStopPlayThreadEvent, 0) == WAIT_OBJECT_0) return S_OK;
        // GetNextPacketSize is only supported for shared-mode streams.
        if (!m_wasapiExclusiveMode)
        {
            UINT32 nextFrames = 0;
            HRESULT hr = capture->GetNextPacketSize(&nextFrames);
            if (FAILED(hr)) return hr;
            if (!nextFrames) return S_OK;
        }
        BYTE* data = NULL;
        UINT32 frames = 0;
        DWORD flags = 0;
        HRESULT hr = capture->GetBuffer(&data, &frames, &flags, NULL, NULL);
        if (hr == AUDCLNT_S_BUFFER_EMPTY || hr == AUDCLNT_E_BUFFER_ERROR) return S_OK;
        if (FAILED(hr)) return hr;
        if (flags & AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY)
        {
            m_captureBuffer.clear();
            m_capturePrimed = false;
        }
        m_captureBuffer.push(data, frames, (flags & AUDCLNT_BUFFERFLAGS_SILENT) != 0);
        hr = capture->ReleaseBuffer(frames);
        if (FAILED(hr)) return hr;
        // Exclusive event-driven capture supplies one whole endpoint buffer.
        if (m_wasapiExclusiveMode) return S_OK;
    }
}

void ASIO2WASAPI::SwitchInputBuffer(int index)
{
    // Leave one ASIO block of headroom so different packet boundaries do not
    // cause periodic underruns even when both endpoints share the same clock.
    if (m_hasOutputBuffers && !m_capturePrimed)
    {
        if (m_captureBuffer.available() < (size_t)(2 * m_bufferSize))
        {
            for (size_t channel = 0; channel < m_inputBuffers[index].size(); ++channel)
                if (!m_inputBuffers[index][channel].empty())
                    memset(&m_inputBuffers[index][channel][0], 0, m_inputBuffers[index][channel].size());
            return;
        }
        m_capturePrimed = true;
    }
    if (m_captureBuffer.available() < (size_t)m_bufferSize) m_capturePrimed = false;
    m_captureBuffer.readPlanarConverted(m_inputBuffers[index], m_bufferSize,
        m_inputWaveFormat.Format.wBitsPerSample / 8,
        m_inputWaveFormat.SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT,
        m_waveFormat.Format.wBitsPerSample / 8,
        m_waveFormat.SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT);
}

void ASIO2WASAPI::DuplexThreadProc(LPVOID pThis)
{
    ASIO2WASAPI* driver = static_cast<ASIO2WASAPI*>(pThis);
    ASIO2WASAPI* pDriver = driver; // shared startup-error macro
    IAudioClient* renderAudio = driver->m_hasOutputBuffers ? driver->m_pAudioClient : NULL;
    IAudioClient* captureAudio = driver->m_pCaptureAudioClient;
    HRESULT hr = S_OK;
    // Cleanup must stop/reset both streams even after a partial startup failure.
    CAudioStopper stopRender(renderAudio), stopCapture(captureAudio);
    IAudioCaptureClient* capture = NULL;
    hr = captureAudio->GetService(__uuidof(IAudioCaptureClient), (void**)&capture);
    RETURN_ON_ERROR(hr)
    CReleaser releaseCapture(capture);
    IAudioRenderClient* render = NULL;
    CReleaser releaseRender(NULL);
    UINT32 renderFrames = 0;
    if (renderAudio)
    {
        if (!driver->m_hCallbackEvent)
        {
            driver->m_hCallbackEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
            hr = driver->m_hCallbackEvent ? renderAudio->SetEventHandle(driver->m_hCallbackEvent) : E_OUTOFMEMORY;
            RETURN_ON_ERROR(hr)
        }
        ResetEvent(driver->m_hCallbackEvent);
        hr = renderAudio->GetService(IID_IAudioRenderClient, (void**)&render);
        RETURN_ON_ERROR(hr)
        releaseRender.reset(render);
        hr = renderAudio->GetBufferSize(&renderFrames);
        RETURN_ON_ERROR(hr)
        if (driver->m_wasapiExclusiveMode)
        {
            hr = driver->LoadData(render);
            RETURN_ON_ERROR(hr)
        }
        else
        {
            BYTE* data = NULL;
            UINT32 startFrames = renderFrames > (UINT32)(4 * driver->m_bufferSize) ?
                renderFrames - driver->m_bufferSize : driver->m_bufferSize;
            startFrames = min(startFrames, renderFrames);
            hr = render->GetBuffer(startFrames, &data);
            RETURN_ON_ERROR(hr)
            hr = render->ReleaseBuffer(startFrames, AUDCLNT_BUFFERFLAGS_SILENT);
            RETURN_ON_ERROR(hr)
        }
    }
    ResetEvent(driver->m_hCaptureEvent);
    DWORD taskIndex = 0;
    HANDLE av = AvSetMmThreadCharacteristics(TEXT("Pro Audio"), &taskIndex);
    CAvReverter avReverter(av);
    if (!av || !AvSetMmThreadPriority(av, AVRT_PRIORITY_CRITICAL))
        SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL);
    hr = captureAudio->Start();
    RETURN_ON_ERROR(hr)
    if (renderAudio)
    {
        hr = renderAudio->Start();
        RETURN_ON_ERROR(hr)
        driver->m_threadStartResult = S_OK;
        SetEvent(driver->m_hThreadStartedEvent);
        getNanoSeconds(&driver->m_theSystemTime);
        driver->SwitchInputBuffer(1 - driver->m_bufferIndex); // startup silence
        driver->m_callbacks->bufferSwitch(1 - driver->m_bufferIndex, ASIOTrue);
    }
    if (!renderAudio)
    {
        driver->m_threadStartResult = S_OK;
        SetEvent(driver->m_hThreadStartedEvent);
    }
    LARGE_INTEGER renderStart = { 0 }, frequency = { 0 };
    QueryPerformanceFrequency(&frequency);
    QueryPerformanceCounter(&renderStart);
    double counterToMs = frequency.QuadPart ? 1000.0 / frequency.QuadPart : 0.0;
    DWORD normalInterval = (DWORD)round(driver->m_bufferSize / (driver->m_nSampleRate * 0.001)) + 1;
    DWORD slowIntervals = 0;
    HANDLE events[3] = { driver->m_hStopPlayThreadEvent, driver->m_hCaptureEvent, driver->m_hCallbackEvent };
    for (;;)
    {
        DWORD wait = WaitForMultipleObjects(renderAudio ? 3 : 2, events, FALSE, INFINITE);
        if (wait == WAIT_OBJECT_0) break;
        if (wait != WAIT_OBJECT_0 + 1 && wait != WAIT_OBJECT_0 + 2) { hr = E_FAIL; break; }
        if (wait == WAIT_OBJECT_0 + 1 || !driver->m_wasapiExclusiveMode)
        {
            hr = driver->DrainCapture(capture);
            if (FAILED(hr)) break;
        }
        if (!renderAudio)
        {
            // Input-only hosts follow the capture clock, with no render stream
            // and no synthetic callbacks before a full block has been recorded.
            while (driver->m_captureBuffer.available() >= (size_t)driver->m_bufferSize &&
                WaitForSingleObject(driver->m_hStopPlayThreadEvent, 0) != WAIT_OBJECT_0)
            {
                driver->m_bufferIndex = 1 - driver->m_bufferIndex;
                driver->SwitchInputBuffer(driver->m_bufferIndex);
                getNanoSeconds(&driver->m_theSystemTime);
                driver->m_callbacks->bufferSwitch(driver->m_bufferIndex, ASIOTrue);
                driver->m_samplePosition += driver->m_bufferSize;
            }
        }
        else if (wait == WAIT_OBJECT_0 + 2)
        {
            if (driver->m_wasapiExclusiveMode && counterToMs)
            {
                LARGE_INTEGER now;
                QueryPerformanceCounter(&now);
                DWORD interval = (DWORD)round((now.QuadPart - renderStart.QuadPart) * counterToMs);
                slowIntervals = interval > normalInterval ? slowIntervals + 1 : 0;
                if (slowIntervals > 10)
                {
                    hr = renderAudio->Stop();
                    if (SUCCEEDED(hr)) hr = renderAudio->Reset();
                    if (SUCCEEDED(hr)) hr = renderAudio->Start();
                    if (FAILED(hr)) break;
                    slowIntervals = 0;
                }
            }
            if (!driver->m_wasapiExclusiveMode)
            {
                UINT32 padding = 0;
                hr = renderAudio->GetCurrentPadding(&padding);
                if (FAILED(hr)) break;
                if (padding > renderFrames || renderFrames - padding < (UINT32)driver->m_bufferSize) continue;
            }
            hr = driver->LoadData(render);
            if (hr == AUDCLNT_E_BUFFER_ERROR) continue;
            if (FAILED(hr)) break;
            int index = 1 - driver->m_bufferIndex;
            driver->SwitchInputBuffer(index);
            getNanoSeconds(&driver->m_theSystemTime);
            driver->m_samplePosition += driver->m_bufferSize;
            driver->m_callbacks->bufferSwitch(index, ASIOTrue);
            QueryPerformanceCounter(&renderStart);
        }
    }
    if (FAILED(hr) && driver->m_callbacks->asioMessage)
        driver->m_callbacks->asioMessage(kAsioResetRequest, 0, NULL, NULL);
}

#undef RETURN_ON_ERROR

HRESULT ASIO2WASAPI::LoadData(IAudioRenderClient * pRenderClient)
{
    if (!pRenderClient)
        return E_INVALIDARG;
    
    HRESULT hr = S_OK;
    BYTE *pData = NULL;
    hr = pRenderClient->GetBuffer(m_bufferSize, &pData);
    if (hr != S_OK) return hr;

    UINT32 sampleSize=m_waveFormat.Format.wBitsPerSample/8;
    
    //switch buffer
    m_bufferIndex = 1 - m_bufferIndex;
    vector <vector <BYTE> > &buffer = m_buffers[m_bufferIndex]; 
    unsigned sampleOffset = 0;
    unsigned nextSampleOffset = sampleSize;
    for (int i = 0;i < m_bufferSize; i++, sampleOffset=nextSampleOffset, nextSampleOffset+=sampleSize)
    {
        for (unsigned j = 0; j < buffer.size(); j++) 
        {
            if (buffer[j].size() >= nextSampleOffset)
            {
                memcpy_s(pData, sampleSize, &buffer[j].at(0) + sampleOffset, sampleSize);               
            }
            else
                memset(pData,0,sampleSize);
            pData+=sampleSize;
        }
    }

    hr = pRenderClient->ReleaseBuffer(m_bufferSize, 0);

    return hr;
}

/*  ASIO driver interface implementation
*/

void ASIO2WASAPI::getDriverName (char *name)
{
	strcpy_s (name, 32, "ASIO2WASAPI");
}

long ASIO2WASAPI::getDriverVersion ()
{
	return 1;
}

void ASIO2WASAPI::getErrorMessage (char *string)
{
    strcpy_s(string,sizeof(m_errorMessage),m_errorMessage);
}

ASIOError ASIO2WASAPI::future (long selector, void* opt)	
{
    //none of the optional features are present
    return ASE_NotPresent;
}

ASIOError ASIO2WASAPI::outputReady ()
{
    //No latency reduction can be achieved, return ASE_NotPresent
    return ASE_NotPresent;
}

ASIOError ASIO2WASAPI::getChannels (long *numInputChannels, long *numOutputChannels)
{
    if (!m_active)
        return ASE_NotPresent;

    if (numInputChannels)
        *numInputChannels = m_nInputChannels;
	if (numOutputChannels)
        *numOutputChannels = m_nChannels;
	return ASE_OK;
}

ASIOError ASIO2WASAPI::controlPanel()
{   
    if (m_hControlPanelHandle) DestroyWindow(m_hControlPanelHandle);
        
    HWND parentWindow = m_hAppWindowHandle ? m_hAppWindowHandle : GetActiveWindow();   
    DialogBoxParam(g_hinstDLL, MAKEINTRESOURCE(IDD_CONTROL_PANEL), parentWindow, (DLGPROC)ControlPanelProc, (LPARAM)this);

    return ASE_OK;
}

void ASIO2WASAPI::setMostReliableFormat()
{
    m_nChannels = 2;
    m_nSampleRate = 48000;
    //m_nBufferSize = m_wasapiExclusiveMode ? 10 : 30;

    memset(&m_waveFormat,0,sizeof(m_waveFormat));
    WAVEFORMATEX& fmt = m_waveFormat.Format;
    fmt.wFormatTag = WAVE_FORMAT_PCM;
    fmt.nChannels = 2;
    fmt.nSamplesPerSec = 48000;
    fmt.nBlockAlign = 4;
    fmt.nAvgBytesPerSec = fmt.nSamplesPerSec * fmt.nBlockAlign;
    fmt.wBitsPerSample = 16;
}

ASIOBool ASIO2WASAPI::init(void* sysRef)
{
	if (m_active)
		return true;

    m_hAppWindowHandle = (HWND) sysRef;
    m_hControlPanelHandle = 0;
    pNotificationClient = NULL;
    m_hCallbackEvent = NULL;   

    HRESULT hr=S_OK;
    IMMDeviceEnumerator *pEnumerator = NULL;
    DWORD flags = 0;

    (void) CoInitialize(NULL);

    hr = CoCreateInstance(
           CLSID_MMDeviceEnumerator, NULL,
           CLSCTX_ALL, IID_IMMDeviceEnumerator,
           (void**)&pEnumerator);
    if (FAILED(hr)) return false;
    CReleaser r1(pEnumerator);

    IMMDeviceCollection *pMMDeviceCollection = NULL;
    hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pMMDeviceCollection);
    if (FAILED(hr)) return false;
    CReleaser r2(pMMDeviceCollection);
    
    UINT nDevices = 0;
    hr = pMMDeviceCollection->GetCount(&nDevices);
    if (FAILED(hr)) return false;
    
    bool bDeviceFound = false;
    
	if (!m_useDefaultDevice)
		for (UINT i = 0; i < nDevices; i++)
		{
			IMMDevice* pMMDevice = NULL;
			hr = pMMDeviceCollection->Item(i, &pMMDevice);
			if (FAILED(hr))	return false;
			CReleaser r(pMMDevice);

			vector<wchar_t> deviceId = getDeviceId(pMMDevice);
			if (deviceId.size() && m_deviceId.size() && wcscmp(&deviceId[0], &m_deviceId[0]) == 0)
			{
				m_pDevice = pMMDevice;
				m_pDevice->AddRef();
				bDeviceFound = true;
				break;
			}
		}
    
    if (!bDeviceFound)
    {//if not found or default device is used
        hr = pEnumerator->GetDefaultAudioEndpoint(
                            eRender, eConsole, &m_pDevice);
        if (FAILED(hr)) return false;
        //setMostReliableFormat();
    }
    else    
        m_deviceId = getDeviceId(m_pDevice);


    BOOL rc = FindStreamFormat(m_pDevice, m_nChannels, m_nSampleRate, m_nBufferSize, m_wasapiExclusiveMode, m_wasapiEnableResampling, m_wasapiLowLatencySharedMode, &m_waveFormat, &m_pAudioClient);
    if (!rc)
    {
        IAudioClient* pAudioClient = NULL;
        hr = m_pDevice->Activate(IID_IAudioClient, CLSCTX_ALL, NULL, (void**)&pAudioClient);
        if (FAILED(hr)) return false;
        CReleaser r(pAudioClient);

        WAVEFORMATEX* devFormat;
        hr = pAudioClient->GetMixFormat(&devFormat);
        if (SUCCEEDED(hr))
        {
           m_nChannels = !m_wasapiExclusiveMode ? devFormat->nChannels : 2;
           m_nSampleRate = devFormat->nSamplesPerSec;
           CoTaskMemFree(devFormat);
           rc = FindStreamFormat(m_pDevice, m_nChannels, m_nSampleRate, m_nBufferSize, m_wasapiExclusiveMode, m_wasapiEnableResampling, m_wasapiLowLatencySharedMode, &m_waveFormat, &m_pAudioClient);
        }        

        if (!rc)
        {//go through all devices and try to find the one that works for 16/48K
            SAFE_RELEASE(m_pDevice)
                setMostReliableFormat();

            pMMDeviceCollection = NULL;
            hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pMMDeviceCollection);
            if (FAILED(hr)) return false;
            CReleaser r3(pMMDeviceCollection);

            nDevices = 0;
            hr = pMMDeviceCollection->GetCount(&nDevices);
            if (FAILED(hr)) return false;

            for (UINT i = 0; i < nDevices; i++)
            {
                IMMDevice* pMMDevice = NULL;
                hr = pMMDeviceCollection->Item(i, &pMMDevice);
                if (FAILED(hr)) continue;
                CReleaser rr(pMMDevice);
                rc = FindStreamFormat(pMMDevice, m_nChannels, m_nSampleRate, m_nBufferSize, m_wasapiExclusiveMode, m_wasapiEnableResampling, m_wasapiLowLatencySharedMode, &m_waveFormat, &m_pAudioClient);
                if (rc)
                {
                    m_pDevice = pMMDevice;
                    rr.deactivate();
                    break;
                }
            }
            if (!m_pAudioClient)
                return false; //suitable device not found
        }
    }
    
    if (!m_pAudioClient)
        return false; 

    if (m_inputEnabled && !initCaptureDevice(pEnumerator))
    {
        shutdown();
        strcpy_s(m_errorMessage, "Recording device unavailable or does not support the selected sample rate/mode.");
        return ASIOFalse;
    }
    if (m_inputEnabled && m_preferCommonFormat &&
        (m_waveFormat.SubFormat != m_inputWaveFormat.SubFormat ||
         m_waveFormat.Format.wBitsPerSample != m_inputWaveFormat.Format.wBitsPerSample ||
         m_waveFormat.Samples.wValidBitsPerSample != m_inputWaveFormat.Samples.wValidBitsPerSample))
    {
        WAVEFORMATEXTENSIBLE oldOutput = m_waveFormat;
        int oldDuration = m_nBufferSize;
        ReleaseInitializedAudioClient(m_pAudioClient);
        WAVEFORMATEXTENSIBLE commonInput;
        if (FindStreamFormat(m_pDevice, m_nChannels, m_nSampleRate, m_nBufferSize,
            m_wasapiExclusiveMode, m_wasapiEnableResampling, m_wasapiLowLatencySharedMode,
            &m_waveFormat, &m_pAudioClient, m_pInputDevice, m_nInputChannels, &commonInput))
            m_inputWaveFormat = commonInput;
        else
        {
            // Keep the independently selected formats/channel counts/rate when
            // the devices have no common format; retain Test 2's conversion.
            m_waveFormat = oldOutput;
            m_nBufferSize = oldDuration;
            m_pAudioClient = getAudioClient(m_pDevice, &m_waveFormat.Format, m_nBufferSize,
                m_wasapiExclusiveMode, m_wasapiEnableResampling, m_wasapiLowLatencySharedMode);
            if (!m_pAudioClient)
            {
                shutdown();
                strcpy_s(m_errorMessage, "Cannot reopen playback device after common-format negotiation.");
                return ASIOFalse;
            }
        }
    }

    if (m_wasapiExclusiveMode)
    {
        UINT32 bufferSize = 0;
        hr = m_pAudioClient->GetBufferSize(&bufferSize);
        if (FAILED(hr)) return false;
        m_bufferSize = bufferSize;
    }
    else
    {
        if (m_wasapiLowLatencySharedMode)
        {
            IAudioClient3* pAudioClient3 = NULL;
            hr = m_pAudioClient->QueryInterface(&pAudioClient3);
            if (FAILED(hr) || !pAudioClient3) return false;

            CReleaser ra3(pAudioClient3);
            UINT32 currentPeriod;
            WAVEFORMATEX* format = NULL;
            hr = pAudioClient3->GetCurrentSharedModeEnginePeriod(&format, &currentPeriod);
            if (FAILED(hr)) return false;
            CoTaskMemFree(format);
            m_bufferSize = (int)currentPeriod;
        }
        else
        {
            REFERENCE_TIME hnsDefaultPeriod = 0;
            hr = m_pAudioClient->GetDevicePeriod(&hnsDefaultPeriod, NULL);
            if (FAILED(hr)) return false;
            m_bufferSize = (int)ceil(hnsDefaultPeriod * 0.0001 * m_nSampleRate * 0.001);
        }
    }
   
    m_active = true;

    pNotificationClient = new CMMNotificationClient(this);    
    pEnumerator->RegisterEndpointNotificationCallback(pNotificationClient);   
       
    return true;
}

bool ASIO2WASAPI::initCaptureDevice(IMMDeviceEnumerator* enumerator)
{
    HRESULT hr = m_inputDeviceId.empty() ?
        enumerator->GetDefaultAudioEndpoint(eCapture, eConsole, &m_pInputDevice) :
        enumerator->GetDevice(&m_inputDeviceId[0], &m_pInputDevice);
    if (FAILED(hr) || !m_pInputDevice) return false;
    // Resolve an explicit saved ID to the endpoint's canonical string. Do not
    // change an empty ID: it intentionally tracks the default recording device.
    if (!m_inputDeviceId.empty())
    {
        vector<wchar_t> canonicalId = getDeviceId(m_pInputDevice);
        if (!canonicalId.empty()) m_inputDeviceId = canonicalId;
    }
    DWORD state = 0;
    if (FAILED(m_pInputDevice->GetState(&state)) || !(state & DEVICE_STATE_ACTIVE)) return false;
    IMMEndpoint* endpoint = NULL;
    hr = m_pInputDevice->QueryInterface(__uuidof(IMMEndpoint), (void**)&endpoint);
    if (FAILED(hr)) return false;
    CReleaser endpointReleaser(endpoint);
    EDataFlow flow;
    if (FAILED(endpoint->GetDataFlow(&flow)) || flow != eCapture) return false;

    IAudioClient* client = NULL;
    hr = m_pInputDevice->Activate(IID_IAudioClient, CLSCTX_ALL, NULL, (void**)&client);
    if (FAILED(hr)) return false;
    CReleaser clientReleaser(client);
    WAVEFORMATEX* mix = NULL;
    hr = client->GetMixFormat(&mix);
    if (FAILED(hr)) return false;
    int channels = mix->nChannels;
    CoTaskMemFree(mix);
    // Keep the configured ASIO rate common to both endpoints. Do not silently
    // substitute the input endpoint's mix rate in exclusive/low latency mode.
    for (; channels > 0; --channels)
    {
        int duration = m_nBufferSize;
        if (FindStreamFormat(m_pInputDevice, channels, m_nSampleRate, duration,
            m_wasapiExclusiveMode, m_wasapiEnableResampling, m_wasapiLowLatencySharedMode, &m_inputWaveFormat))
        {
            m_nInputChannels = channels;
            return true;
        }
    }
    return false;
}

ASIOError ASIO2WASAPI::getSampleRate (ASIOSampleRate *sampleRate)
{
    if (!sampleRate)
        return ASE_InvalidParameter;
    if (!m_active)
        return ASE_NotPresent;
    *sampleRate = m_nSampleRate;
    
    return ASE_OK;
}

ASIOError ASIO2WASAPI::setSampleRate (ASIOSampleRate sampleRate)
{
    if (!m_active)
        return ASE_NotPresent;

    if (sampleRate == m_nSampleRate)
        return ASE_OK;

    ASIOError err = canSampleRate(sampleRate);
    if (err != ASE_OK)
        return err;    
    
    if (m_callbacks)
    {//ask the host ro reset us
        int nPrevSampleRate = m_nSampleRate;
        m_nSampleRate = (int)sampleRate;
        writeToRegistry();
        m_nSampleRate = nPrevSampleRate;
        m_callbacks->asioMessage(kAsioResetRequest, 0, NULL, NULL);
    }
    else return ASE_NoClock;
    /* In case of Cubase 5 getBufferSize has been called at this point and buffersize can change due to AUDCLNT_E_BUFFER_SIZE_NOT_ALIGNED which results in failed createBuffers thus no sound at all.  
    {//reinitialize us with the new sample rate
        HWND hAppWindowHandle = m_hAppWindowHandle;
        shutdown();
        readFromRegistry();
        init(hAppWindowHandle);
    }
    */
    
    return ASE_OK;
}

//all buffer sizes are in frames
ASIOError ASIO2WASAPI::getBufferSize (long *minSize, long *maxSize,
	long *preferredSize, long *granularity)
{
    if (!m_active)
        return ASE_NotPresent;

    if (minSize)
        *minSize = m_bufferSize;
    if (maxSize)
        *maxSize = m_bufferSize;
    if (preferredSize)
        *preferredSize = m_bufferSize;
    if (granularity)
        *granularity = 0;

    return ASE_OK;
}

ASIOError ASIO2WASAPI::createBuffers(ASIOBufferInfo* bufferInfos, long numChannels,
    long bufferSize, ASIOCallbacks* callbacks)
{
    if (!m_active) return ASE_NotPresent;
    if (!callbacks || !callbacks->bufferSwitch || !bufferInfos || numChannels <= 0 ||
        numChannels > m_nChannels + m_nInputChannels) return ASE_InvalidParameter;
    if (bufferSize != m_bufferSize) return ASE_InvalidMode;
    bool hasInput = false, hasOutput = false;
    for (long i = 0; i < numChannels; ++i)
    {
        const ASIOBufferInfo& info = bufferInfos[i];
        int channels = info.isInput ? m_nInputChannels : m_nChannels;
        if (info.channelNum < 0 || info.channelNum >= channels) return ASE_InvalidParameter;
        for (long j = 0; j < i; ++j)
            if (!!bufferInfos[j].isInput == !!info.isInput && bufferInfos[j].channelNum == info.channelNum)
                return ASE_InvalidParameter;
        if (info.isInput) hasInput = true; else hasOutput = true;
    }
    ASIOError err = disposeBuffers();
    if (err != ASE_OK) return err;
    if (hasInput)
    {
        int duration = m_nBufferSize; // capture alignment must not change ASIO's render block size
        m_pCaptureAudioClient = getAudioClient(m_pInputDevice, &m_inputWaveFormat.Format,
            duration, m_wasapiExclusiveMode, m_wasapiEnableResampling, m_wasapiLowLatencySharedMode);
        if (!m_pCaptureAudioClient)
        {
            strcpy_s(m_errorMessage, "Cannot open recording device (format, exclusive access or microphone permissions).");
            return ASE_HWMalfunction;
        }
        m_hCaptureEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
        HRESULT hr = m_hCaptureEvent ? m_pCaptureAudioClient->SetEventHandle(m_hCaptureEvent) : E_OUTOFMEMORY;
        UINT32 captureFrames = 0;
        if (SUCCEEDED(hr)) hr = m_pCaptureAudioClient->GetBufferSize(&captureFrames);
        if (FAILED(hr)) { disposeBuffers(); return ASE_HWMalfunction; }
        try
        {
            m_captureBuffer.configure(max((UINT32)(4 * bufferSize), captureFrames + (UINT32)(2 * bufferSize)),
                m_inputWaveFormat.Format.nBlockAlign);
        }
        catch (const std::bad_alloc&) { disposeBuffers(); return ASE_NoMemory; }
    }
    try
    {
        for (int index = 0; index < 2; ++index)
        {
            m_buffers[index].resize(m_nChannels);
            m_inputBuffers[index].resize(m_nInputChannels);
        }
        for (long i = 0; i < numChannels; ++i)
        {
            ASIOBufferInfo& info = bufferInfos[i];
            int bytes = m_waveFormat.Format.wBitsPerSample / 8;
            for (int index = 0; index < 2; ++index)
            {
                vector<BYTE>& channel = (info.isInput ? m_inputBuffers[index] : m_buffers[index])[info.channelNum];
                channel.assign(bufferSize * bytes, 0);
                info.buffers[index] = &channel[0];
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        disposeBuffers();
        for (long i = 0; i < numChannels; ++i) bufferInfos[i].buffers[0] = bufferInfos[i].buffers[1] = NULL;
        return ASE_NoMemory;
    }
    m_hasInputBuffers = hasInput;
    m_hasOutputBuffers = hasOutput;
    m_callbacks = callbacks;
    return ASE_OK;
}

ASIOError ASIO2WASAPI::disposeBuffers()
{
    ASIOError err = stop();
    if (err != ASE_OK) return err;
    m_callbacks = NULL;
    SAFE_RELEASE(m_pCaptureAudioClient)
    if (m_hCaptureEvent) CloseHandle(m_hCaptureEvent);
    m_hCaptureEvent = NULL;
    m_hasInputBuffers = m_hasOutputBuffers = false;
    for (int index = 0; index < 2; ++index)
    {
        m_buffers[index].clear();
        m_inputBuffers[index].clear();
    }
    m_captureBuffer.configure(0, 0);
    return ASE_OK;
}

ASIOError ASIO2WASAPI::getChannelInfo(ASIOChannelInfo* info)
{
    if (!m_active)
        return ASE_NotPresent;

    if (!info || info->channel < 0 || info->channel >= (info->isInput ? m_nInputChannels : m_nChannels))
        return ASE_InvalidParameter;

    info->type = getASIOSampleType(info->isInput != 0);
    info->channelGroup = info->isInput ? 1 : 0;
    const vector<vector<BYTE> >& buffers = info->isInput ? m_inputBuffers[0] : m_buffers[0];
    info->isActive = info->channel < (long)buffers.size() && !buffers[info->channel].empty() ? ASIOTrue : ASIOFalse;
    if (info->isInput)
    {
        sprintf_s(info->name, sizeof(info->name), "Input %ld", info->channel + 1);
        return ASE_OK;
    }

    vector<UINT32> speakerBits;
    speakerBits.reserve(m_nChannels);
    for (UINT32 i = 0; i < KNOWN_SPEAKERNAMES; i++)
    {
        UINT32 testedBit = 1 << i;
        if (m_waveFormat.dwChannelMask & testedBit) speakerBits.push_back(testedBit);
        if (speakerBits.size() == m_nChannels) break;
    }

    const char* speakerName = NULL;

    if (info->channel < (long)speakerBits.size())
    {
        switch (speakerBits[info->channel])
        {
        case SPEAKER_FRONT_LEFT: speakerName = "Front Left"; break;
        case SPEAKER_FRONT_RIGHT: speakerName = "Front Right"; break;
        case SPEAKER_FRONT_CENTER: speakerName = "Front Center"; break;
        case SPEAKER_LOW_FREQUENCY: speakerName = "Low Frequency"; break;
        case SPEAKER_BACK_LEFT: speakerName = "Back Left"; break;
        case SPEAKER_BACK_RIGHT: speakerName = "Back Right"; break;
        case SPEAKER_FRONT_LEFT_OF_CENTER: speakerName = "Front Left Center"; break;
        case SPEAKER_FRONT_RIGHT_OF_CENTER: speakerName = "Front Right Center"; break;
        case SPEAKER_BACK_CENTER: speakerName = "Back Center"; break;
        case SPEAKER_SIDE_LEFT: speakerName = "Side Left"; break;
        case SPEAKER_SIDE_RIGHT: speakerName = "Side Right"; break;
        case SPEAKER_TOP_CENTER: speakerName = "Top Center"; break;
        case SPEAKER_TOP_FRONT_LEFT: speakerName = "Top Front Left"; break;
        case SPEAKER_TOP_FRONT_CENTER: speakerName = "Top Front Center"; break;
        case SPEAKER_TOP_FRONT_RIGHT: speakerName = "Top Front Right"; break;
        case SPEAKER_TOP_BACK_LEFT: speakerName = "Top Back left"; break;
        case SPEAKER_TOP_BACK_CENTER: speakerName = "Top Back Center"; break;
        case SPEAKER_TOP_BACK_RIGHT: speakerName = "Top Back Right"; break;
        }
    }

    if (speakerName)
    {
        strcpy_s(info->name, sizeof(info->name), speakerName);
    }
    else
    {
        char unknownTxt[12] = "Unknown ";
        char convTxt[12] = { 0 };
        strncat_s(unknownTxt, _itoa(info->channel, convTxt, 10), 3);
        strcpy_s(info->name, sizeof(info->name), unknownTxt);
    }

    return ASE_OK;
}

ASIOError ASIO2WASAPI::canSampleRate (ASIOSampleRate sampleRate)
{
    if (!m_active)
        return ASE_NotPresent;

	int nSampleRate = static_cast<int>(sampleRate);
    if (!IsFormatSupported(m_pDevice, m_nChannels, nSampleRate, m_wasapiExclusiveMode, m_wasapiEnableResampling)) return ASE_NoClock;
    if (m_nInputChannels && !IsFormatSupported(m_pInputDevice, m_nInputChannels, nSampleRate,
        m_wasapiExclusiveMode, m_wasapiEnableResampling)) return ASE_NoClock;
    return ASE_OK;
}

unsigned __stdcall ASIO2WASAPI::AudioThreadProc(void* pThis)
{
    ASIO2WASAPI* driver = static_cast<ASIO2WASAPI*>(pThis);
    HRESULT hr = CoInitialize(NULL);
    if (SUCCEEDED(hr))
    {
        if (driver->m_hasInputBuffers) DuplexThreadProc(pThis);
        else if (driver->m_wasapiExclusiveMode) PlayThreadProc(pThis);
        else PlayThreadProcShared(pThis);
        driver->m_samplePosition = 0;
        CoUninitialize();
    }
    else driver->m_threadStartResult = hr;
    return 0;
}

ASIOError ASIO2WASAPI::start()
{
    if (!m_active || !m_callbacks) return ASE_NotPresent;
    if (m_hPlayThread)
    {
        if (WaitForSingleObject(m_hPlayThread, 0) == WAIT_TIMEOUT) return ASE_OK;
        stop();
    }
    m_bufferIndex = 0;
    m_samplePosition = 0;
    m_captureBuffer.clear();
    m_capturePrimed = false;
    m_threadStartResult = E_FAIL;
    m_hStopPlayThreadEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
    m_hThreadStartedEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
    if (!m_hStopPlayThreadEvent || !m_hThreadStartedEvent)
    {
        if (m_hStopPlayThreadEvent) CloseHandle(m_hStopPlayThreadEvent);
        if (m_hThreadStartedEvent) CloseHandle(m_hThreadStartedEvent);
        m_hStopPlayThreadEvent = m_hThreadStartedEvent = NULL;
        return ASE_NoMemory;
    }
    m_hPlayThread = (HANDLE)_beginthreadex(NULL, 0, AudioThreadProc, this, 0, &m_playThreadId);
    if (!m_hPlayThread)
    {
        CloseHandle(m_hStopPlayThreadEvent);
        CloseHandle(m_hThreadStartedEvent);
        m_hStopPlayThreadEvent = m_hThreadStartedEvent = NULL;
        return ASE_NoMemory;
    }
    HANDLE startupEvents[2] = { m_hThreadStartedEvent, m_hPlayThread };
    DWORD wait = WaitForMultipleObjects(2, startupEvents, FALSE, INFINITE);
    if (wait != WAIT_OBJECT_0 || FAILED(m_threadStartResult))
    {
        sprintf_s(m_errorMessage, sizeof(m_errorMessage), "WASAPI stream startup failed (HRESULT 0x%08lX).", (unsigned long)m_threadStartResult);
        stop();
        return ASE_HWMalfunction;
    }
    return ASE_OK;
}

ASIOError ASIO2WASAPI::stop()
{
    if (!m_hPlayThread) return ASE_OK;
    SetEvent(m_hStopPlayThreadEvent);
    // ASIO lifecycle calls must not run from bufferSwitch/asioMessage. Do not
    // deadlock if a host nevertheless attempts a synchronous stop there.
    if (GetCurrentThreadId() == m_playThreadId) return ASE_InvalidMode;
    WaitForSingleObject(m_hPlayThread, INFINITE);
    CloseHandle(m_hPlayThread);
    CloseHandle(m_hStopPlayThreadEvent);
    CloseHandle(m_hThreadStartedEvent);
    m_hThreadStartedEvent = NULL;
    m_hPlayThread = m_hStopPlayThreadEvent = NULL;
    m_playThreadId = 0;
    return ASE_OK;
}

ASIOError ASIO2WASAPI::getClockSources (ASIOClockSource *clocks, long *numSources)
{
    if (!numSources || *numSources == 0)
        return ASE_OK;
    clocks->index = 0;
    clocks->associatedChannel = -1;
    clocks->associatedGroup = -1;
    clocks->isCurrentSource = ASIOTrue;
    strcpy_s(clocks->name,"Internal clock");
    *numSources = 1;
    return ASE_OK;
}

ASIOError ASIO2WASAPI::setClockSource (long index)
{
    return (index == 0) ? ASE_OK : ASE_NotPresent;
}

ASIOError ASIO2WASAPI::getSamplePosition (ASIOSamples *sPos, ASIOTimeStamp *tStamp)
{
    if (!m_active || !m_callbacks)
        return ASE_NotPresent;
	if (tStamp)
    {
        tStamp->lo = m_theSystemTime.lo;
	    tStamp->hi = m_theSystemTime.hi;
    }
	if (sPos)
    {
        if (m_samplePosition >= twoRaisedTo32)
	    {
		    sPos->hi = (unsigned long)(m_samplePosition * twoRaisedTo32Reciprocal);
		    sPos->lo = (unsigned long)(m_samplePosition - (sPos->hi * twoRaisedTo32));
	    }
	    else
	    {
		    sPos->hi = 0;
		    sPos->lo = (unsigned long)m_samplePosition;
	    }
    }
    return ASE_OK;
}

ASIOError ASIO2WASAPI::getLatencies(long* _inputLatency, long* _outputLatency)
{
    if (!m_active || !m_callbacks)
        return ASE_NotPresent;
    if (_inputLatency)
    {
        REFERENCE_TIME latency = 0;
        *_inputLatency = m_hasInputBuffers ? m_bufferSize * (m_hasOutputBuffers ? 2 : 1) : 0;
        if (m_pCaptureAudioClient && SUCCEEDED(m_pCaptureAudioClient->GetStreamLatency(&latency)))
            *_inputLatency += refTimeToBufferSize(latency);
    }
    if (_outputLatency)
    {
        UINT32 latency = 0;
        HRESULT hr = E_FAIL;
        if (m_pAudioClient) hr = m_pAudioClient->GetBufferSize(&latency);
        if (SUCCEEDED(hr))
        {
            if (m_wasapiExclusiveMode)
                *_outputLatency = 2 * latency;
            else
                *_outputLatency = latency + m_bufferSize;
        }
        else
            *_outputLatency = 2 * m_bufferSize;
    }
    return ASE_OK;
}

