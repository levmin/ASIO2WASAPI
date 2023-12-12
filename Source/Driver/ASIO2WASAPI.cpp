/*  ASIO2WASAPI Universal ASIO Driver
    Copyright (C) Lev Minkovsky
    
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
#include <Mmdeviceapi.h>
#include <Audioclient.h>
#include <versionhelpers.h>
#include <process.h>
#include "Avrt.h" //used for AvSetMmThreadCharacteristics
#include <Functiondiscoverykeys_devpkey.h>
#include "ASIO2WASAPI.h"
#include "resource.h"

extern HINSTANCE g_hinstDLL;

CLSID CLSID_ASIO2WASAPI_DRIVER = { 0x3981c4c8, 0xfe12, 0x4b0f, { 0x98, 0xa0, 0xd1, 0xb6, 0x67, 0xbd, 0xa6, 0x15 } };

const CLSID CLSID_MMDeviceEnumerator = __uuidof(MMDeviceEnumerator);
const IID IID_IMMDeviceEnumerator = __uuidof(IMMDeviceEnumerator);
const IID IID_IAudioClient = __uuidof(IAudioClient);
const IID IID_IAudioRenderClient = __uuidof(IAudioRenderClient);

const char * szPrefsRegKey = "Software\\ASIO2WASAPI";

#pragma comment(lib,"Version.lib") 
wchar_t* GetFileVersion(wchar_t* result)
{
    DWORD               dwSize = 0;
    BYTE* pVersionInfo = NULL;
    VS_FIXEDFILEINFO* pFileInfo = NULL;
    UINT                pLenFileInfo = 0;
    wchar_t tmpBuff[MAX_PATH];

    GetModuleFileNameW(g_hinstDLL, tmpBuff, MAX_PATH);

    dwSize = GetFileVersionInfoSizeW(tmpBuff, NULL);
    if (dwSize == 0)
    {
        return NULL;
    }

    pVersionInfo = new BYTE[dwSize];

    if (!GetFileVersionInfoW(tmpBuff, 0, dwSize, pVersionInfo))
    {
        delete[] pVersionInfo;
        return NULL;
    }

    if (!VerQueryValue(pVersionInfo, TEXT("\\"), (LPVOID*)&pFileInfo, &pLenFileInfo))
    {
        delete[] pVersionInfo;
        return NULL;
    }

    lstrcatW(result, L"version: ");
    _ultow_s((pFileInfo->dwFileVersionMS >> 16) & 0xffff, tmpBuff, MAX_PATH, 10);
    lstrcatW(result, tmpBuff);
    lstrcatW(result, L".");
    _ultow_s((pFileInfo->dwFileVersionMS) & 0xffff, tmpBuff, MAX_PATH, 10);
    lstrcatW(result, tmpBuff);
    lstrcatW(result, L".");
    _ultow_s((pFileInfo->dwFileVersionLS >> 16) & 0xffff, tmpBuff, MAX_PATH, 10);
    lstrcatW(result, tmpBuff);
    //lstrcatW(result, L".");
    //lstrcatW(result, _ultow((pFileInfo->dwFileVersionLS) & 0xffff, tmpBuff, 10));

    return result;
}

class CReleaser 
{
    IUnknown * m_pUnknown;
public:
    CReleaser(IUnknown * pUnk) : m_pUnknown(pUnk) {}
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
    CHandleCloser(HANDLE h) : m_h(h) {}
    ~CHandleCloser() 
    {
        if (m_h != NULL)
            CloseHandle(m_h);
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
    if (flow == eRender && role == eConsole)
    {
        ASIOCallbacks* callbacks = _asio2Wasapi->getCallbacks();
        if (_asio2Wasapi->getUseDefaultDevice() && callbacks)
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
    if (FAILED(hr))
        return id;
    size_t nDeviceIdLength=wcslen(pDeviceId);
    if (nDeviceIdLength == 0)
        return id;
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
    if (SUCCEEDED(hr) && devFormat->nChannels == nChannels)
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
        UINT32 minPeriod, maxPeriod, fundamentalPeriod, defaultPeriod, currentPeriod;
        WAVEFORMATEX* format = NULL;
        hr = pAudioClient3->GetCurrentSharedModeEnginePeriod(&format, &currentPeriod);
        hr = pAudioClient3->GetSharedModeEnginePeriod(format, &defaultPeriod, &fundamentalPeriod, &minPeriod, &maxPeriod);
        if (FAILED(hr)) return NULL;
        
        UINT32 buffsizeInSamples = (UINT32)(bufferSize * (format->nSamplesPerSec * 0.001));
        while (minPeriod < (UINT32)(buffsizeInSamples * 0.44)) minPeriod += fundamentalPeriod; //derive update period from given buffer size (that cannot be set directly in low latency mode, but it seems to be 2.2 * update period)
        if (minPeriod > maxPeriod) minPeriod = maxPeriod;
        hr = pAudioClient3->InitializeSharedAudioStream(AUDCLNT_STREAMFLAGS_EVENTCALLBACK, minPeriod, format, NULL);
        if (hr == AUDCLNT_E_ENGINE_PERIODICITY_LOCKED)
            hr = pAudioClient3->InitializeSharedAudioStream(AUDCLNT_STREAMFLAGS_EVENTCALLBACK, currentPeriod, format, NULL);
        CoTaskMemFree(format);
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

        if (FAILED(hr))
            return NULL;


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
            if (FAILED(hr))
                return NULL;

            const double REFTIME_UNITS_PER_SECOND = 10000000.;
            REFERENCE_TIME hnsAlignedDuration = static_cast<REFERENCE_TIME>((REFTIME_UNITS_PER_SECOND / pWaveFormat->nSamplesPerSec * tmpBuffSize) + 0.5);
            r.deactivate();
            SAFE_RELEASE(pAudioClient);
            hr = pDevice->Activate(IID_IAudioClient, CLSCTX_ALL, NULL, (void**)&pAudioClient);
            if (FAILED(hr) || !pAudioClient)
                return false;
            r.reset(pAudioClient);
            hr = pAudioClient->Initialize(
                shareMode,
                AUDCLNT_STREAMFLAGS_EVENTCALLBACK | (!shareMode && doResampling ? (AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM | AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY) : 0),
                hnsAlignedDuration,
                shareMode ? hnsAlignedDuration : 0,
                pWaveFormat,
                NULL);
        }

        if (FAILED(hr))
            return NULL;
    }
    
    pAudioClient->GetBufferSize(&tmpBuffSize); //calculate real latency/buffer size
    bufferSize = (int)round(tmpBuffSize / (pWaveFormat->nSamplesPerSec * 0.001));

    r.deactivate();
    return pAudioClient;
}

BOOL FindStreamFormat(IMMDevice* pDevice, int nChannels, int nSampleRate, int& nbufferSize, AUDCLNT_SHAREMODE shareMode, BOOL doResampling, BOOL doLowLatency, WAVEFORMATEXTENSIBLE* pwfxt = NULL, IAudioClient** ppAudioClient = NULL)
{
    if (!pDevice)
         return FALSE;
    
    IAudioClient* pAudioClient = NULL;
    HRESULT hr = pDevice->Activate(
        IID_IAudioClient, CLSCTX_ALL,
        NULL, (void**)&pAudioClient);
    if (FAILED(hr) || !pAudioClient)
        return NULL;
    CReleaser r(pAudioClient);

    WAVEFORMATEX* devFormat;
    hr = pAudioClient->GetMixFormat(&devFormat);
    r.deactivate();
    SAFE_RELEASE(pAudioClient);

    DWORD dwChannelMask = 0;
    if (SUCCEEDED(hr) && devFormat->nChannels == nChannels)
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
   waveFormat.Format.nBlockAlign = waveFormat.Format.wBitsPerSample * waveFormat.Format.nChannels/8;
   waveFormat.Format.nAvgBytesPerSec = waveFormat.Format.nSamplesPerSec * waveFormat.Format.nBlockAlign;
   waveFormat.Format.cbSize = sizeof(WAVEFORMATEXTENSIBLE)-sizeof(WAVEFORMATEX);
   waveFormat.Samples.wValidBitsPerSample=waveFormat.Format.wBitsPerSample;
   waveFormat.dwChannelMask = dwChannelMask;
   //test for native support for 32-bit float first
   waveFormat.SubFormat = KSDATAFORMAT_SUBTYPE_IEEE_FLOAT;

   pAudioClient = getAudioClient(pDevice, (WAVEFORMATEX*)&waveFormat, nbufferSize, shareMode, doResampling, doLowLatency);
  
   if (pAudioClient)
       goto Finish;
   
   //then try integer formats, first 32-bit int
   waveFormat.SubFormat = KSDATAFORMAT_SUBTYPE_PCM;

   pAudioClient = getAudioClient(pDevice, (WAVEFORMATEX*)&waveFormat, nbufferSize, shareMode, doResampling, doLowLatency);
   
   if (pAudioClient)
       goto Finish;

   //try 24-bit containered next
    waveFormat.Samples.wValidBitsPerSample = 24;

   pAudioClient = getAudioClient(pDevice, (WAVEFORMATEX*)&waveFormat, nbufferSize, shareMode, doResampling, doLowLatency);

   if (pAudioClient)
       goto Finish; 
   
   //try 24-bit packed next
   waveFormat.Format.wBitsPerSample = 24;
   waveFormat.Format.nBlockAlign = waveFormat.Format.wBitsPerSample * waveFormat.Format.nChannels/8;
   waveFormat.Format.nAvgBytesPerSec = waveFormat.Format.nSamplesPerSec * waveFormat.Format.nBlockAlign;
   waveFormat.Samples.wValidBitsPerSample=waveFormat.Format.wBitsPerSample;

   pAudioClient = getAudioClient(pDevice,(WAVEFORMATEX*)&waveFormat, nbufferSize, shareMode, doResampling, doLowLatency);

   if (pAudioClient)
       goto Finish;

   //finally, try 16-bit   
   waveFormat.Format.wBitsPerSample = 16;
   waveFormat.Format.nBlockAlign = waveFormat.Format.wBitsPerSample * waveFormat.Format.nChannels/8;
   waveFormat.Format.nAvgBytesPerSec = waveFormat.Format.nSamplesPerSec * waveFormat.Format.nBlockAlign;
   waveFormat.Samples.wValidBitsPerSample=waveFormat.Format.wBitsPerSample;
   
   pAudioClient = getAudioClient(pDevice,(WAVEFORMATEX*)&waveFormat, nbufferSize, shareMode, doResampling, doLowLatency);

Finish:
   BOOL bSuccess = (pAudioClient!=NULL);
   if (bSuccess)        
   {
      if (pwfxt)
        memcpy_s(pwfxt,sizeof(WAVEFORMATEXTENSIBLE),&waveFormat,sizeof(WAVEFORMATEXTENSIBLE));
      if (ppAudioClient)
        *ppAudioClient = pAudioClient;
      else 
      {
          HANDLE hEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
          CHandleCloser cl(hEvent);
          pAudioClient->SetEventHandle(hEvent); //Without this you have to wait long seconds in Win7...
          SAFE_RELEASE(pAudioClient)
      }
   }
   return bSuccess;
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

ASIOSampleType ASIO2WASAPI::getASIOSampleType() const
{
    if (m_waveFormat.SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT) return ASIOSTFloat32LSB;

    switch (m_waveFormat.Format.wBitsPerSample)
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
        RegCloseKey(key);
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
        RegSetValueExW(key,szDeviceId,NULL,REG_SZ,(const BYTE *)&m_deviceId[0],size);
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
    memset(&m_waveFormat,0,sizeof(m_waveFormat));
    m_bufferIndex = 0;
    m_hAppWindowHandle = NULL;   

    //fields filled by createBuffers()/cleaned by disposeBuffers()
    m_buffers[0].clear();
    m_buffers[1].clear();
    m_callbacks = NULL;

    //fields filled by start()/cleaned by stop()
    m_hPlayThreadIsRunningEvent = NULL;
    m_bufferSize = 0;
    m_theSystemTime.hi = 0;
    m_theSystemTime.lo = 0;
    m_samplePosition = 0;
}

ASIO2WASAPI::ASIO2WASAPI (LPUNKNOWN pUnk, HRESULT *phr)
	: CUnknown("ASIO2WASAPI", pUnk, phr)
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
	disposeBuffers();
    
    if (!m_hCallbackEvent)
    {
        HANDLE hEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
        CHandleCloser cl(hEvent);
        if (m_pAudioClient) m_pAudioClient->SetEventHandle(hEvent); //Without this you have to wait long seconds in Win7...
    }
    SAFE_RELEASE(m_pAudioClient)
    
    SAFE_RELEASE(m_pDevice)
    clearState();
    
    hr = CoCreateInstance(
        CLSID_MMDeviceEnumerator, NULL,
        CLSCTX_ALL, IID_IMMDeviceEnumerator,
        (void**)&pEnumerator);
    if (FAILED(hr))
        return;
    CReleaser r2(pEnumerator);
    
    if (pNotificationClient)
    {
        pEnumerator->UnregisterEndpointNotificationCallback(pNotificationClient);
        delete(pNotificationClient);
        pNotificationClient = NULL;
    }

    if (m_hCallbackEvent)
    {
        CloseHandle(m_hCallbackEvent);
        m_hCallbackEvent = NULL;
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

    IAudioClient3* pAudioClient3 = NULL;
    hr = pAudioClient->QueryInterface(&pAudioClient3);
    if (SUCCEEDED(hr) && pAudioClient3) CReleaser ra3(pAudioClient3);
    if (!pDriver->m_wasapiExclusiveMode && SUCCEEDED(hr) && pAudioClient3)
    {
        UINT32 minPeriod, maxPeriod, fundamentalPeriod, defaultPeriod, currentPeriod;
        WAVEFORMATEX* format;
        pAudioClient3->GetCurrentSharedModeEnginePeriod(&format, &currentPeriod);
        pAudioClient3->GetSharedModeEnginePeriod(format, &defaultPeriod, &fundamentalPeriod, &minPeriod, &maxPeriod);
        CoTaskMemFree(format);
        if (minPeriod < defaultPeriod)
        {
            EnableWindow(GetDlgItem(hwndDlg, IDC_LOWLATENCY), TRUE);
        }
        else
        {
            SendDlgItemMessage(hwndDlg, IDC_LOWLATENCY, BM_SETCHECK, BST_UNCHECKED, 0);
            EnableWindow(GetDlgItem(hwndDlg, IDC_LOWLATENCY), FALSE);
            pDriver->m_wasapiLowLatencySharedMode = FALSE;
        }

    }
    else
    {
        SendDlgItemMessage(hwndDlg, IDC_LOWLATENCY, BM_SETCHECK, BST_UNCHECKED, 0);
        EnableWindow(GetDlgItem(hwndDlg, IDC_LOWLATENCY), FALSE);
        pDriver->m_wasapiLowLatencySharedMode = FALSE;
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

BOOL CALLBACK ASIO2WASAPI::ControlPanelProc(HWND hwndDlg, 
        UINT message, WPARAM wParam, LPARAM lParam)
{ 
    static ASIO2WASAPI * pDriver = NULL;
    static vector< vector<wchar_t> > deviceStringIds; 
    
    switch (message) 
    { 
         case WM_DESTROY:
            if (pDriver) pDriver->m_hControlPanelHandle = 0;
            pDriver = NULL;
            deviceStringIds.clear();
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

            case IDC_DEVICE:
            {
                if (HIWORD(wParam) == CBN_SELCHANGE)
                {
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
                    if (FAILED(hr))
                        return 0;
                    CReleaser r1(pEnumerator);

                    if (!pDriver->getUseDefaultDevice())
                    {                      
                        IMMDeviceCollection* pMMDeviceCollection = NULL;
                        hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pMMDeviceCollection);
                        if (FAILED(hr))
                            return 0;
                        CReleaser r2(pMMDeviceCollection);

                        UINT nDevices = 0;
                        hr = pMMDeviceCollection->GetCount(&nDevices);
                        if (FAILED(hr))
                            return 0;

                        for (UINT i = 0; i < nDevices; i++)
                        {
                            IMMDevice* pMMDevice = NULL;
                            hr = pMMDeviceCollection->Item(i, &pMMDevice);
                            if (FAILED(hr))
                                continue;
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
                        //find this device
                        IMMDevice * pDevice = NULL;
                        {
                            IMMDeviceEnumerator *pEnumerator = NULL;
                            HRESULT hr = CoCreateInstance(
                                   CLSID_MMDeviceEnumerator, NULL,
                                   CLSCTX_ALL, IID_IMMDeviceEnumerator,
                                   (void**)&pEnumerator);
                            if (FAILED(hr))
                                return 0;
                            CReleaser r1(pEnumerator);

                            IMMDeviceCollection *pMMDeviceCollection = NULL;
                            hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pMMDeviceCollection);
                            if (FAILED(hr)) 
                                return 0;
                            CReleaser r2(pMMDeviceCollection);
    
                            UINT nDevices = 0;
                            hr = pMMDeviceCollection->GetCount(&nDevices);
                            if (FAILED(hr)) 
                                return 0;
    
							if (!pDriver->getUseDefaultDevice())
								for (UINT i = 0; i < nDevices; i++)
								{
									IMMDevice* pMMDevice = NULL;
									hr = pMMDeviceCollection->Item(i, &pMMDevice);
									if (FAILED(hr))
										continue;
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
                            CCallbackCaller(ASIOCallbacks * pCallbacks) : m_pCallbacks(pCallbacks) {}
                            ~CCallbackCaller() 
                            {
                                if (m_pCallbacks)
                                    m_pCallbacks->asioMessage(kAsioResetRequest,0,NULL,NULL);
                            }
                        } caller(pDriver->m_callbacks);
                        
                        //shut down the driver so no exclusive WASAPI connection would stand in our way
                        HWND hAppWindowHandle = pDriver->m_hAppWindowHandle;
                        pDriver->shutdown();

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
                            strcat_s(msgTxt, itoa(nBufferSize, convTxt, 10));
                            strcat_s(msgTxt, " ms.");
                            MessageBox(hwndDlg, msgTxt, szDescription, MB_OK | MB_ICONINFORMATION);
                        }
                        
                        //copy selected device/sample rate/channel combination into the driver
                        pDriver->m_nSampleRate = nSampleRate;
                        pDriver->m_nChannels = nChannels;
                        pDriver->m_nBufferSize = nBufferSize;
                        pDriver->m_deviceId.resize(selectedDeviceId.size());
                        wcscpy_s(&pDriver->m_deviceId[0],selectedDeviceId.size(),&selectedDeviceId.at(0));
                        //try to init the driver
                        if (pDriver->init(hAppWindowHandle) == ASIOFalse)
                        {    
                            MessageBox(hwndDlg,"ASIO driver failed to initialize",szDescription,MB_OK);
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
            pDriver = (ASIO2WASAPI*)lParam;
            if (!pDriver)
                return FALSE;
            pDriver->m_hControlPanelHandle = hwndDlg;
#ifdef _WIN64           
            SetWindowText(hwndDlg, "ASIO2WASAPI x64");
#else	
            SetWindowText(hwndDlg, "ASIO2WASAPI x86");
#endif	
            wchar_t fileversionBuff[32] = { 0 }; ;
            SetWindowTextW(GetDlgItem(hwndDlg, IDC_VERSIONINFO), GetFileVersion(fileversionBuff));
            
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
                rcOwner.left + (rc.right / 2),
                rcOwner.top + (rc.bottom / 2),
                0, 0,
                SWP_NOSIZE);

            if (GetDlgCtrlID((HWND)wParam) != IDC_DEVICE) SetFocus(GetDlgItem(hwndDlg, IDC_DEVICE));

            //SetDlgItemInt(hwndDlg,IDC_CHANNELS,(UINT)pDriver->m_nChannels,TRUE);
            //SetDlgItemInt(hwndDlg, IDC_SAMPLE_RATE, (UINT)pDriver->m_nSampleRate, TRUE);
            SetDlgItemInt(hwndDlg, IDC_BUFFERSIZE, (UINT)pDriver->m_nBufferSize, TRUE);

            IMMDeviceEnumerator *pEnumerator = NULL;
            DWORD flags = 0;
            CoInitialize(NULL);

            HRESULT hr = CoCreateInstance(
                   CLSID_MMDeviceEnumerator, NULL,
                   CLSCTX_ALL, IID_IMMDeviceEnumerator,
                   (void**)&pEnumerator);
            if (FAILED(hr))
                return false;
            CReleaser r1(pEnumerator);

            IMMDeviceCollection *pMMDeviceCollection = NULL;
            hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pMMDeviceCollection);
            if (FAILED(hr)) 
                return false;
            CReleaser r2(pMMDeviceCollection);
    
            UINT nDevices = 0;
            hr = pMMDeviceCollection->GetCount(&nDevices);
            if (FAILED(hr)) 
                return false;
                        
    
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
                if (FAILED(hr)) 
                    return false;
                CReleaser r(pMMDevice);

                vector<wchar_t> deviceId = getDeviceId(pMMDevice);
                if (deviceId.size() == 0)
                    return false;
                deviceIds.push_back(deviceId);

                IPropertyStore *pPropertyStore;
                hr = pMMDevice->OpenPropertyStore(STGM_READ, &pPropertyStore);
                if (FAILED(hr)) 
                    return false;
                CReleaser r2(pPropertyStore);
                PROPVARIANT var; PropVariantInit(&var);
                hr = pPropertyStore->GetValue(PKEY_Device_FriendlyName,&var);
                if (FAILED(hr))
                    return false;
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


            SendDlgItemMessage(hwndDlg, IDC_BUFFERSIZE, EM_LIMITTEXT, 4, 0);

            return TRUE;
            }
            break;
    } 
    return FALSE; 
} 

#define RETURN_ON_ERROR(hres)  \
              if (FAILED(hres)) return;

void ASIO2WASAPI::PlayThreadProcShared(LPVOID pThis)
{
    ASIO2WASAPI* pDriver = static_cast<ASIO2WASAPI*>(pThis);
    struct CExitEventSetter
    {
        HANDLE& m_hEvent;
        CExitEventSetter(ASIO2WASAPI* pDriver) :m_hEvent(pDriver->m_hPlayThreadIsRunningEvent)
        {
            m_hEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
        }
        ~CExitEventSetter()
        {
            SetEvent(m_hEvent);
            CloseHandle(m_hEvent);
            m_hEvent = NULL;
        }
    } setter(pDriver);

    HRESULT hr = S_OK;

    IAudioClient* pAudioClient = pDriver->m_pAudioClient;
    IAudioRenderClient* pRenderClient = NULL;
    BYTE* pData = NULL;

    hr = CoInitialize(NULL);
    RETURN_ON_ERROR(hr)    
   
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
    if (hAv) AvSetMmThreadPriority(hAv, AVRT_PRIORITY_CRITICAL);

    // Pre-load the first buffer with data  
   
    UINT32 bufferFrameCount;
    UINT32 numFramesPadding;
    hr = pAudioClient->GetBufferSize(&bufferFrameCount);
    hr = pAudioClient->GetCurrentPadding(&numFramesPadding);
    RETURN_ON_ERROR(hr)
    
    UINT32 startFrames;
    if (bufferFrameCount > (4 * pDriver->m_bufferSize))
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
        if (pDriver->m_bufferSize > (int)(bufferFrameCount - numFramesPadding))
        {           
            continue; //we will get more space in next turn
        }

        hr = pDriver->LoadData(pRenderClient);
        if (hr != S_OK && hr != AUDCLNT_E_BUFFER_ERROR)
        {
            //OutputDebugString(itoa(hr, convTxt, 10));
            pDriver->m_callbacks->asioMessage(kAsioResetRequest, 0, NULL, NULL);
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
    struct CExitEventSetter
    {
        HANDLE & m_hEvent;
        CExitEventSetter(ASIO2WASAPI * pDriver):m_hEvent(pDriver->m_hPlayThreadIsRunningEvent)
        {
            m_hEvent = CreateEvent(NULL,FALSE,FALSE,NULL);
        }
        ~CExitEventSetter()
        {
            SetEvent(m_hEvent);
            CloseHandle(m_hEvent);
            m_hEvent = NULL;
        }
    } setter(pDriver);

    HRESULT hr=S_OK;

    IAudioClient *pAudioClient = pDriver->m_pAudioClient;
    IAudioRenderClient *pRenderClient = NULL;
    BYTE *pData = NULL;
                                              
    hr = CoInitialize(NULL);
    RETURN_ON_ERROR(hr)

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
    if (hAv) AvSetMmThreadPriority(hAv, AVRT_PRIORITY_CRITICAL);

    // Pre-load the first buffer with data
    // from the audio source before starting the stream.
    hr = pDriver->LoadData(pRenderClient);  //Does this actually add latency of the length of 1 buffer?     
    RETURN_ON_ERROR(hr)

    hr = pAudioClient->Start();  // Start playing.
    RETURN_ON_ERROR(hr)

    getNanoSeconds(&pDriver->m_theSystemTime);
    pDriver->m_samplePosition = 0;    
    
    if (pDriver->m_callbacks)
        pDriver->m_callbacks->bufferSwitch(1-pDriver->m_bufferIndex,ASIOTrue);

    DWORD normalInterval = ((DWORD)round(pDriver->m_bufferSize / (pDriver->m_nSampleRate * 0.001))) + 1;
    DWORD counter = 0;
    LARGE_INTEGER startTime = { 0 };
    DWORD endTime = 0;
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
        endTime = (DWORD)round((tmpCounter.QuadPart - startTime.QuadPart) * queryPerformanceUnit);
        //OutputDebugString(itoa(endTime, convTxt, 10));

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
        if (hr != S_OK && hr != AUDCLNT_E_BUFFER_ERROR)
        {
            pDriver->m_callbacks->asioMessage(kAsioResetRequest, 0, NULL, NULL);
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
        *numInputChannels = 0;
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

    CoInitialize(NULL);

    hr = CoCreateInstance(
           CLSID_MMDeviceEnumerator, NULL,
           CLSCTX_ALL, IID_IMMDeviceEnumerator,
           (void**)&pEnumerator);
    if (FAILED(hr))
        return false;
    CReleaser r1(pEnumerator);

    IMMDeviceCollection *pMMDeviceCollection = NULL;
    hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pMMDeviceCollection);
    if (FAILED(hr)) 
        return false;
    CReleaser r2(pMMDeviceCollection);
    
    UINT nDevices = 0;
    hr = pMMDeviceCollection->GetCount(&nDevices);
    if (FAILED(hr)) 
        return false;
    
    bool bDeviceFound = false;
    
	if (!m_useDefaultDevice)
		for (UINT i = 0; i < nDevices; i++)
		{
			IMMDevice* pMMDevice = NULL;
			hr = pMMDeviceCollection->Item(i, &pMMDevice);
			if (FAILED(hr))
				return false;
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
        if (FAILED(hr))
            return false;
        //setMostReliableFormat();
    }
    else    
        m_deviceId = getDeviceId(m_pDevice);


    BOOL rc = FindStreamFormat(m_pDevice, m_nChannels, m_nSampleRate, m_nBufferSize, m_wasapiExclusiveMode, m_wasapiEnableResampling, m_wasapiLowLatencySharedMode, &m_waveFormat, &m_pAudioClient);
    if (!rc)
    {
        IAudioClient* pAudioClient = NULL;
        hr = m_pDevice->Activate(IID_IAudioClient, CLSCTX_ALL, NULL, (void**)&pAudioClient);
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

            IMMDeviceCollection* pMMDeviceCollection = NULL;
            hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pMMDeviceCollection);
            if (FAILED(hr))
                return false;
            CReleaser r2(pMMDeviceCollection);

            UINT nDevices = 0;
            hr = pMMDeviceCollection->GetCount(&nDevices);
            if (FAILED(hr))
                return false;

            for (UINT i = 0; i < nDevices; i++)
            {
                IMMDevice* pMMDevice = NULL;
                hr = pMMDeviceCollection->Item(i, &pMMDevice);
                if (FAILED(hr))
                    continue;
                CReleaser r(pMMDevice);
                rc = FindStreamFormat(pMMDevice, m_nChannels, m_nSampleRate, m_nBufferSize, m_wasapiExclusiveMode, m_wasapiEnableResampling, m_wasapiLowLatencySharedMode, &m_waveFormat, &m_pAudioClient);
                if (rc)
                {
                    m_pDevice = pMMDevice;
                    r.deactivate();
                    break;
                }
            }
            if (!m_pAudioClient)
                return false; //suitable device not found
        }
    }
    
    if (!m_pAudioClient)
        return false; 

    if (m_wasapiExclusiveMode)
    {
        UINT32 bufferSize = 0;
        hr = m_pAudioClient->GetBufferSize(&bufferSize);
        if (FAILED(hr))
            return false;
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
            m_bufferSize = (int)currentPeriod;
        }
        else
        {
            REFERENCE_TIME hnsDefaultPeriod = 0;
            hr = m_pAudioClient->GetDevicePeriod(&hnsDefaultPeriod, NULL);
            if (FAILED(hr))
                return false;
            m_bufferSize = (int)ceil(hnsDefaultPeriod * 0.0001 * m_nSampleRate * 0.001);
        }
    }
   
    m_active = true;

    pNotificationClient = new CMMNotificationClient(this);    
    pEnumerator->RegisterEndpointNotificationCallback(pNotificationClient);   
       
    return true;
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

ASIOError ASIO2WASAPI::createBuffers (ASIOBufferInfo *bufferInfos, long numChannels,
	long bufferSize, ASIOCallbacks *callbacks)
{
    if (!m_active)
        return ASE_NotPresent;

    //some sanity checks
    if (!callbacks || numChannels < 0 || numChannels > m_nChannels)
        return ASE_InvalidParameter;
    if (bufferSize != m_bufferSize)
        return ASE_InvalidMode;
    for (int i = 0; i < numChannels; i++)
	{
        ASIOBufferInfo &info = bufferInfos[i];
        if (info.isInput || info.channelNum < 0 || info.channelNum >= m_nChannels)
            return ASE_InvalidMode;
    }
    
    //dispose exiting buffers
    disposeBuffers();

    m_callbacks = callbacks;
    int sampleContainerLength = m_waveFormat.Format.wBitsPerSample/8;
    int bufferByteLength=bufferSize*sampleContainerLength;

    //the very allocation
    m_buffers[0].resize(m_nChannels);
    m_buffers[1].resize(m_nChannels);
    
    for (int i = 0; i < numChannels; i++)
	{
        ASIOBufferInfo &info = bufferInfos[i];
        m_buffers[0].at(info.channelNum).resize(bufferByteLength);
        m_buffers[1].at(info.channelNum).resize(bufferByteLength);
        info.buffers[0] = &m_buffers[0].at(info.channelNum)[0];
        info.buffers[1] = &m_buffers[1].at(info.channelNum)[0];
    }		
    return ASE_OK;
}

ASIOError ASIO2WASAPI::disposeBuffers()
{
	stop();
	//wait for the play thread to finish
    WaitForSingleObject(m_hPlayThreadIsRunningEvent,INFINITE);
    m_callbacks = 0;
    m_buffers[0].clear();
    m_buffers[1].clear();
    
    return ASE_OK;
}

ASIOError ASIO2WASAPI::getChannelInfo (ASIOChannelInfo *info)
{
    if (!m_active)
        return ASE_NotPresent;

    if (info->channel < 0 || info->channel >=m_nChannels ||  info->isInput)
		return ASE_InvalidParameter;

    info->type = getASIOSampleType();
    info->channelGroup = 0;
    info->isActive = (m_buffers[0].size() > 0) ? ASIOTrue:ASIOFalse;
    const char* knownChannelNames[] =
    {
        "Front left",
        "Front right",
        "Front center",
        "Low frequency",
        "Back left",
        "Back right",
        "Front left of center",
        "Front right of center",
        "Back center",
        "Side left",
        "Side right",
        "Top center",
        "Top front left",
        "Top front center",
        "Top front right",
        "Top back left",
        "Top back center",
        "Top back right"
    };

    if (info->channel < sizeof(knownChannelNames) / sizeof(knownChannelNames[0]))
    {
        if (m_waveFormat.dwChannelMask == KSAUDIO_SPEAKER_QUAD)
        {
            switch (info->channel)
            {
            case 2: strcpy_s(info->name, sizeof(info->name), knownChannelNames[4]); break;
            case 3: strcpy_s(info->name, sizeof(info->name), knownChannelNames[5]); break;
            default: strcpy_s(info->name, sizeof(info->name), knownChannelNames[info->channel]); break;
            }
        }
        else if (m_waveFormat.dwChannelMask == KSAUDIO_SPEAKER_5POINT1_SURROUND)
        {
            switch (info->channel)
            {
            case 4: strcpy_s(info->name, sizeof(info->name), knownChannelNames[9]); break;
            case 5: strcpy_s(info->name, sizeof(info->name), knownChannelNames[10]); break;
            default: strcpy_s(info->name, sizeof(info->name), knownChannelNames[info->channel]); break;
            }
        }
        else if (m_waveFormat.dwChannelMask == KSAUDIO_SPEAKER_7POINT1_SURROUND)
        {
            switch (info->channel)
            {
            case 6: strcpy_s(info->name, sizeof(info->name), knownChannelNames[9]); break;
            case 7: strcpy_s(info->name, sizeof(info->name), knownChannelNames[10]); break;
            default: strcpy_s(info->name, sizeof(info->name), knownChannelNames[info->channel]); break;
            }
        }
        else
            strcpy_s(info->name, sizeof(info->name), knownChannelNames[info->channel]);

    }
    else
    {
        char unknownTxt[12] = "Unknown ";
        char convTxt[12] = { 0 };
        strncat_s(unknownTxt, itoa(info->channel, convTxt, 10), 3);
        strcpy_s(info->name, sizeof(info->name), unknownTxt);
    }

    return ASE_OK;
}

ASIOError ASIO2WASAPI::canSampleRate (ASIOSampleRate sampleRate)
{
    if (!m_active)
        return ASE_NotPresent;

	int nSampleRate = static_cast<int>(sampleRate);
    return IsFormatSupported(m_pDevice, m_nChannels, nSampleRate, m_wasapiExclusiveMode, m_wasapiEnableResampling) ? ASE_OK : ASE_NoClock;
}

ASIOError ASIO2WASAPI::start()
{
    if (!m_active || !m_callbacks)
        return ASE_NotPresent;
    if (m_hStopPlayThreadEvent)
        return ASE_OK;// we are already playing
    //make sure the previous play thread exited
    WaitForSingleObject(m_hPlayThreadIsRunningEvent,INFINITE);
    
    m_hStopPlayThreadEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
    if (m_wasapiExclusiveMode)
        _beginthread(PlayThreadProc, 16384, this); //Prevents thread leaks caused by CreateThread and static crt
    else
        _beginthread(PlayThreadProcShared, 16384, this); //Prevents thread leaks caused by CreateThread and static crt

    return ASE_OK;
}

ASIOError ASIO2WASAPI::stop()
{
    if (!m_active)
        return ASE_NotPresent;
    if (!m_hStopPlayThreadEvent)
        return ASE_OK; //we already stopped

    //set the thead stopping event, thus initiating the thread termination process
    SetEvent(m_hStopPlayThreadEvent);
    CloseHandle(m_hStopPlayThreadEvent);
    m_hStopPlayThreadEvent = NULL;

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
        *_inputLatency = m_bufferSize;
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

