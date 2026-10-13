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

#pragma once
#pragma warning(disable : 4996) // disable crt secure warnings
#pragma warning(disable : 26812) // disable ASIO enum warnings

#include <vector>
#include "CaptureBuffer.h"
using namespace std;
#ifndef _INC_MMREG
    #include <MMReg.h> // needed for WAVEFORMATEXTENSIBLE
#endif
struct IMMDevice;
struct IAudioClient;
struct IAudioRenderClient;

#define KNOWN_SPEAKERNAMES   18       

#define SAFE_RELEASE(punk)  \
              if ((punk) != NULL)  \
                { (punk)->Release(); (punk) = NULL; }

#include <Mmdeviceapi.h>
#include <Audioclient.h>
#include "COMBaseClasses.h"
#include "asiosys.h"
#include "iasiodrv.h"

extern CLSID CLSID_ASIO2WASAPI_DRIVER;
const char * const szDescription = "ASIO2WASAPI";

class ASIO2WASAPI;

class CMMNotificationClient : public IMMNotificationClient
{
    LONG _cRef;
    IMMDeviceEnumerator* _pEnumerator;
    ASIO2WASAPI* _asio2Wasapi;

public:
    explicit CMMNotificationClient(ASIO2WASAPI* asio2Wasapi);

    ~CMMNotificationClient();

    // IUnknown methods -- AddRef, Release, and QueryInterface

    ULONG STDMETHODCALLTYPE AddRef();

    ULONG STDMETHODCALLTYPE Release();

    HRESULT STDMETHODCALLTYPE QueryInterface(
        REFIID riid, VOID** ppvInterface);

    // Callback methods for device-event notifications.

    HRESULT STDMETHODCALLTYPE OnDefaultDeviceChanged(
        EDataFlow flow, ERole role,
        LPCWSTR pwstrDeviceId);

    HRESULT STDMETHODCALLTYPE OnDeviceAdded(LPCWSTR pwstrDeviceId);;

    HRESULT STDMETHODCALLTYPE OnDeviceRemoved(LPCWSTR pwstrDeviceId);

    HRESULT STDMETHODCALLTYPE OnDeviceStateChanged(
        LPCWSTR pwstrDeviceId,
        DWORD dwNewState);

    HRESULT STDMETHODCALLTYPE OnPropertyValueChanged(
        LPCWSTR pwstrDeviceId,
        const PROPERTYKEY key);
};


class ASIO2WASAPI : public IASIO, public CUnknown
{
public:
	ASIO2WASAPI(LPUNKNOWN pUnk, HRESULT *phr);
	~ASIO2WASAPI();

    STDMETHODIMP QueryInterface(REFIID riid, void **ppv) {     
        return GetOwner()->QueryInterface(riid,ppv);           
    };                                                         
    STDMETHODIMP_(ULONG) AddRef() {                            
        return GetOwner()->AddRef();                           
    };                                                         
    STDMETHODIMP_(ULONG) Release() {                           
        return GetOwner()->Release();                          
    };

	// Factory method
	static CUnknown *CreateInstance(LPUNKNOWN pUnk, HRESULT *phr);
	// IUnknown
	virtual HRESULT STDMETHODCALLTYPE NonDelegatingQueryInterface(REFIID riid,void **ppvObject);

    

	ASIOBool init (void* sysRef);
	void getDriverName (char *name);	// max 32 bytes incl. terminating zero
	long getDriverVersion ();
	void getErrorMessage (char *string);	// max 128 bytes incl. terminating zero

    ASIOCallbacks* getCallbacks() { return m_callbacks; }
      
	ASIOError start ();
	ASIOError stop ();

	ASIOError getChannels (long *numInputChannels, long *numOutputChannels);
	ASIOError getLatencies (long *inputLatency, long *outputLatency);
	ASIOError getBufferSize (long *minSize, long *maxSize,long *preferredSize, long *granularity);

	ASIOError canSampleRate (ASIOSampleRate sampleRate);
	ASIOError getSampleRate (ASIOSampleRate *sampleRate);
	ASIOError setSampleRate (ASIOSampleRate sampleRate);
	ASIOError getClockSources (ASIOClockSource *clocks, long *numSources);
	ASIOError setClockSource (long index);

	ASIOError getSamplePosition (ASIOSamples *sPos, ASIOTimeStamp *tStamp);
	ASIOError getChannelInfo (ASIOChannelInfo *info);

	ASIOError createBuffers (ASIOBufferInfo *bufferInfos, long numChannels,
		long bufferSize, ASIOCallbacks *callbacks);
	ASIOError disposeBuffers ();

	ASIOError controlPanel ();
	ASIOError future (long selector, void *opt);
	ASIOError outputReady ();

/// WASAPI specific
    bool getUseDefaultInputDevice() const { return m_inputEnabled && m_inputDeviceId.empty(); };
    bool getUseDefaultDevice() { return m_useDefaultDevice; };
    void setUseDefaultDevice(bool value) { m_useDefaultDevice = value; };
private:   
    //for default device changed notification
    CMMNotificationClient* pNotificationClient = NULL;
    HANDLE m_hCallbackEvent = NULL;   
    
    static unsigned __stdcall AudioThreadProc(void* pThis);
    static void DuplexThreadProc(LPVOID pThis);
    static void PlayThreadProc(LPVOID pThis);
    static void PlayThreadProcShared(LPVOID pThis);
    static BOOL CALLBACK ControlPanelProc(HWND hwndDlg, UINT message, WPARAM wParam, LPARAM lParam);    
    static void initInputFields(IMMDevice* pDevice, ASIO2WASAPI* pDriver, const HWND hwndDlg);
    static void updateLowLatencyAvailability(ASIO2WASAPI* pDriver, HWND hwndDlg,
        const vector< vector<wchar_t> >& deviceIds, const vector< vector<wchar_t> >& inputDeviceIds);
    
    HRESULT LoadData(IAudioRenderClient * pRenderClient);
    HRESULT DrainCapture(IAudioCaptureClient* pCaptureClient);
    void SwitchInputBuffer(int index);
    bool initCaptureDevice(IMMDeviceEnumerator* pEnumerator);
    long refTimeToBufferSize(REFERENCE_TIME time) const;
    LONGLONG bufferSizeToRefTime(long bufferSize) const;
    void updateFormatDiagnostics(HWND dialog) const;
    void writeToRegistry();
    void readFromRegistry();
    void shutdown();   
    ASIOSampleType getASIOSampleType(bool input = false) const;
    void clearState();
    void setMostReliableFormat();

    //fields valid before initialization
    int m_nChannels;
    int m_nSampleRate;
    int m_nBufferSize;
    char m_errorMessage[128];
    vector<wchar_t> m_deviceId;
    bool m_inputEnabled = false;
    BOOL m_preferCommonFormat = FALSE;
    vector<wchar_t> m_inputDeviceId;
    
    //fields filled by init()/cleaned by shutdown()
    bool m_active;
    IMMDevice * m_pDevice;
    IAudioClient * m_pAudioClient;
    IMMDevice* m_pInputDevice;
    IAudioClient* m_pCaptureAudioClient;
    WAVEFORMATEXTENSIBLE m_inputWaveFormat;
    int m_nInputChannels;
    HANDLE m_hCaptureEvent = NULL;
    WAVEFORMATEXTENSIBLE m_waveFormat;
    int m_bufferSize;           //in audio frames
    HWND m_hAppWindowHandle; 
    HWND m_hControlPanelHandle;
    
    //WASAPI specific
    bool m_useDefaultDevice = true;    
    AUDCLNT_SHAREMODE m_wasapiExclusiveMode = AUDCLNT_SHAREMODE_EXCLUSIVE;
    BOOL m_wasapiEnableResampling = FALSE;
    BOOL m_wasapiLowLatencySharedMode = FALSE; 

    //fields filled by createBuffers()/cleaned by disposeBuffers()
    vector< vector<BYTE> > m_buffers[2];
    vector< vector<BYTE> > m_inputBuffers[2];
    CaptureBuffer m_captureBuffer;
    bool m_hasInputBuffers;
    bool m_hasOutputBuffers;
    bool m_capturePrimed;
    ASIOCallbacks* m_callbacks;

    //fields filled by start()/cleaned by stop()
    HANDLE m_hPlayThread;
    unsigned m_playThreadId;
    HANDLE m_hThreadStartedEvent = NULL;
    HRESULT m_threadStartResult = E_FAIL;
    int m_bufferIndex;
    HANDLE m_hStopPlayThreadEvent;
    ASIOTimeStamp m_theSystemTime;
	double m_samplePosition;
};


