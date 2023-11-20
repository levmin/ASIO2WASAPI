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

#pragma once

#include <vector>
using namespace std;
#ifndef _INC_MMREG
    #include <MMReg.h> // needed for WAVEFORMATEXTENSIBLE
#endif
struct IMMDevice;
struct IAudioClient;
struct IAudioRenderClient;

#include "COMBaseClasses.h"
#include "asiosys.h"
#include "iasiodrv.h"

extern CLSID CLSID_ASIO2WASAPI_DRIVER;
const char * const szDescription = "ASIO2WASAPI";


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

private:

    static DWORD WINAPI PlayThreadProc(LPVOID pThis);
    static BOOL CALLBACK ControlPanelProc(HWND hwndDlg, 
         UINT message, WPARAM wParam, LPARAM lParam);
    HRESULT LoadData(IAudioRenderClient * pRenderClient);
    long refTimeToBufferSize(LONGLONG time) const;
    LONGLONG bufferSizeToRefTime(long bufferSize) const;
    void readFromRegistry();
    void writeToRegistry();
    ASIOSampleType getASIOSampleType() const;
    void shutdown();
    void clearState();
    void setMostReliableFormat();

    //fields valid before initialization
    int m_nChannels;
    int m_nSampleRate;
    int m_nBufferSize;
    char m_errorMessage[128];
    vector<wchar_t> m_deviceId;
    
    //fields filled by init()/cleaned by shutdown()
    bool m_active;
    IMMDevice * m_pDevice;
    IAudioClient * m_pAudioClient;
    WAVEFORMATEXTENSIBLE m_waveFormat;
    int m_bufferSize;           //in audio frames
    HWND m_hAppWindowHandle;

    //fields filled by createBuffers()/cleaned by disposeBuffers()
    vector< vector<BYTE> > m_buffers[2];
	ASIOCallbacks * m_callbacks;

    //fields filled by start()/cleaned by stop()
    HANDLE m_hPlayThreadIsRunningEvent;
    int m_bufferIndex;
    HANDLE m_hStopPlayThreadEvent;
    ASIOTimeStamp m_theSystemTime;
	double m_samplePosition;
};


