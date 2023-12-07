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
#include "COMBaseClasses.h"
#include "ASIO2WASAPI.h"

LONG UnregisterAsioDriver (CLSID clsid,const char *szDllPathName,const char *szregname);
LONG RegisterAsioDriver (CLSID clsid,const char *szDllPathName,const char *szregname,const char *szasiodesc,const char *szthreadmodel);


HINSTANCE hinstance;

static CFactoryTemplate s_Templates[1] = {
    {L"ASIO2WASAPI", &CLSID_ASIO2WASAPI_DRIVER, ASIO2WASAPI::CreateInstance
    } 
};
static int s_cTemplates = sizeof(s_Templates) / sizeof(s_Templates[0]);

static void InitClasses(BOOL bLoading)
{
   for (int i = 0; i < s_cTemplates; i++) {
		const CFactoryTemplate * pT = &s_Templates[i];
      if (pT->m_lpfnInit != NULL) {
			(*pT->m_lpfnInit)(bLoading, pT->m_ClsID);
      }
   }
}

class CClassFactory : public IClassFactory
{

private:
    const CFactoryTemplate * m_pTemplate;

    ULONG m_cRef;

    static int m_cLocked;
public:
    CClassFactory(const CFactoryTemplate *);

    // IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, void ** ppv);
    STDMETHODIMP_(ULONG)AddRef();
    STDMETHODIMP_(ULONG)Release();

    // IClassFactory
    STDMETHODIMP CreateInstance(LPUNKNOWN pUnkOuter, REFIID riid, void **pv);
    STDMETHODIMP LockServer(BOOL fLock);

    // allow DLLGetClassObject to know about global server lock status
    static BOOL IsLocked() {
        return (m_cLocked > 0);
    };
};


HINSTANCE g_hinstDLL;

BOOL WINAPI DllMain(
  __in  HINSTANCE hinstDLL,
  __in  DWORD fdwReason,
  __in  LPVOID 
)
{
	switch (fdwReason) {
	
		case DLL_PROCESS_ATTACH:
			g_hinstDLL = hinstDLL;
#ifndef _MT
            DisableThreadLibraryCalls(hinstDLL); //Do not call this in case of static crt
#endif           

			hinstance = hinstDLL;
			InitClasses(TRUE);
			
			break;

		case DLL_PROCESS_DETACH:
			InitClasses(FALSE);
			break;
    }
    return TRUE;
}

STDAPI DllGetClassObject(REFCLSID rClsID,REFIID riid,void **pv)
{

    if (!(riid == IID_IUnknown) && !(riid == IID_IClassFactory)) {
            return E_NOINTERFACE;
    }

    // traverse the array of templates looking for one with this
    // class id
    for (int i = 0; i < s_cTemplates; i++) {
        const CFactoryTemplate * pT = &s_Templates[i];
        if (pT->IsClassID(rClsID)) {

            // found a template - make a class factory based on this
            // template

            *pv = (LPVOID) (LPUNKNOWN) new CClassFactory(pT);
            if (*pv == NULL) {
                return E_OUTOFMEMORY;
            }
            ((LPUNKNOWN)*pv)->AddRef();
            return NOERROR;
        }
    }
    return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllCanUnloadNow()
{
   if (CClassFactory::IsLocked() || CBaseObject::ObjectsActive()) {
	
		return S_FALSE;
   } 
	else {
		return S_OK;
   }
}


HRESULT DllRegisterServer()
{
    char szDllPathName[MAX_PATH] = {0};
    GetModuleFileName(g_hinstDLL,szDllPathName,MAX_PATH);
    LONG	rc = RegisterAsioDriver (CLSID_ASIO2WASAPI_DRIVER,szDllPathName,szDescription,szDescription,"Apartment");

	if (rc) {
		MessageBox(NULL,(LPCTSTR)"DllRegisterServer failed!",szDescription,MB_OK);
		return -1;
	}

	return S_OK;
}

HRESULT DllUnregisterServer()
{
    char szDllPathName[MAX_PATH] = {0};
    GetModuleFileName(g_hinstDLL,szDllPathName,MAX_PATH);
	LONG	rc = UnregisterAsioDriver (CLSID_ASIO2WASAPI_DRIVER,szDllPathName,szDescription);

	if (rc) {
		MessageBox(NULL,(LPCTSTR)"DllUnregisterServer failed!",szDescription,MB_OK);
		return -1;
	}

	return S_OK;
}


// process-wide dll locked state
int CClassFactory::m_cLocked = 0;

CClassFactory::CClassFactory(const CFactoryTemplate *pTemplate)
{
    m_cRef = 0;
    m_pTemplate = pTemplate;
}


STDMETHODIMP CClassFactory::QueryInterface(REFIID riid,void **ppv)
{
    *ppv = NULL;

    // any interface on this object is the object pointer.
    if ((riid == IID_IUnknown) || (riid == IID_IClassFactory)) {
        *ppv = (LPVOID) this;
	// AddRef returned interface pointer
        ((LPUNKNOWN) *ppv)->AddRef();
        return NOERROR;
    }

    return ResultFromScode(E_NOINTERFACE);
}


STDMETHODIMP_(ULONG) CClassFactory::AddRef()
{
    return ++m_cRef;
}

STDMETHODIMP_(ULONG) CClassFactory::Release()
{
	LONG	rc;

    if (--m_cRef == 0) {
		delete this;
		rc = 0;
    } else rc = m_cRef;

	return rc;
}

STDMETHODIMP CClassFactory::CreateInstance(LPUNKNOWN pUnkOuter,REFIID riid,void **pv)
{
    /* Enforce the normal OLE rules regarding interfaces and delegation */

    if (pUnkOuter != NULL) {
        if (IsEqualIID(riid,IID_IUnknown) == FALSE) {
            return ResultFromScode(E_NOINTERFACE);
        }
    }

    /* Create the new object through the derived class's create function */

    HRESULT hr = NOERROR;
    CUnknown *pObj = m_pTemplate->CreateInstance(pUnkOuter, &hr);

    if (pObj == NULL) {
        return E_OUTOFMEMORY;
    }

    /* Delete the object if we got a construction error */

	if (FAILED(hr)) {
		delete pObj;
      return hr;
   }

    /* Get a reference counted interface on the object */

    /* We wrap the non-delegating QI with NDAddRef & NDRelease. */
    /* This protects any outer object from being prematurely    */
    /* released by an inner object that may have to be created  */
    /* in order to supply the requested interface.              */
    pObj->NonDelegatingAddRef();
    hr = pObj->NonDelegatingQueryInterface(riid, pv);
    pObj->NonDelegatingRelease();
    /* Note that if NonDelegatingQueryInterface fails, it will  */
    /* not increment the ref count, so the NonDelegatingRelease */
    /* will drop the ref back to zero and the object will "self-*/
    /* destruct".  Hence we don't need additional tidy-up code  */
    /* to cope with NonDelegatingQueryInterface failing.        */

    return hr;
}

STDMETHODIMP CClassFactory::LockServer(BOOL fLock)
{
    if (fLock) {
        m_cLocked++;
    } else {
        m_cLocked--;
    }
    return NOERROR;
}


