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
#ifndef _INC_WINDOWS
    #include <windows.h>
#endif
#ifndef __OBJBASE_H__
    #include "objbase.h"
#endif

extern HINSTANCE g_hInst;

extern OSVERSIONINFO g_osInfo;     // Filled in by GetVersionEx

DECLARE_INTERFACE(INonDelegatingUnknown)
{
    STDMETHOD(NonDelegatingQueryInterface) (THIS_ REFIID, LPVOID *) PURE;
    STDMETHOD_(ULONG, NonDelegatingAddRef)(THIS) PURE;
    STDMETHOD_(ULONG, NonDelegatingRelease)(THIS) PURE;
};

typedef INonDelegatingUnknown *PNDUNKNOWN;

class CBaseObject
{
    // Disable the copy constructor and assignment by default so you will get
    //   compiler errors instead of unexpected behaviour if you pass objects
    //   by value or assign objects.
    CBaseObject(const CBaseObject& objectSrc);          // no implementation
    void operator=(const CBaseObject& objectSrc);       // no implementation

    static LONG m_cObjects;     /* Total number of objects active */

public:

    CBaseObject(TCHAR *pName);
    ~CBaseObject();

    static LONG ObjectsActive() {
        return m_cObjects;
    };
};

class CUnknown : public INonDelegatingUnknown,
                 public CBaseObject
{

private:
    IUnknown * const m_pUnknown; /* Owner of this object */

protected:                      /* So we can override NonDelegatingRelease() */
    volatile LONG m_cRef;       /* Number of reference counts */

public:

    CUnknown(TCHAR *pName, IUnknown * pUnk, HRESULT *phr);
    virtual ~CUnknown() {};

    IUnknown * GetOwner() const {
        return m_pUnknown;
    };

    /* Non delegating unknown implementation */

    STDMETHODIMP NonDelegatingQueryInterface(REFIID, void **);
    STDMETHODIMP_(ULONG) NonDelegatingAddRef();
    STDMETHODIMP_(ULONG) NonDelegatingRelease();

    /* Return an interface pointer to a requesting client
       performing a thread safe AddRef as necessary */

    HRESULT GetInterface(IUnknown * pUnk, void **ppv);


};

typedef CUnknown *(*LPFNNewCOMObject)(IUnknown * pUnkOuter, HRESULT *phr);
typedef void (*LPFNInitRoutine)(BOOL bLoading, const CLSID *rclsid);

/* Create one of these per object class in an array so that
   the default class factory code can create new instances */

class CFactoryTemplate {

public:

    const WCHAR *m_Name;
    const CLSID *m_ClsID;
    LPFNNewCOMObject m_lpfnNew;
    LPFNInitRoutine  m_lpfnInit;

    BOOL IsClassID(REFCLSID rclsid) const {
        return (IsEqualCLSID(*m_ClsID,rclsid));
    };

    CUnknown *CreateInstance(LPUNKNOWN pUnk, HRESULT *phr) const {
        return m_lpfnNew(pUnk, phr);
    };
};

