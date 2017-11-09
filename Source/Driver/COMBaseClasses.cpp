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

LONG CBaseObject::m_cObjects = 0;

CBaseObject::CBaseObject(TCHAR *pName)
{
    InterlockedIncrement(&m_cObjects);
}


CBaseObject::~CBaseObject()
{
    InterlockedDecrement(&m_cObjects);
}

#pragma warning( disable : 4355 ) 

CUnknown::CUnknown(TCHAR *pName, IUnknown * pUnk, HRESULT *phr) 
: CBaseObject(pName) , m_cRef(0), m_pUnknown (pUnk != 0 ? pUnk : reinterpret_cast<IUnknown *>( static_cast<INonDelegatingUnknown *>(this) ))
{
}

#pragma warning( default : 4355 ) 

STDMETHODIMP CUnknown::NonDelegatingQueryInterface(REFIID riid, void ** ppv)
{
    if (riid == IID_IUnknown) {
        GetInterface((LPUNKNOWN) (PNDUNKNOWN) this, ppv);
        return NOERROR;
    } else {
        *ppv = NULL;
        return E_NOINTERFACE;
    }
}

STDMETHODIMP_(ULONG) CUnknown::NonDelegatingAddRef()
{
    LONG lRef = InterlockedIncrement( &m_cRef );
    return max(ULONG(lRef), 1ul);
}

STDMETHODIMP_(ULONG) CUnknown::NonDelegatingRelease()
{
    LONG lRef = InterlockedDecrement( &m_cRef );
    if (lRef == 0) {
        m_cRef++;
        delete this;
        return ULONG(0);
    } else {
        return max(ULONG(lRef), 1ul);
    }
}

HRESULT CUnknown::GetInterface(LPUNKNOWN pUnk, void **ppv)
{
    *ppv = pUnk;
    pUnk->AddRef();
    return NOERROR;
}

