#include <new>
#include <atomic>
#include "ClassFactory.h"
#include "EfxApo.h"        //  必须：否则 CMyCompanyEfxApo 是不完整类型
#include "MyApoGuids.h"    // CLSID_MyCompanyEfxApo 声明

std::atomic<long> g_cDllRefs{0};

// IUnknown
HRESULT CMyClassFactory::QueryInterface(REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    if (riid == __uuidof(IUnknown) || riid == __uuidof(IClassFactory)) {
        *ppv = static_cast<IClassFactory*>(this);
        AddRef(); return S_OK;
    }
    *ppv = nullptr; return E_NOINTERFACE;
}
ULONG CMyClassFactory::AddRef()  { return ++_ref; }
ULONG CMyClassFactory::Release() { auto n = --_ref; if (!n) delete this; return n; }

// IClassFactory
HRESULT CMyClassFactory::CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    *ppv = nullptr;
    if (pUnkOuter) return CLASS_E_NOAGGREGATION;

    CMyCompanyEfxApo* pObj = new(std::nothrow) CMyCompanyEfxApo();
    if (!pObj) return E_OUTOFMEMORY;

    ++g_cDllRefs;
    HRESULT hr = pObj->QueryInterface(riid, ppv);
    pObj->Release();
    if (FAILED(hr)) --g_cDllRefs;
    return hr;
}

HRESULT CMyClassFactory::LockServer(BOOL fLock) {
    if (fLock) ++g_cDllRefs; else --g_cDllRefs;
    return S_OK;
}
