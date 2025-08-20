#pragma once
#include <unknwn.h>
#include <atomic>

extern std::atomic<long> g_cDllRefs;

class CMyClassFactory : public IClassFactory {
    std::atomic<ULONG> _ref{1};
public:
    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override;
    ULONG   STDMETHODCALLTYPE AddRef()  override;
    ULONG   STDMETHODCALLTYPE Release() override;

    // IClassFactory
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv) override;
    HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock) override;
};
