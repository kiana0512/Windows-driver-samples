// dllmain.cpp
#include <windows.h>
#include <unknwn.h>
#include "MyApoGuids.h"
#include "ClassFactory.h"

extern std::atomic<long> g_cDllRefs;

BOOL APIENTRY DllMain(HMODULE, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        DisableThreadLibraryCalls(GetModuleHandleW(nullptr));
    }
    return TRUE;
}

extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
    if (!ppv)
        return E_POINTER;
    *ppv = nullptr;

    if (rclsid != CLSID_MyCompanyEfxApo)
        return CLASS_E_CLASSNOTAVAILABLE;

    CMyClassFactory *f = new (std::nothrow) CMyClassFactory();
    if (!f)
        return E_OUTOFMEMORY;
    HRESULT hr = f->QueryInterface(riid, ppv);
    f->Release();
    return hr;
}

extern "C" HRESULT __stdcall DllCanUnloadNow()
{
    return (g_cDllRefs.load() == 0) ? S_OK : S_FALSE;
}

// exports.cpp 或 dllmain.cpp 末尾
extern "C" __declspec(dllexport) HRESULT __stdcall DllRegisterServer(void)
{
    return S_OK; // 交给 INF 写 COM 注册，这里留空即可
}
extern "C" __declspec(dllexport) HRESULT __stdcall DllUnregisterServer(void)
{
    return S_OK;
}