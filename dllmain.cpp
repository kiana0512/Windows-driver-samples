// dllmain.cpp —— 在你原始基础上做了两点：
// 1) 修正 DisableThreadLibraryCalls 的用法（传本 DLL 的 hModule，而不是 EXE 的 nullptr 句柄）
// 2) 增加 OutputDebugStringW 调试日志（可在 Sysinternals DebugView 里看到）

#include <windows.h>
#include <unknwn.h>
#include <new>
#include <atomic>
#include <cwchar>      // swprintf
#include "MyApoGuids.h"
#include "ClassFactory.h"

extern std::atomic<long> g_cDllRefs;

// 轻量日志助手（避免额外依赖）
static void DbgLog(const wchar_t* fmt, ...)
{
    wchar_t buf[512];
    va_list ap; va_start(ap, fmt);
    _vsnwprintf_s(buf, _countof(buf), _TRUNCATE, fmt, ap);
    va_end(ap);
    OutputDebugStringW(buf);
    OutputDebugStringW(L"\n");
}

// 获取当前进程可执行名（仅用于日志）
static const wchar_t* GetProcBaseName()
{
    static wchar_t name[MAX_PATH] = L"";
    if (!name[0]) {
        DWORD n = GetModuleFileNameW(nullptr, name, _countof(name));
        if (n && n < _countof(name)) {
            // 取文件名部分
            wchar_t* p = wcsrchr(name, L'\\');
            if (p && p[1]) return p + 1;
        }
    }
    // 返回全路径或空
    return name[0] ? name : L"(unknown)";
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD reason, LPVOID)
{
    switch (reason)
    {
    case DLL_PROCESS_ATTACH:
        // 重要修正：这里必须传入本 DLL 的 hModule，避免无效的 EXE 句柄
        DisableThreadLibraryCalls(hModule);
        DbgLog(L"[MyAPO] DllMain: PROCESS_ATTACH in %s (pid=%lu)", GetProcBaseName(), GetCurrentProcessId());
        break;
    case DLL_PROCESS_DETACH:
        DbgLog(L"[MyAPO] DllMain: PROCESS_DETACH in %s (pid=%lu)", GetProcBaseName(), GetCurrentProcessId());
        break;
    default:
        break;
    }
    return TRUE;
}

extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
    if (!ppv)
        return E_POINTER;
    *ppv = nullptr;

    // 打印来访 CLSID，帮助确认 audiodg/COM 请求的到底是谁
    if (IsEqualCLSID(rclsid, CLSID_MyCompanyEfxApo)) {
        DbgLog(L"[MyAPO] DllGetClassObject: CLSID match, creating ClassFactory (req IID=?).");
    } else {
        // 常见定位点：如果这里经常不匹配，表示注册/INF 指向的 CLSID 与 DLL 内导出的不一致
        DbgLog(L"[MyAPO] DllGetClassObject: CLSID mismatch, not ours.");
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    CMyClassFactory *f = new (std::nothrow) CMyClassFactory();
    if (!f)
        return E_OUTOFMEMORY;

    HRESULT hr = f->QueryInterface(riid, ppv);
    // QueryInterface 会对工厂 AddRef；我们这里调用一次 Release 平衡 new 初始引用
    f->Release();

    if (SUCCEEDED(hr)) {
        DbgLog(L"[MyAPO] DllGetClassObject: ClassFactory QI succeeded.");
    } else {
        DbgLog(L"[MyAPO] DllGetClassObject: ClassFactory QI failed, hr=0x%08X.", hr);
    }
    return hr;
}

extern "C" HRESULT __stdcall DllCanUnloadNow()
{
    long refs = g_cDllRefs.load();
    // 当引用计数为 0 时，COM 允许卸载
    HRESULT hr = (refs == 0) ? S_OK : S_FALSE;
    DbgLog(L"[MyAPO] DllCanUnloadNow: g_cDllRefs=%ld -> %s", refs, (hr==S_OK?L"S_OK":L"S_FALSE"));
    return hr;
}

// 你已有的导出保持不变（COM 注册交给 INF）
extern "C" __declspec(dllexport) HRESULT __stdcall DllRegisterServer(void)
{
    DbgLog(L"[MyAPO] DllRegisterServer called (no-op; INF handles COM registration).");
    return S_OK; // 交给 INF 写 COM 注册，这里留空即可
}

extern "C" __declspec(dllexport) HRESULT __stdcall DllUnregisterServer(void)
{
    DbgLog(L"[MyAPO] DllUnregisterServer called (no-op).");
    return S_OK;
}
