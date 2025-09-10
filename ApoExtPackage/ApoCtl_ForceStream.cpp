// ApoCtl_ForceStream.cpp — force a shared-mode stream to trigger APO load
// No link to Audioclient.lib / Mmdevapi.lib / Ole32.lib. COM entry points are loaded at runtime.
// Matching: Parent -> InstanceId -> ContainerId -> DeviceInterface_FriendlyName -> FriendlyName -> EndpointId.
// If no match, auto-dump endpoints to help copying a working substring.
//
// Build: cl ApoCtl_ForceStream.cpp /nologo /W3 /EHsc /utf-8
// Note: wrap substrings containing '&' with quotes.

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif
#define _WIN32_WINNT 0x0A00
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <mmdeviceapi.h>
#include <audioclient.h>
#include <functiondiscoverykeys_devpkey.h> // PKEY_Device_*
#include <propidl.h>   // PROPVARIANT
#include <cstdio>
#include <cwchar>
#include <cwctype>     // towupper
#include <string>
#include <vector>
#include <algorithm>

// ---- dynamic ole32 entry points (avoid import libs) ----
typedef HRESULT (WINAPI *PFN_CoInitializeEx)(LPVOID, DWORD);
typedef void    (WINAPI *PFN_CoUninitialize)(void);
typedef HRESULT (WINAPI *PFN_CoCreateInstance)(REFCLSID, LPUNKNOWN, DWORD, REFIID, LPVOID*);
typedef void    (WINAPI *PFN_CoTaskMemFree)(LPVOID);
typedef HRESULT (WINAPI *PFN_PropVariantClear)(PROPVARIANT*);

struct Ole32Fns {
    HMODULE                 h = nullptr;
    PFN_CoInitializeEx      CoInitializeEx      = nullptr;
    PFN_CoUninitialize      CoUninitialize      = nullptr;
    PFN_CoCreateInstance    CoCreateInstance    = nullptr;
    PFN_CoTaskMemFree       CoTaskMemFree       = nullptr;
    PFN_PropVariantClear    PropVariantClear    = nullptr;
} g_ole;

static bool LoadOle32() {
    g_ole.h = LoadLibraryW(L"ole32.dll");
    if (!g_ole.h) return false;
    g_ole.CoInitializeEx   = (PFN_CoInitializeEx)   GetProcAddress(g_ole.h, "CoInitializeEx");
    g_ole.CoUninitialize   = (PFN_CoUninitialize)   GetProcAddress(g_ole.h, "CoUninitialize");
    g_ole.CoCreateInstance = (PFN_CoCreateInstance) GetProcAddress(g_ole.h, "CoCreateInstance");
    g_ole.CoTaskMemFree    = (PFN_CoTaskMemFree)    GetProcAddress(g_ole.h, "CoTaskMemFree");
    g_ole.PropVariantClear = (PFN_PropVariantClear) GetProcAddress(g_ole.h, "PropVariantClear");
    return g_ole.CoInitializeEx && g_ole.CoUninitialize &&
           g_ole.CoCreateInstance && g_ole.CoTaskMemFree && g_ole.PropVariantClear;
}

static void PrintHr(const wchar_t* tag, HRESULT hr) {
    wprintf(L"%s: 0x%08X\n", tag, hr);
}

static void Usage(const wchar_t* exe) {
    wprintf(L"Usage:\n");
    wprintf(L"  %s                              (default render endpoint, 3000 ms)\n", exe);
    wprintf(L"  %s --ms N                       (duration in ms)\n", exe);
    wprintf(L"  %s --match-substr \"...\"        (match endpoint by Parent/Id/Name/Container/EndpointId)\n", exe);
    wprintf(L"  %s --dump                       (list all endpoints and properties)\n", exe);
    wprintf(L"Note: use quotes if substring contains '&'.\n");
}

// ---- utils ----
static std::wstring ToUpper(const std::wstring& s) {
    std::wstring t; t.reserve(s.size());
    for (wchar_t c : s) t.push_back((wchar_t)towupper((wint_t)c));
    return t;
}

static std::wstring GuidToString(REFGUID g) {
    wchar_t buf[64];
    swprintf(buf, 64,
        L"{%08lX-%04hX-%04hX-%02X%02X-%02X%02X%02X%02X%02X%02X}",
        (unsigned long)g.Data1,
        (unsigned short)g.Data2,
        (unsigned short)g.Data3,
        (unsigned)g.Data4[0], (unsigned)g.Data4[1],
        (unsigned)g.Data4[2], (unsigned)g.Data4[3],
        (unsigned)g.Data4[4], (unsigned)g.Data4[5],
        (unsigned)g.Data4[6], (unsigned)g.Data4[7]);
    return std::wstring(buf);
}

// Return string prop or empty; optionally return whether key existed
static std::wstring GetPropStr(IPropertyStore* ps, REFPROPERTYKEY key, bool* pHas = nullptr) {
    if (pHas) *pHas = false;
    if (!ps) return L"";
    PROPVARIANT v; PropVariantInit(&v);
    std::wstring out;
    if (SUCCEEDED(ps->GetValue(key, &v))) {
        if (v.vt == VT_LPWSTR && v.pwszVal) {
            out = v.pwszVal; if (pHas) *pHas = true;
        } else if (v.vt == VT_CLSID && v.puuid) {
            out = GuidToString(*v.puuid); if (pHas) *pHas = true;
        }
    }
    g_ole.PropVariantClear(&v);
    return out;
}

static bool ContainsI(const std::wstring& hay, const std::wstring& needle) {
    if (needle.empty()) return false;
    return ToUpper(hay).find(ToUpper(needle)) != std::wstring::npos;
}

// Get IMMDevice::GetId string
static std::wstring GetEndpointId(IMMDevice* dev) {
    if (!dev) return L"";
    LPWSTR sid = nullptr;
    std::wstring out;
    if (SUCCEEDED(dev->GetId(&sid)) && sid) {
        out = sid;
        g_ole.CoTaskMemFree(sid);
    }
    return out;
}

// Try match this endpoint by several keys
static bool DeviceMatches(IMMDevice* dev, const std::wstring& needle) {
    if (!dev || needle.empty()) return false;

    // 1) EndpointId (IMMDevice::GetId)
    std::wstring eid = GetEndpointId(dev);
    if (!eid.empty() && ContainsI(eid, needle)) return true;

    // 2) PropertyStore
    IPropertyStore* ps = nullptr;
    if (FAILED(dev->OpenPropertyStore(STGM_READ, &ps))) return false;

    bool hasParent=false, hasId=false, hasIfcName=false, hasName=false, hasCid=false;
    std::wstring parent   = GetPropStr(ps, PKEY_Device_Parent, &hasParent);
    std::wstring id       = GetPropStr(ps, PKEY_Device_InstanceId, &hasId);
    std::wstring ifcName  = GetPropStr(ps, PKEY_DeviceInterface_FriendlyName, &hasIfcName);
    std::wstring name     = GetPropStr(ps, PKEY_Device_FriendlyName, &hasName);
    std::wstring cid      = GetPropStr(ps, PKEY_Device_ContainerId, &hasCid);

    bool hit =
        (hasParent && ContainsI(parent, needle)) ||
        (hasId     && ContainsI(id, needle))     ||
        (hasCid    && ContainsI(cid, needle))    ||
        (hasIfcName&& ContainsI(ifcName, needle))||
        (hasName   && ContainsI(name, needle));

    ps->Release();
    return hit;
}

struct Picked {
    IMMDevice* dev = nullptr;
    EDataFlow flow{};
};

// Enumerate both eRender and eCapture, try to match; collect for dump
static bool FindEndpoint(const std::wstring& needle, Picked& out, std::vector<IMMDevice*>& all) {
    IMMDeviceEnumerator* en = nullptr;
    HRESULT hr = g_ole.CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
                                        __uuidof(IMMDeviceEnumerator), (void**)&en);
    if (FAILED(hr)) { PrintHr(L"CoCreateInstance(MMDeviceEnumerator) failed", hr); return false; }

    bool found = false;
    for (int pass = 0; pass < 2 && !found; ++pass) {
        EDataFlow flow = (pass==0) ? eRender : eCapture;
        IMMDeviceCollection* col = nullptr;
        if (FAILED(en->EnumAudioEndpoints(flow, DEVICE_STATE_ACTIVE, &col))) continue;

        UINT count = 0; col->GetCount(&count);
        for (UINT i = 0; i < count; ++i) {
            IMMDevice* dev = nullptr;
            if (FAILED(col->Item(i, &dev))) continue;
            all.push_back(dev); // keep for dump later

            if (!needle.empty() && DeviceMatches(dev, needle)) {
                out.dev = dev; out.flow = flow;
                found = true;
                // do not release dev; ownership moves to out
                break;
            }
        }
        col->Release();
    }

    // fallback: default render
    if (!found) {
        IMMDevice* defDev = nullptr;
        HRESULT hr2 = en->GetDefaultAudioEndpoint(eRender, eConsole, &defDev);
        if (SUCCEEDED(hr2)) {
            out.dev = defDev; out.flow = eRender; found = true;
        }
    }

    en->Release();
    return found;
}

static void DumpEndpoints(const std::vector<IMMDevice*>& all) {
    wprintf(L"== Dumping active endpoints (render + capture) ==\n");
    for (size_t i = 0; i < all.size(); ++i) {
        IMMDevice* dev = all[i];
        std::wstring eid = GetEndpointId(dev);

        IPropertyStore* ps = nullptr;
        std::wstring id, par, cid, ifn, fn;
        if (SUCCEEDED(dev->OpenPropertyStore(STGM_READ, &ps))) {
            id  = GetPropStr(ps, PKEY_Device_InstanceId);
            par = GetPropStr(ps, PKEY_Device_Parent);
            cid = GetPropStr(ps, PKEY_Device_ContainerId);
            ifn = GetPropStr(ps, PKEY_DeviceInterface_FriendlyName);
            fn  = GetPropStr(ps, PKEY_Device_FriendlyName);
            ps->Release();
        }
        wprintf(L"[%u]\n", (unsigned)i);
        wprintf(L"  EndpointId : %s\n", eid.c_str());
        wprintf(L"  InstanceId : %s\n", id.c_str());
        wprintf(L"  Parent     : %s\n", par.c_str());
        wprintf(L"  ContainerId: %s\n", cid.c_str());
        wprintf(L"  IfcName    : %s\n", ifn.c_str());
        wprintf(L"  Name       : %s\n", fn.c_str());
    }
}

int wmain(int argc, wchar_t** argv)
{
    // Parse args
    DWORD runMs = 3000;
    std::wstring needle;
    bool wantDump = false;

    for (int i = 1; i < argc; ++i) {
        if (!_wcsicmp(argv[i], L"--ms") && i + 1 < argc) {
            runMs = _wtoi(argv[++i]);
        } else if (!_wcsicmp(argv[i], L"--match-substr") && i + 1 < argc) {
            needle = argv[++i];
        } else if (!_wcsicmp(argv[i], L"--dump")) {
            wantDump = true;
        } else if (!_wcsicmp(argv[i], L"--help") || !_wcsicmp(argv[i], L"-h") ) {
            Usage(argv[0]); return 0;
        }
    }

    if (!LoadOle32()) { wprintf(L"[!] Failed to load ole32 exports\n"); return 1; }
    HRESULT hr = g_ole.CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(hr)) { PrintHr(L"CoInitializeEx failed", hr); return 1; }

    // Find endpoint
    Picked picked{};
    std::vector<IMMDevice*> all; all.reserve(16);
    bool ok = FindEndpoint(needle, picked, all);
    if (!ok || !picked.dev) {
        wprintf(L"[!] Could not enumerate endpoints.\n");
        g_ole.CoUninitialize();
        return 1;
    }

    // Show chosen endpoint info
    {
        std::wstring eid = GetEndpointId(picked.dev);
        IPropertyStore* ps = nullptr;
        std::wstring id, par, cid, nm;
        if (SUCCEEDED(picked.dev->OpenPropertyStore(STGM_READ, &ps))) {
            id  = GetPropStr(ps, PKEY_Device_InstanceId);
            par = GetPropStr(ps, PKEY_Device_Parent);
            cid = GetPropStr(ps, PKEY_Device_ContainerId);
            nm  = GetPropStr(ps, PKEY_Device_FriendlyName);
            ps->Release();
        }
        wprintf(L"[i] Using endpoint:\n");
        wprintf(L"    EndpointId : %s\n", eid.c_str());
        wprintf(L"    InstanceId : %s\n", id.c_str());
        wprintf(L"    Parent     : %s\n", par.c_str());
        wprintf(L"    ContainerId: %s\n", cid.c_str());
        wprintf(L"    Name       : %s\n", nm.c_str());
    }

    // Dump if requested or mismatch suspicion
    if (wantDump || !needle.empty()) {
        bool hit = false;
        std::wstring eid = GetEndpointId(picked.dev);
        if (ContainsI(eid, needle)) hit = true;
        else {
            IPropertyStore* ps = nullptr;
            if (SUCCEEDED(picked.dev->OpenPropertyStore(STGM_READ, &ps))) {
                hit = ContainsI(GetPropStr(ps, PKEY_Device_Parent), needle) ||
                      ContainsI(GetPropStr(ps, PKEY_Device_InstanceId), needle) ||
                      ContainsI(GetPropStr(ps, PKEY_Device_ContainerId), needle) ||
                      ContainsI(GetPropStr(ps, PKEY_DeviceInterface_FriendlyName), needle) ||
                      ContainsI(GetPropStr(ps, PKEY_Device_FriendlyName), needle);
                ps->Release();
            }
        }
        if (!hit || wantDump) DumpEndpoints(all);
    }

    // Activate IAudioClient on chosen endpoint (render path)
    IAudioClient* ac = nullptr;
    hr = picked.dev->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, (void**)&ac);
    if (FAILED(hr)) { PrintHr(L"Activate(IAudioClient) failed", hr); for (auto* d: all) if (d!=picked.dev) d->Release(); picked.dev->Release(); g_ole.CoUninitialize(); return 1; }

    WAVEFORMATEX* mix = nullptr;
    hr = ac->GetMixFormat(&mix);
    if (FAILED(hr)) { PrintHr(L"GetMixFormat failed", hr); ac->Release(); for (auto* d: all) if (d!=picked.dev) d->Release(); picked.dev->Release(); g_ole.CoUninitialize(); return 1; }

    REFERENCE_TIME dur = 2 * 10'000'000; // 2s
    hr = ac->Initialize(AUDCLNT_SHAREMODE_SHARED, 0, dur, 0, mix, nullptr);
    if (FAILED(hr)) { PrintHr(L"IAudioClient::Initialize failed", hr); g_ole.CoTaskMemFree(mix); ac->Release(); for (auto* d: all) if (d!=picked.dev) d->Release(); picked.dev->Release(); g_ole.CoUninitialize(); return 1; }

    UINT32 bufFrames = 0; ac->GetBufferSize(&bufFrames);

    IAudioRenderClient* rc = nullptr;
    hr = ac->GetService(__uuidof(IAudioRenderClient), (void**)&rc);
    if (FAILED(hr)) { PrintHr(L"GetService(IAudioRenderClient) failed", hr); g_ole.CoTaskMemFree(mix); ac->Release(); for (auto* d: all) if (d!=picked.dev) d->Release(); picked.dev->Release(); g_ole.CoUninitialize(); return 1; }

    // Prefill silence
    {
        BYTE* p = nullptr;
        if (SUCCEEDED(rc->GetBuffer(bufFrames, &p))) {
            ZeroMemory(p, bufFrames * mix->nBlockAlign);
            rc->ReleaseBuffer(bufFrames, 0);
        }
    }

    hr = ac->Start();
    if (FAILED(hr)) { PrintHr(L"IAudioClient::Start failed", hr); rc->Release(); g_ole.CoTaskMemFree(mix); ac->Release(); for (auto* d: all) if (d!=picked.dev) d->Release(); picked.dev->Release(); g_ole.CoUninitialize(); return 1; }

    wprintf(L"[i] Stream started (shared). Running for %u ms...\n", runMs);
    Sleep(runMs);
    ac->Stop();

    rc->Release();
    g_ole.CoTaskMemFree(mix);
    ac->Release();

    for (auto* d: all) if (d!=picked.dev) d->Release();
    picked.dev->Release();

    g_ole.CoUninitialize();
    return 0;
}
