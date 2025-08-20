// ============================================================================
//  ApoCtl.cpp —— 给自研 EFX APO 发送/查询参数（IKsControl）【端点列举/选择/多路径获取】
//  要点：
//   1) “pnp 子串匹配”同时支持 USB#VID... 和 USB\VID...（分隔符归一化）
//   2) list 同时打印：接口路径(ifPath)、硬件ID(hwid)、友好名
//   3) 获取 IKsControl 的 3 条路径：IAudioClient -> IPart::Activate -> IConnector::QI
//      若仍失败，多半是 APO 未在该路径注册属性集或尚未被加载
// ============================================================================

#define _WIN32_WINNT 0x0A00

#include <windows.h>
#include <mmdeviceapi.h>
#include <audioclient.h>
#include <mmreg.h>
#include <devicetopology.h>
#include <functiondiscoverykeys_devpkey.h>
#include <initguid.h>
#include <ks.h>

// ---- 优先官方 ksproxy.h，缺失则兜底声明 IKsControl ----
#if __has_include(<ksproxy.h>)
#include <ksproxy.h>
#else
#include <Unknwn.h>
MIDL_INTERFACE("28F54685-06FD-11d2-B27A-00A0C9223196")
IKsControl : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE KsProperty(
        PKSPROPERTY Property, ULONG PropertyLength,
        PVOID PropertyData, ULONG DataLength, ULONG * BytesReturned) = 0;
    virtual HRESULT STDMETHODCALLTYPE KsMethod(
        PKSMETHOD Method, ULONG MethodLength,
        PVOID MethodData, ULONG DataLength, ULONG * BytesReturned) = 0;
    virtual HRESULT STDMETHODCALLTYPE KsEvent(
        PKSEVENT Event, ULONG EventLength,
        PVOID EventData, ULONG DataLength, ULONG * BytesReturned) = 0;
};
#endif

// ---- 额外：用 SetupAPI 通过接口路径取硬件ID ----
#include <setupapi.h>
#include <cfgmgr32.h>

#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Mmdevapi.lib")
#pragma comment(lib, "Uuid.lib")
#pragma comment(lib, "Ksproxy.lib")
#pragma comment(lib, "Setupapi.lib")
#pragma comment(lib, "Cfgmgr32.lib")

#include <cstdio>
#include <cstdlib>
#include <vector>
#include <string>
#include <algorithm>

// ===== 你的 APO 属性集与 PID 定义（需与 APO 内一致）=====
DEFINE_GUID(MYCOMPANY_APO_PROPSETID,
            0xd4d9a040, 0x8b5f, 0x4c0e, 0xaa, 0xd1, 0xaa, 0xbb, 0xcc, 0xdd, 0xee, 0xff);

enum : ULONG
{
    PID_Gain = 1,
    PID_EQBand = 2,
    PID_Reverb = 3,
    PID_Limiter = 4,
    PID_ParamsBlob = 10
};

struct EQBandParam
{
    LONG bandIndex;
    FLOAT gainLinear;
};

// ===== 宏 & 小工具 =====
#define HR_OK(hr) (SUCCEEDED((hr)))
#define HR_FAIL(hr, msg) wprintf(L"[!] %s (hr=0x%08X)\n", L##msg, (hr))
#define RETURN_IF_FAILED(hr, msg) \
    do                            \
    {                             \
        if (FAILED(hr))           \
        {                         \
            HR_FAIL(hr, msg);     \
            goto done;            \
        }                         \
    } while (0)
// 新增：去掉 "{n}." 前缀（比如 "{2}.\\?\usb#vid_..."）
// 放在文件顶部小工具区域
static std::wstring SanitizeIfPath(const std::wstring &s)
{
    if (s.size() >= 4 && s[0] == L'{')
    {
        size_t pos = s.find(L".\\");
        if (pos != std::wstring::npos)
            return s.substr(pos + 1); // 去掉 "{n}."
    }
    return s;
}

static std::wstring ToUpper(std::wstring s)
{
    std::transform(s.begin(), s.end(), s.begin(), ::towupper);
    return s;
}

// 统一把 ID 变大写，并把 '#' 归一化成 '\'，这样 USB#VID 和 USB\VID 都能匹配
static std::wstring NormalizeId(std::wstring s)
{
    for (auto &ch : s)
    {
        if (ch == L'#')
            ch = L'\\';
        ch = (wchar_t)towupper(ch);
    }
    return s;
}

// —— 默认 pnp 子串（允许留空，不传则按名字或索引/第一项）
static const wchar_t *kDefaultPnPSubstr = L"USB\\VID_0A67&PID_30A2";

// 从 endpoint 拿到“设备侧接口路径”（KS 设备接口符号链接，形如 \\?\usb#vid...#{KSCATEGORY_AUDIO}\global）
static HRESULT GetInterfacePath(IMMDevice *endpoint, std::wstring &ifPath)
{
    ifPath.clear();
    if (!endpoint)
        return E_POINTER;

    HRESULT hr = S_OK;
    IDeviceTopology *topoEp = nullptr;
    hr = endpoint->Activate(__uuidof(IDeviceTopology), CLSCTX_ALL, nullptr, (void **)&topoEp);
    if (FAILED(hr))
        return hr;

    UINT cc = 0;
    hr = topoEp->GetConnectorCount(&cc);
    if (FAILED(hr) || cc == 0)
    {
        topoEp->Release();
        return FAILED(hr) ? hr : E_FAIL;
    }

    IConnector *epConn = nullptr;
    hr = topoEp->GetConnector(0, &epConn);
    if (FAILED(hr))
    {
        topoEp->Release();
        return hr;
    }

    IConnector *devConn = nullptr;
    hr = epConn->GetConnectedTo(&devConn);
    if (FAILED(hr))
    {
        epConn->Release();
        topoEp->Release();
        return hr;
    }

    // 设备侧 Connector -> Part -> TopologyObject -> GetDeviceId() 得到接口路径
    IPart *part = nullptr;
    hr = devConn->QueryInterface(__uuidof(IPart), (void **)&part);
    if (FAILED(hr))
    {
        devConn->Release();
        epConn->Release();
        topoEp->Release();
        return hr;
    }

    IDeviceTopology *topoDev = nullptr;
    hr = part->GetTopologyObject(&topoDev);
    if (FAILED(hr))
    {
        part->Release();
        devConn->Release();
        epConn->Release();
        topoEp->Release();
        return hr;
    }

    LPWSTR devPath = nullptr;
    hr = topoDev->GetDeviceId(&devPath); // 注意：这是 KS 接口路径
    if (SUCCEEDED(hr) && devPath)
    {
        ifPath.assign(devPath);
        CoTaskMemFree(devPath);
    }
    else if (SUCCEEDED(hr))
    {
        hr = E_FAIL;
    }

    topoDev->Release();
    part->Release();
    devConn->Release();
    epConn->Release();
    topoEp->Release();

    return ifPath.empty() ? E_FAIL : S_OK;
}

// 通过 KS 接口路径(ifPath) 反查硬件 ID（SPDRP_HARDWAREID）
static HRESULT GetHardwareIdFromInterfacePath(const std::wstring &ifPath, std::wstring &hwid)
{
    hwid.clear();
    HDEVINFO hdi = SetupDiCreateDeviceInfoList(nullptr, nullptr);
    if (hdi == INVALID_HANDLE_VALUE)
        return HRESULT_FROM_WIN32(GetLastError());

    SP_DEVICE_INTERFACE_DATA ifData{};
    ifData.cbSize = sizeof(ifData);
    // GetHardwareIdFromInterfacePath() 起始处：
    std::wstring clean = SanitizeIfPath(ifPath);
    if (!SetupDiOpenDeviceInterfaceW(hdi, ifPath.c_str(), 0, &ifData))
    {
        DWORD err = GetLastError();
        SetupDiDestroyDeviceInfoList(hdi);
        return HRESULT_FROM_WIN32(err);
    }

    // 两段式获取 detail + 取得 devInfo（设备节点）
    DWORD need = 0;
    SetupDiGetDeviceInterfaceDetailW(hdi, &ifData, nullptr, 0, &need, nullptr);
    std::vector<BYTE> buf(need);
    auto *detail = reinterpret_cast<SP_DEVICE_INTERFACE_DETAIL_DATA_W *>(buf.data());
    detail->cbSize = sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA_W);
    SP_DEVINFO_DATA di{};
    di.cbSize = sizeof(di);

    if (!SetupDiGetDeviceInterfaceDetailW(hdi, &ifData, detail, need, nullptr, &di))
    {
        DWORD err = GetLastError();
        SetupDiDestroyDeviceInfoList(hdi);
        return HRESULT_FROM_WIN32(err);
    }

    WCHAR multi[4096] = {};
    if (!SetupDiGetDeviceRegistryPropertyW(hdi, &di, SPDRP_HARDWAREID, nullptr,
                                           reinterpret_cast<PBYTE>(multi), sizeof(multi), nullptr))
    {
        DWORD err = GetLastError();
        SetupDiDestroyDeviceInfoList(hdi);
        return HRESULT_FROM_WIN32(err);
    }

    // MULTI_SZ 的第一项即主硬件ID
    hwid.assign(multi);
    SetupDiDestroyDeviceInfoList(hdi);
    return hwid.empty() ? E_FAIL : S_OK;
}

enum class Flow
{
    Render,
    Capture
};

// 列出端点：友好名 + 接口路径 + 硬件ID（尽量打印全）
static void ListEndpoints(Flow flow)
{
    IMMDeviceEnumerator *e = nullptr;
    IMMDeviceCollection *col = nullptr;
    if (FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
                                __uuidof(IMMDeviceEnumerator), (void **)&e)))
        return;
    EDataFlow df = (flow == Flow::Render) ? eRender : eCapture;
    if (FAILED(e->EnumAudioEndpoints(df, DEVICE_STATE_ACTIVE, &col)))
    {
        e->Release();
        return;
    }

    UINT n = 0;
    if (FAILED(col->GetCount(&n)))
    {
        col->Release();
        e->Release();
        return;
    }
    wprintf(L"--- %s endpoints ---\n", (flow == Flow::Render) ? L"Render" : L"Capture");
    for (UINT i = 0; i < n; ++i)
    {
        IMMDevice *ep = nullptr;
        if (FAILED(col->Item(i, &ep)))
            continue;
        IPropertyStore *ps = nullptr;
        PROPVARIANT v;
        PropVariantInit(&v);

        std::wstring name = L"(unknown)";
        std::wstring ifPath, hwid;

        if (SUCCEEDED(ep->OpenPropertyStore(STGM_READ, &ps)))
        {
            if (SUCCEEDED(ps->GetValue(PKEY_Device_FriendlyName, &v)) && v.pwszVal)
                name = v.pwszVal;
            PropVariantClear(&v);
            ps->Release();
        }
        GetInterfacePath(ep, ifPath);
        if (!ifPath.empty())
            GetHardwareIdFromInterfacePath(ifPath, hwid);

        wprintf(L"[%u] %s\n     ifPath: %s\n     hwid:   %s\n",
                i, name.c_str(),
                ifPath.empty() ? L"(n/a)" : ifPath.c_str(),
                hwid.empty() ? L"(n/a)" : hwid.c_str());

        ep->Release();
    }
    col->Release();
    e->Release();
}

// 查找端点：优先 index；否则 pnp/hwid/ifPath 子串；再否则 name 子串；都没给就取第一项
static HRESULT FindEndpoint(Flow flow,
                            const std::wstring &pnpOrHwidSubstr, // 统一做 NormalizeId 后匹配
                            const std::wstring &nameSubstr,
                            int index,
                            IMMDevice **ppDev,
                            std::wstring *pIfPath /*可选返回，便于调试*/)
{
    if (pIfPath)
        pIfPath->clear();
    *ppDev = nullptr;

    HRESULT hr = S_OK;
    IMMDeviceEnumerator *e = nullptr;
    IMMDeviceCollection *col = nullptr;

    hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
                          __uuidof(IMMDeviceEnumerator), (void **)&e);
    if (FAILED(hr))
        return hr;

    EDataFlow df = (flow == Flow::Render) ? eRender : eCapture;
    hr = e->EnumAudioEndpoints(df, DEVICE_STATE_ACTIVE, &col);
    if (FAILED(hr))
    {
        e->Release();
        return hr;
    }

    UINT n = 0;
    hr = col->GetCount(&n);
    if (FAILED(hr))
    {
        col->Release();
        e->Release();
        return hr;
    }

    if (index >= 0 && (UINT)index < n)
    {
        hr = col->Item(index, ppDev);
        if (SUCCEEDED(hr) && pIfPath)
        {
            std::wstring ifp;
            GetInterfacePath(*ppDev, ifp);
            *pIfPath = ifp;
        }
        col->Release();
        e->Release();
        return hr;
    }

    const std::wstring key = NormalizeId(pnpOrHwidSubstr);
    const std::wstring nameKey = ToUpper(nameSubstr);

    for (UINT i = 0; i < n; ++i)
    {
        IMMDevice *ep = nullptr;
        if (FAILED(col->Item(i, &ep)))
            continue;
        bool ok = false;

        // 先按 hwid/ifPath 子串（统一归一化）匹配
        if (!key.empty())
        {
            std::wstring ifp;
            if (SUCCEEDED(GetInterfacePath(ep, ifp)) && !ifp.empty())
            {
                std::wstring hwid;
                std::wstring cleanIf = SanitizeIfPath(ifp);
                GetHardwareIdFromInterfacePath(ifp, hwid);
                if (!hwid.empty())
                {
                    if (NormalizeId(hwid).find(key) != std::wstring::npos)
                        ok = true;
                }
                if (!ok)
                {
                    if (NormalizeId(ifp).find(key) != std::wstring::npos)
                        ok = true;
                }
            }
        }

        // 再按名字子串匹配
        if (!ok && !nameKey.empty())
        {
            IPropertyStore *ps = nullptr;
            PROPVARIANT v;
            PropVariantInit(&v);
            if (SUCCEEDED(ep->OpenPropertyStore(STGM_READ, &ps)))
            {
                if (SUCCEEDED(ps->GetValue(PKEY_Device_FriendlyName, &v)) && v.pwszVal)
                {
                    if (ToUpper(v.pwszVal).find(nameKey) != std::wstring::npos)
                        ok = true;
                }
                PropVariantClear(&v);
                ps->Release();
            }
        }

        // 都没给筛选条件时，取第一项
        if (!ok && key.empty() && nameKey.empty())
            ok = true;

        if (ok)
        {
            if (pIfPath)
            {
                std::wstring ifp;
                GetInterfacePath(ep, ifp);
                *pIfPath = ifp;
            }
            *ppDev = ep;
            col->Release();
            e->Release();
            return S_OK;
        }
        ep->Release();
    }

    col->Release();
    e->Release();
    return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
}

// 取 IKsControl：三条路径，能拿到任一即返回
static HRESULT GetKsControl(IMMDevice *dev, IKsControl **ppKs)
{
    *ppKs = nullptr;
    HRESULT hr = S_OK;

    // 1) IAudioClient 路径
    {
        IAudioClient *ac = nullptr;
        WAVEFORMATEX *mix = nullptr;
        hr = dev->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, (void **)&ac);
        if (SUCCEEDED(hr))
        {
            hr = ac->GetMixFormat(&mix);
            if (SUCCEEDED(hr) && mix)
            {
                hr = ac->Initialize(AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_NOPERSIST,
                                    10'000'000, 0, mix, nullptr);
            }
            else
            {
                WAVEFORMATEX fmt = {};
                fmt.wFormatTag = WAVE_FORMAT_PCM;
                fmt.nChannels = 1;
                fmt.nSamplesPerSec = 16000;
                fmt.wBitsPerSample = 16;
                fmt.nBlockAlign = (fmt.nChannels * fmt.wBitsPerSample) / 8;
                fmt.nAvgBytesPerSec = fmt.nSamplesPerSec * fmt.nBlockAlign;
                hr = ac->Initialize(AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_NOPERSIST,
                                    10'000'000, 0, &fmt, nullptr);
            }
            if (SUCCEEDED(hr))
            {
                hr = ac->GetService(__uuidof(IKsControl), (void **)ppKs);
                if (SUCCEEDED(hr))
                {
                    if (mix)
                        CoTaskMemFree(mix);
                    ac->Release();
                    wprintf(L"[+] IKsControl via IAudioClient\n");
                    return S_OK;
                }
            }
            if (mix)
                CoTaskMemFree(mix);
            if (ac)
                ac->Release();
        }
    }

    // 2) DeviceTopology：在“设备侧 Part”上 Activate(IKsControl)
    {
        IDeviceTopology *topo = nullptr;
        hr = dev->Activate(__uuidof(IDeviceTopology), CLSCTX_ALL, nullptr, (void **)&topo);
        if (SUCCEEDED(hr))
        {
            UINT cc = 0;
            hr = topo->GetConnectorCount(&cc);
            if (SUCCEEDED(hr) && cc > 0)
            {
                IConnector *epConn = nullptr;
                if (SUCCEEDED(topo->GetConnector(0, &epConn)))
                {
                    IConnector *devConn = nullptr;
                    if (SUCCEEDED(epConn->GetConnectedTo(&devConn)))
                    {
                        IPart *part = nullptr;
                        if (SUCCEEDED(devConn->QueryInterface(__uuidof(IPart), (void **)&part)))
                        {
                            // 关键：很多设备只在 Part 上能 Activate 到 IKsControl
                            hr = part->Activate(CLSCTX_ALL, __uuidof(IKsControl), (void **)ppKs);
                            part->Release();
                            if (SUCCEEDED(hr) && *ppKs)
                            {
                                devConn->Release();
                                epConn->Release();
                                topo->Release();
                                wprintf(L"[+] IKsControl via DeviceTopology Part::Activate\n");
                                return S_OK;
                            }
                        }
                        // 3) 再在设备侧 Connector 上尝试 QI(IKsControl) 兜底
                        if (FAILED(hr))
                        {
                            hr = devConn->QueryInterface(__uuidof(IKsControl), (void **)ppKs);
                            if (SUCCEEDED(hr) && *ppKs)
                            {
                                devConn->Release();
                                epConn->Release();
                                topo->Release();
                                wprintf(L"[+] IKsControl via DeviceTopology Connector::QI\n");
                                return S_OK;
                            }
                        }
                        devConn->Release();
                    }
                    epConn->Release();
                }
            }
            topo->Release();
        }
    }

    return E_NOINTERFACE;
}

// ---- KsProperty SET/GET ----
template <typename T>
static HRESULT KsSet(IKsControl *ks, ULONG pid, const T &data)
{
    KSPROPERTY prop = {};
    prop.Set = MYCOMPANY_APO_PROPSETID;
    prop.Id = pid;
    prop.Flags = KSPROPERTY_TYPE_SET;
    ULONG ret = 0;
    return ks->KsProperty(&prop, sizeof(prop), (PVOID)&data, sizeof(T), &ret);
}
template <typename T>
static HRESULT KsGet(IKsControl *ks, ULONG pid, T &out)
{
    KSPROPERTY prop = {};
    prop.Set = MYCOMPANY_APO_PROPSETID;
    prop.Id = pid;
    prop.Flags = KSPROPERTY_TYPE_GET;
    ULONG ret = 0;
    return ks->KsProperty(&prop, sizeof(prop), (PVOID)&out, sizeof(T), &ret);
}

static bool HexToBytes(const std::wstring &hex, std::vector<BYTE> &out)
{
    if (hex.size() % 2)
        return false;
    out.clear();
    out.reserve(hex.size() / 2);
    for (size_t i = 0; i < hex.size(); i += 2)
    {
        wchar_t b[3] = {hex[i], hex[i + 1], 0};
        out.push_back((BYTE)wcstol(b, nullptr, 16));
    }
    return true;
}

// ---- 选项解析 ----
enum class FlowSel
{
    Render,
    Capture
};
struct Options
{
    FlowSel flow = FlowSel::Render;       // --render / --capture
    std::wstring pnp = kDefaultPnPSubstr; // --pnp "<substr>"（hwid 或 接口路径子串都行）
    std::wstring name;                    // --name "<substr>"
    int index = -1;                       // --index N
    int argi = 1;                         // 第一个非选项参数的下标
    bool verbose = false;                 // --verbose
    bool forceStream = false;
    DWORD forceMs = 150;
};
static Options Parse(int argc, wchar_t **argv)
{
    Options o;
    while (o.argi < argc && wcsncmp(argv[o.argi], L"--", 2) == 0)
    {
        std::wstring s = argv[o.argi];
        if (s == L"--capture")
            o.flow = FlowSel::Capture;
        else if (s == L"--render")
            o.flow = FlowSel::Render;
        else if (s == L"--pnp" && o.argi + 1 < argc)
            o.pnp = argv[++o.argi];
        else if (s == L"--name" && o.argi + 1 < argc)
            o.name = argv[++o.argi];
        else if (s == L"--index" && o.argi + 1 < argc)
            o.index = _wtoi(argv[++o.argi]);
        else if (s == L"--verbose")
            o.verbose = true;
        else if (s == L"--force-stream")
        {
            o.forceStream = true;
            if (o.argi + 1 < argc && iswdigit(argv[o.argi + 1][0]))
                o.forceMs = (DWORD)_wtoi(argv[++o.argi]);
        }
        else
            break;
        ++o.argi;
    }
    return o;
}
// 拉流函数
static HRESULT EnsureApoLoaded(IMMDevice *dev, DWORD runMs)
{
    IAudioClient *ac = nullptr;
    HRESULT hr = dev->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, (void **)&ac);
    if (FAILED(hr))
        return hr;

    WAVEFORMATEX *mix = nullptr;
    if (SUCCEEDED(ac->GetMixFormat(&mix)) && mix)
    {
        hr = ac->Initialize(AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_NOPERSIST, 10'000'000, 0, mix, nullptr);
        CoTaskMemFree(mix);
    }
    else
    {
        WAVEFORMATEX fmt{};
        fmt.wFormatTag = WAVE_FORMAT_PCM;
        fmt.nChannels = 2;
        fmt.nSamplesPerSec = 48000;
        fmt.wBitsPerSample = 16;
        fmt.nBlockAlign = (fmt.nChannels * fmt.wBitsPerSample) / 8;
        fmt.nAvgBytesPerSec = fmt.nSamplesPerSec * fmt.nBlockAlign;
        hr = ac->Initialize(AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_NOPERSIST, 10'000'000, 0, &fmt, nullptr);
    }
    if (FAILED(hr))
    {
        ac->Release();
        return hr;
    }

    ac->Start();
    Sleep(runMs);
    ac->Stop();
    ac->Release();
    return S_OK;
}
int wmain(int argc, wchar_t **argv)
{
    // 提前声明会在 goto done 之后析构的对象，避免 C2362
    std::wstring ifPath; // 供 FindEndpoint 返回接口路径
    IMMDevice *dev = nullptr;
    IKsControl *ks = nullptr;

    if (argc < 2)
    {
        wprintf(L"用法:\n");
        wprintf(L"  ApoCtl.exe [--render|--capture] [--pnp <substr>|--name <substr>|--index N] list\n");
        wprintf(L"  ApoCtl.exe [选择器...] gain <linear>\n");
        wprintf(L"  ApoCtl.exe [选择器...] eq <bandIndex> <linear>\n");
        wprintf(L"  ApoCtl.exe [选择器...] reverb <mix0..1>\n");
        wprintf(L"  ApoCtl.exe [选择器...] limiter <thresLinear>\n");
        wprintf(L"  ApoCtl.exe [选择器...] blob <hex_no_spaces>\n");
        wprintf(L"  ApoCtl.exe [选择器...] get gain|reverb|limiter\n");
        wprintf(L"  例：ApoCtl.exe --render --pnp \"USB\\VID_0A67&PID_30A2&MI_00\" gain 0.5\n");
        return 0; // 这里直接 return，避免 goto 跳过构造
    }

    HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(hr))
    {
        HR_FAIL(hr, "CoInitializeEx");
        return 1;
    }

    Options opt = Parse(argc, argv);
    if (opt.argi >= argc)
    {
        wprintf(L"[!] 缺少命令；使用 list/gain/eq/... \n");
        CoUninitialize();
        return 0;
    }

    // 列出端点并退出（直接 return，避免 goto）
    if (_wcsicmp(argv[opt.argi], L"list") == 0)
    {
        ListEndpoints(opt.flow == FlowSel::Render ? Flow::Render : Flow::Capture);
        CoUninitialize();
        return 0;
    }

    // 按选择器找端点
    hr = FindEndpoint(opt.flow == FlowSel::Render ? Flow::Render : Flow::Capture,
                      opt.pnp, opt.name, opt.index, &dev, &ifPath);
    RETURN_IF_FAILED(hr, "Find endpoint (按 pnp/name/index)");

    if (opt.verbose)
    {
        std::wstring hwid;
        GetHardwareIdFromInterfacePath(ifPath, hwid);
        wprintf(L"[v] Chosen ifPath: %s\n[v] Chosen hwid: %s\n",
                ifPath.c_str(), hwid.empty() ? L"(n/a)" : hwid.c_str());
    }

    // 取 IKsControl（多路径）
    if (opt.forceStream)
    {
        HRESULT hrfs = EnsureApoLoaded(dev, opt.forceMs);
        if (FAILED(hrfs))
            wprintf(L"[!] force-stream failed (0x%08X)\n", hrfs);
    }
    hr = GetKsControl(dev, &ks);
    RETURN_IF_FAILED(hr, "Get IKsControl");

    { // 命令处理放内层作用域，避免 goto 跨越初始化
        std::wstring cmd = argv[opt.argi];
        std::transform(cmd.begin(), cmd.end(), cmd.begin(), ::towlower);

        if (cmd == L"gain" && opt.argi + 1 < argc)
        {
            float g = (float)_wtof(argv[opt.argi + 1]);
            hr = KsSet(ks, PID_Gain, g);
            HR_OK(hr) ? wprintf(L"[OK] Set Gain = %f\n", g) : HR_FAIL(hr, "Set Gain");
        }
        else if (cmd == L"eq" && opt.argi + 2 < argc)
        {
            EQBandParam p{(LONG)_wtol(argv[opt.argi + 1]), (float)_wtof(argv[opt.argi + 2])};
            hr = KsSet(ks, PID_EQBand, p);
            HR_OK(hr) ? wprintf(L"[OK] Set EQ band %ld -> %f\n", p.bandIndex, p.gainLinear)
                      : HR_FAIL(hr, "Set EQ");
        }
        else if (cmd == L"reverb" && opt.argi + 1 < argc)
        {
            float mix = (float)_wtof(argv[opt.argi + 1]);
            hr = KsSet(ks, PID_Reverb, mix);
            HR_OK(hr) ? wprintf(L"[OK] Set Reverb mix = %f\n", mix) : HR_FAIL(hr, "Set Reverb");
        }
        else if (cmd == L"limiter" && opt.argi + 1 < argc)
        {
            float thr = (float)_wtof(argv[opt.argi + 1]);
            hr = KsSet(ks, PID_Limiter, thr);
            HR_OK(hr) ? wprintf(L"[OK] Set Limiter threshold = %f\n", thr)
                      : HR_FAIL(hr, "Set Limiter");
        }
        else if (cmd == L"blob" && opt.argi + 1 < argc)
        {
            std::vector<BYTE> bytes;
            if (!HexToBytes(argv[opt.argi + 1], bytes))
            {
                wprintf(L"[!] blob 需要偶数字节的16进制串（无空格）\n");
                goto done;
            }
            KSPROPERTY prop = {};
            prop.Set = MYCOMPANY_APO_PROPSETID;
            prop.Id = PID_ParamsBlob;
            prop.Flags = KSPROPERTY_TYPE_SET;
            ULONG ret = 0;
            hr = ks->KsProperty(&prop, sizeof(prop), bytes.data(), (ULONG)bytes.size(), &ret);
            HR_OK(hr) ? wprintf(L"[OK] Set Blob (%u bytes)\n", (UINT)bytes.size())
                      : HR_FAIL(hr, "Set Blob");
        }
        else if (cmd == L"get" && opt.argi + 1 < argc)
        {
            std::wstring which = argv[opt.argi + 1];
            std::transform(which.begin(), which.end(), which.begin(), ::towlower);
            if (which == L"gain")
            {
                float g = 0;
                hr = KsGet(ks, PID_Gain, g);
                HR_OK(hr) ? wprintf(L"[OK] Gain=%f\n", g) : HR_FAIL(hr, "Get Gain");
            }
            else if (which == L"reverb")
            {
                float m = 0;
                hr = KsGet(ks, PID_Reverb, m);
                HR_OK(hr) ? wprintf(L"[OK] Reverb=%f\n", m) : HR_FAIL(hr, "Get Reverb");
            }
            else if (which == L"limiter")
            {
                float t = 0;
                hr = KsGet(ks, PID_Limiter, t);
                HR_OK(hr) ? wprintf(L"[OK] Limiter=%f\n", t) : HR_FAIL(hr, "Get Limiter");
            }
            else
                wprintf(L"[!] 未实现 get %s\n", which.c_str());
        }
        else
        {
            wprintf(L"[!] 命令/参数不完整；可先用 list 确认端点\n");
        }
    }

done:
    if (dev)
        dev->Release();
    if (ks)
        ks->Release();
    CoUninitialize();
    return 0;
}
