// EfxApo.cpp —— 最小可编译/可实例化骨架（与 EfxApo.h 对齐，旧 SDK 签名）
// 变更点（相对你原始版本）：
// 1) QueryInterface 增加 IAudioSystemEffects “字面量 IID” 兜底，避免不同 SDK 头导致 IID 不一致而返回 E_NOINTERFACE。
// 2) GetEffectsList 修正：返回“效果 GUID”（你的 APO CLSID），而不是处理模式 GUID（DEFAULT 模式由注册表 FX\0/PM7 告知）。
// 3) 加入 DbgLog 输出，在 Initialize / LockForProcess / APOProcess 打点，调试时用 DebugView.exe 观察。
// 4) APOProcess 做一个 20% 的线性衰减，便于立刻用耳朵验证效果链是否生效。

#include "EfxApo.h"
#include "MyApoGuids.h"      // 声明 CLSID_MyCompanyEfxApo（你工程已有的 Guids 声明/定义）
#include <audioclient.h>
#include <mmreg.h>
#include <new>               // std::nothrow
#include <cstring>           // memcpy
#include <strsafe.h>         // DbgLog 安全格式化

// 旧接口 IID 的字面量常量（防止不同 SDK 头导致 __uuidof(IAudioSystemEffects) 的 GUID 不一致）
static const GUID IID_IAudioSystemEffects_Legacy =
    {0xB61C2C5F, 0x31A8, 0x49CB, {0xAF, 0xA5, 0xF1, 0xF1, 0x0E, 0xB3, 0xC1, 0xDC}};

// 某些环境未定义默认处理模式 GUID，兜底定义（只是为了本文件内引用，不会影响系统）
#ifndef INITGUID
#define INITGUID
#endif
#ifndef AUDIO_SIGNALPROCESSINGMODE_DEFAULT
// {C18E2F7E-933D-4965-B7D1-1EEF228D2AF3}
DEFINE_GUID(AUDIO_SIGNALPROCESSINGMODE_DEFAULT,
            0xC18E2F7E, 0x933D, 0x4965, 0xB7, 0xD1, 0x1E, 0xEF, 0x22, 0x8D, 0x2A, 0xF3);
#endif

// ======= 简单日志辅助（用 DebugView.exe 观察 OutputDebugString） =======
static void DbgLog(const wchar_t* fmt, ...)
{
    wchar_t buf[256];
    va_list ap; va_start(ap, fmt);
    StringCchVPrintfW(buf, _countof(buf), fmt, ap);
    va_end(ap);
    OutputDebugStringW(buf);
    OutputDebugStringW(L"\n");
}

// ====== 构造 / 析构 ======
CMyCompanyEfxApo::CMyCompanyEfxApo() {}
CMyCompanyEfxApo::~CMyCompanyEfxApo()
{
    if (m_hStopEvt) { SetEvent(m_hStopEvt); }
    if (m_hPipeThread)
    {
        WaitForSingleObject(m_hPipeThread, 2000);
        CloseHandle(m_hPipeThread);
        m_hPipeThread = nullptr;
    }
    if (m_hStopEvt) { CloseHandle(m_hStopEvt); m_hStopEvt = nullptr; }
    // TODO: if (m_dspCtx) { /* dsp_destroy(m_dspCtx); */ m_dspCtx = nullptr; }
}

// ====================== IUnknown ======================
STDMETHODIMP CMyCompanyEfxApo::QueryInterface(REFIID riid, void **ppv)
{
    if (!ppv) return E_POINTER;
    *ppv = nullptr;

    if (riid == __uuidof(IUnknown)) {
        // 多继承下到 IUnknown 的二义性：选一个父接口子对象返回（APO 基类通常是 IAudioProcessingObject）
        *ppv = static_cast<IAudioProcessingObject *>(this);
    }
    else if (riid == __uuidof(IAudioProcessingObject)) {
        *ppv = static_cast<IAudioProcessingObject *>(this);
    }
    else if (riid == __uuidof(IAudioProcessingObjectRT)) {
        *ppv = static_cast<IAudioProcessingObjectRT *>(this);
    }
    else if (riid == __uuidof(IAudioSystemEffects) || riid == IID_IAudioSystemEffects_Legacy) {
        // ★关键：加入字面量 IID 兜底，避免与测试程序/系统不同头文件导致的 IID 偏差
        *ppv = static_cast<IAudioSystemEffects *>(this);
    }
    else if (riid == __uuidof(IPropertyStore)) {
        *ppv = static_cast<IPropertyStore *>(this);
    }
    else {
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}
ULONG STDMETHODCALLTYPE CMyCompanyEfxApo::AddRef() { return ++m_ref; }
ULONG STDMETHODCALLTYPE CMyCompanyEfxApo::Release()
{
    ULONG n = --m_ref;
    if (!n) delete this;
    return n;
}

// ============= IAudioProcessingObject（最小实现） =============
STDMETHODIMP CMyCompanyEfxApo::Initialize(UINT32 cbDataSize, BYTE *pbyData)
{
    // 你的 SDK 是旧签名：Initialize(cbDataSize, pbyData)
    UNREFERENCED_PARAMETER(cbDataSize);
    UNREFERENCED_PARAMETER(pbyData);

    // TODO: 若使用 APOINIT_SYSTEMEFFECTS 等，这里解析 pbyData
    m_sr = 48000;
    m_ch = 2;

    // m_dspCtx = dsp_create_context(m_sr, m_ch);

    ZeroMemory(&m_paramsActive, sizeof(m_paramsActive));
    ZeroMemory(&m_paramsPending, sizeof(m_paramsPending));

    m_hStopEvt    = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    m_hPipeThread = CreateThread(nullptr, 0, &CMyCompanyEfxApo::PipeThreadMain, this, 0, nullptr);

    DbgLog(L"[MyAPO] Initialize: sr=%u ch=%u", m_sr, m_ch);
    return S_OK;
}

STDMETHODIMP CMyCompanyEfxApo::LockForProcess(
    UINT32 inCount, _In_reads_(inCount) APO_CONNECTION_DESCRIPTOR **inDesc,
    UINT32 outCount, _In_reads_(outCount) APO_CONNECTION_DESCRIPTOR **outDesc)
{
    // 最小实现：从第一个输出连接的格式推导采样率/通道
    if (outCount && outDesc && outDesc[0] && outDesc[0]->pFormat)
    {
        IAudioMediaType *pType = outDesc[0]->pFormat;
        auto pWfx = reinterpret_cast<const WAVEFORMATEX *>(pType->GetAudioFormat());
        if (pWfx) {
            m_sr = pWfx->nSamplesPerSec;
            m_ch = pWfx->nChannels;
            DbgLog(L"[MyAPO] LockForProcess: sr=%u ch=%u", m_sr, m_ch);
        }
    }
    UNREFERENCED_PARAMETER(inCount);
    UNREFERENCED_PARAMETER(inDesc);
    return S_OK;
}

STDMETHODIMP CMyCompanyEfxApo::UnlockForProcess() { return S_OK; }

STDMETHODIMP CMyCompanyEfxApo::GetLatency(_Out_ HNSTIME *pLatency)
{
    if (!pLatency) return E_POINTER;
    *pLatency = 0; // 最小实现：报告 0；实际可按滤波器组延迟换算
    return S_OK;
}
STDMETHODIMP CMyCompanyEfxApo::Reset() { return S_OK; }

STDMETHODIMP CMyCompanyEfxApo::GetRegistrationProperties(
    _Outptr_result_maybenull_ APO_REG_PROPERTIES **ppRegProps)
{
    // TODO：如需在注册层面暴露更多属性（类名、友好名、UAPO 标志等），在这里返回结构体。
    if (ppRegProps) *ppRegProps = nullptr;
    return E_NOTIMPL;
}

STDMETHODIMP CMyCompanyEfxApo::IsInputFormatSupported(
    _In_opt_ IAudioMediaType * /*pOutputFormat*/,
    _In_ IAudioMediaType * /*pRequestedInputFormat*/,
    _Outptr_ IAudioMediaType **ppSupportedInputFormat)
{
    if (ppSupportedInputFormat) *ppSupportedInputFormat = nullptr;
    return E_NOTIMPL; // 最小实现先不做格式协商（Shared 模式下通常走引擎混音格式）
}

STDMETHODIMP CMyCompanyEfxApo::IsOutputFormatSupported(
    _In_opt_ IAudioMediaType * /*pInputFormat*/,
    _In_ IAudioMediaType * /*pRequestedOutputFormat*/,
    _Outptr_ IAudioMediaType **ppSupportedOutputFormat)
{
    if (ppSupportedOutputFormat) *ppSupportedOutputFormat = nullptr;
    return E_NOTIMPL;
}

STDMETHODIMP CMyCompanyEfxApo::GetInputChannelCount(_Out_ UINT32 *pu32ChannelCount)
{
    if (!pu32ChannelCount) return E_POINTER;
    *pu32ChannelCount = m_ch;
    return S_OK;
}
STDMETHODIMP CMyCompanyEfxApo::GetOutputChannelCount(_Out_ UINT32 *pu32ChannelCount)
{
    if (!pu32ChannelCount) return E_POINTER;
    *pu32ChannelCount = m_ch;
    return S_OK;
}

// ============= IAudioProcessingObjectRT（最小实现） =============
STDMETHODIMP_(void)
CMyCompanyEfxApo::APOProcess(
    UINT32 inC,  _Inout_updates_(inC)  APO_CONNECTION_PROPERTY **inP,
    UINT32 outC, _Inout_updates_(outC) APO_CONNECTION_PROPERTY **outP)
{
    // 最小可听效果：把输出整体衰减到 20%（约 -14 dB），便于“耳朵验证”
    if (!outC || !outP || !outP[0] || !outP[0]->pBuffer) return;

    const UINT32 frames = outP[0]->u32ValidFrameCount;
    if (frames == 0 || m_ch == 0) return;

    // 假定混音格式为 float32 interleaved（WASAPI 引擎内部常见）
    float* out = reinterpret_cast<float*>(outP[0]->pBuffer);
    const float kGain = 0.2f;
    const size_t samples = static_cast<size_t>(frames) * static_cast<size_t>(m_ch);
    for (size_t i = 0; i < samples; ++i) out[i] *= kGain;

    // 节流日志：每秒一条，避免刷屏
    static DWORD s_lastTick = 0;
    DWORD now = GetTickCount();
    if (now - s_lastTick > 1000) {
        s_lastTick = now;
        DbgLog(L"[MyAPO] APOProcess tick: frames=%u ch=%u", frames, m_ch);
    }

    UNREFERENCED_PARAMETER(inC);
    UNREFERENCED_PARAMETER(inP);
}

STDMETHODIMP_(UINT32)
CMyCompanyEfxApo::CalcInputFrames(UINT32 u32OutputFrameCount)
{
    // 线性 1:1。如做 resample/latency，这里需要换算。
    return u32OutputFrameCount;
}
STDMETHODIMP_(UINT32)
CMyCompanyEfxApo::CalcOutputFrames(UINT32 u32InputFrameCount)
{
    return u32InputFrameCount;
}

// ================ IAudioSystemEffects（v1） ================
STDMETHODIMP CMyCompanyEfxApo::GetEffectsList(
    _Outptr_result_buffer_(*pcEffects) GUID **ppEffects,
    _Out_ UINT *pcEffects,
    _Outptr_result_maybenull_ LPWSTR **ppName)
{
    // ★关键修正：返回“效果 GUID”（通常就是本 APO 的 CLSID），不是处理模式 GUID。
    if (!ppEffects || !pcEffects) return E_POINTER;

    *ppEffects = (GUID*)CoTaskMemAlloc(sizeof(GUID));
    if (!*ppEffects) return E_OUTOFMEMORY;

    **ppEffects = CLSID_MyCompanyEfxApo;  // <-- 由 MyApoGuids.cpp 提供的全局 CLSID
    *pcEffects = 1;

    if (ppName) *ppName = nullptr; // 可不返回友好名
    return S_OK;
}

// ===================== IPropertyStore（最小） =====================
STDMETHODIMP CMyCompanyEfxApo::GetCount(DWORD *cProps)
{
    if (!cProps) return E_POINTER;
    *cProps = 0;
    return S_OK;
}
STDMETHODIMP CMyCompanyEfxApo::GetAt(DWORD /*i*/, PROPERTYKEY * /*pkey*/) { return E_NOTIMPL; }
STDMETHODIMP CMyCompanyEfxApo::GetValue(REFPROPERTYKEY /*key*/, PROPVARIANT * /*pv*/) { return E_NOTIMPL; }
STDMETHODIMP CMyCompanyEfxApo::SetValue(REFPROPERTYKEY /*key*/, REFPROPVARIANT /*propvar*/) { return E_NOTIMPL; }
STDMETHODIMP CMyCompanyEfxApo::Commit() { return S_OK; }

// ======================= 内部辅助（占位） =======================
DWORD WINAPI CMyCompanyEfxApo::PipeThreadMain(LPVOID self)
{
    auto p = static_cast<CMyCompanyEfxApo *>(self);
    // TODO: 命名管道/共享内存等接收配置更新；收到后更新 m_paramsPending 并递增 m_paramsSeq
    if (p && p->m_hStopEvt) WaitForSingleObject(p->m_hStopEvt, INFINITE);
    return 0;
}
void CMyCompanyEfxApo::ApplyParams_NoLock(const MyDspParams & /*prm*/)
{
    // TODO: 把 pending 参数应用到 DSP；这里留空
}
