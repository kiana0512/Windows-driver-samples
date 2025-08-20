// EfxApo.cpp  —— 最小可编译/可实例化骨架（与 EfxApo.h 对齐，旧 SDK 签名）
#include "EfxApo.h"
#include <audioclient.h>
#include <mmreg.h>
#include <new>          // std::nothrow
#include <cstring>      // memcpy

// 某些环境未定义默认处理模式 GUID，兜底定义
#ifndef INITGUID
#define INITGUID
#endif
#ifndef AUDIO_SIGNALPROCESSINGMODE_DEFAULT
// {C18E2F7E-933D-4965-B7D1-1EEF228D2AF3}
DEFINE_GUID(AUDIO_SIGNALPROCESSINGMODE_DEFAULT,
            0xC18E2F7E,0x933D,0x4965,0xB7,0xD1,0x1E,0xEF,0x22,0x8D,0x2A,0xF3);
#endif

// ====== 构造 / 析构 ======
CMyCompanyEfxApo::CMyCompanyEfxApo() {}
CMyCompanyEfxApo::~CMyCompanyEfxApo() {
    if (m_hStopEvt)    { SetEvent(m_hStopEvt); }
    if (m_hPipeThread) { WaitForSingleObject(m_hPipeThread, 2000); CloseHandle(m_hPipeThread); }
    if (m_hStopEvt)    { CloseHandle(m_hStopEvt); m_hStopEvt = nullptr; }
    // TODO: if (m_dspCtx) { /* dsp_destroy(m_dspCtx); */ m_dspCtx = nullptr; }
}

// ====================== IUnknown ======================
STDMETHODIMP CMyCompanyEfxApo::QueryInterface(REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    *ppv = nullptr;

    if (riid == __uuidof(IUnknown)) {
        // 多继承下到 IUnknown 的二义性：选一个父接口子对象返回
        *ppv = static_cast<IAudioProcessingObject*>(this);
    }
    else if (riid == __uuidof(IAudioProcessingObject))   *ppv = static_cast<IAudioProcessingObject*>(this);
    else if (riid == __uuidof(IAudioProcessingObjectRT)) *ppv = static_cast<IAudioProcessingObjectRT*>(this);
    else if (riid == __uuidof(IAudioSystemEffects))      *ppv = static_cast<IAudioSystemEffects*>(this);
    else if (riid == __uuidof(IPropertyStore))           *ppv = static_cast<IPropertyStore*>(this);
    else return E_NOINTERFACE;

    AddRef();
    return S_OK;
}
ULONG STDMETHODCALLTYPE CMyCompanyEfxApo::AddRef()  { return ++m_ref; }
ULONG STDMETHODCALLTYPE CMyCompanyEfxApo::Release() { ULONG n = --m_ref; if (!n) delete this; return n; }

// ============= IAudioProcessingObject（最小实现） =============
STDMETHODIMP CMyCompanyEfxApo::Initialize(UINT32 cbDataSize, BYTE* pbyData) {
    // 你的 SDK 是旧签名：Initialize(cbDataSize, pbyData)
    UNREFERENCED_PARAMETER(cbDataSize);
    UNREFERENCED_PARAMETER(pbyData);

    // TODO: 解析 pbyData（若使用 APOINIT_SYSTEMEFFECTS 等）
    m_sr = 48000; 
    m_ch = 2;
    // m_dspCtx = dsp_create_context(m_sr, m_ch);

    ZeroMemory(&m_paramsActive,  sizeof(m_paramsActive));
    ZeroMemory(&m_paramsPending, sizeof(m_paramsPending));

    m_hStopEvt    = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    m_hPipeThread = CreateThread(nullptr, 0, &CMyCompanyEfxApo::PipeThreadMain, this, 0, nullptr);
    return S_OK;
}

STDMETHODIMP CMyCompanyEfxApo::LockForProcess(
    UINT32 inCount,  _In_reads_(inCount)  APO_CONNECTION_DESCRIPTOR** inDesc,
    UINT32 outCount, _In_reads_(outCount) APO_CONNECTION_DESCRIPTOR** outDesc)
{
    // 最小实现：从第一个输出连接的格式推导采样率/通道
    if (outCount && outDesc && outDesc[0] && outDesc[0]->pFormat) {
        IAudioMediaType* pType = outDesc[0]->pFormat;
        auto pWfx = reinterpret_cast<const WAVEFORMATEX*>(pType->GetAudioFormat());
        if (pWfx) { m_sr = pWfx->nSamplesPerSec; m_ch = pWfx->nChannels; }
    }
    UNREFERENCED_PARAMETER(inCount);
    UNREFERENCED_PARAMETER(inDesc);
    return S_OK;
}

STDMETHODIMP CMyCompanyEfxApo::UnlockForProcess() { return S_OK; }

STDMETHODIMP CMyCompanyEfxApo::GetLatency(_Out_ HNSTIME* pLatency) {
    if (!pLatency) return E_POINTER; 
    *pLatency = 0; 
    return S_OK;
}
STDMETHODIMP CMyCompanyEfxApo::Reset() { return S_OK; }

STDMETHODIMP CMyCompanyEfxApo::GetRegistrationProperties(
    _Outptr_result_maybenull_ APO_REG_PROPERTIES** ppRegProps)
{
    if (ppRegProps) *ppRegProps = nullptr;   // 后续需要再返回真正属性
    return E_NOTIMPL;
}

STDMETHODIMP CMyCompanyEfxApo::IsInputFormatSupported(
    _In_opt_ IAudioMediaType* /*pOutputFormat*/,
    _In_     IAudioMediaType* /*pRequestedInputFormat*/,
    _Outptr_ IAudioMediaType** ppSupportedInputFormat)
{
    if (ppSupportedInputFormat) *ppSupportedInputFormat = nullptr;
    return E_NOTIMPL;
}

STDMETHODIMP CMyCompanyEfxApo::IsOutputFormatSupported(
    _In_opt_ IAudioMediaType* /*pInputFormat*/,
    _In_     IAudioMediaType* /*pRequestedOutputFormat*/,
    _Outptr_ IAudioMediaType** ppSupportedOutputFormat)
{
    if (ppSupportedOutputFormat) *ppSupportedOutputFormat = nullptr;
    return E_NOTIMPL;
}

STDMETHODIMP CMyCompanyEfxApo::GetInputChannelCount(_Out_ UINT32* pu32ChannelCount) {
    if (!pu32ChannelCount) return E_POINTER; 
    *pu32ChannelCount = m_ch; 
    return S_OK;
}
STDMETHODIMP CMyCompanyEfxApo::GetOutputChannelCount(_Out_ UINT32* pu32ChannelCount) {
    if (!pu32ChannelCount) return E_POINTER; 
    *pu32ChannelCount = m_ch; 
    return S_OK;
}

// ============= IAudioProcessingObjectRT（最小实现） =============
STDMETHODIMP_(void) CMyCompanyEfxApo::APOProcess(
    UINT32 inC,  _Inout_updates_(inC)  APO_CONNECTION_PROPERTY** inP,
    UINT32 outC, _Inout_updates_(outC) APO_CONNECTION_PROPERTY** outP)
{
    UNREFERENCED_PARAMETER(inC);
    UNREFERENCED_PARAMETER(inP);
    UNREFERENCED_PARAMETER(outC);
    UNREFERENCED_PARAMETER(outP);
    // 最小实现：不改数据；后续把你的 DSP 放到这里
}

// 你这版 SDK：Calc*Frames 只有一个参数，返回值就是帧数
STDMETHODIMP_(UINT32) CMyCompanyEfxApo::CalcInputFrames(UINT32 u32OutputFrameCount) {
    // 线性 1:1。如做变采样，这里返回对应输入帧数。
    return u32OutputFrameCount;
}
STDMETHODIMP_(UINT32) CMyCompanyEfxApo::CalcOutputFrames(UINT32 u32InputFrameCount) {
    return u32InputFrameCount;
}

// ================ IAudioSystemEffects（v1） ================
STDMETHODIMP CMyCompanyEfxApo::GetEffectsList(
    _Outptr_result_buffer_(*pcEffects) GUID** ppEffects,
    _Out_ UINT* pcEffects,
    _Outptr_result_maybenull_ LPWSTR** ppName)
{
    if (!ppEffects || !pcEffects) return E_POINTER;
    *ppEffects = (GUID*)CoTaskMemAlloc(sizeof(GUID));
    if (!*ppEffects) return E_OUTOFMEMORY;
    **ppEffects = AUDIO_SIGNALPROCESSINGMODE_DEFAULT;
    *pcEffects  = 1;
    if (ppName) *ppName = nullptr;
    return S_OK;
}

// ===================== IPropertyStore（最小） =====================
STDMETHODIMP CMyCompanyEfxApo::GetCount(DWORD* cProps) {
    if (!cProps) return E_POINTER; 
    *cProps = 0; 
    return S_OK;
}
STDMETHODIMP CMyCompanyEfxApo::GetAt(DWORD /*i*/, PROPERTYKEY* /*pkey*/) { return E_NOTIMPL; }
STDMETHODIMP CMyCompanyEfxApo::GetValue(REFPROPERTYKEY /*key*/, PROPVARIANT* /*pv*/) { return E_NOTIMPL; }
STDMETHODIMP CMyCompanyEfxApo::SetValue(REFPROPERTYKEY /*key*/, REFPROPVARIANT /*propvar*/) { return E_NOTIMPL; }
STDMETHODIMP CMyCompanyEfxApo::Commit() { return S_OK; }

// ======================= 内部辅助（占位） =======================
DWORD WINAPI CMyCompanyEfxApo::PipeThreadMain(LPVOID self) {
    auto p = static_cast<CMyCompanyEfxApo*>(self);
    // TODO: 命名管道/共享内存等接收配置更新；收到后更新 m_paramsPending 并递增 m_paramsSeq
    // 简化：直接等待停止事件
    if (p && p->m_hStopEvt) WaitForSingleObject(p->m_hStopEvt, INFINITE);
    return 0;
}
void CMyCompanyEfxApo::ApplyParams_NoLock(const MyDspParams& /*p*/) {
    // TODO: 把 pending 参数应用到 DSP；这里留空
}
