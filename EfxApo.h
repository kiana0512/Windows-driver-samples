// EfxApo.h  —— 你的 EFX APO 主类（签名与 SDK 对齐）
#pragma once

// Windows & COM
#include <windows.h>
#include <unknwn.h>
#include <mmreg.h>

// Property Store
#include <propsys.h>
#include <propvarutil.h>
#include <functiondiscoverykeys_devpkey.h>

// APO 基类/类型
#include <audioenginebaseapo.h> // IAudioProcessingObject / RT 等
// #include <audioengineextensionapo.h>  // 没有可先注掉

// 你的 GUID/参数 与 DSP
#include "MyApoGuids.h"
#include "MyApoParams.h"
#include "dsp_wrapper.h"

#include <atomic> // 用到 std::atomic

// 前向声明（真实定义在 SDK 头里）
struct APOInit;
struct APO_CONNECTION_DESCRIPTOR;
struct APO_CONNECTION_PROPERTY;

class CMyCompanyEfxApo : public IAudioProcessingObject,   // 必须继承：否则 override 不匹配
                         public IAudioProcessingObjectRT, //  RT 接口：APOProcess/Calc*Frames
                         public IAudioSystemEffects,      //  EFX（v1）
                         public IPropertyStore
{
public:
    CMyCompanyEfxApo();
    ~CMyCompanyEfxApo();

    // IUnknown
    STDMETHOD(QueryInterface)(REFIID riid, void **ppv);
    STDMETHOD_(ULONG, AddRef)();
    STDMETHOD_(ULONG, Release)();

    // ===== IAudioProcessingObject =====
    STDMETHOD(Initialize)(_In_ UINT32 cbDataSize,
                          _In_reads_bytes_opt_(cbDataSize) BYTE *pbyData);
    STDMETHOD(LockForProcess)(
        UINT32 u32NumInputConnections,
        _In_reads_(u32NumInputConnections) APO_CONNECTION_DESCRIPTOR **ppInputConnectionDescriptors,
        UINT32 u32NumOutputConnections,
        _In_reads_(u32NumOutputConnections) APO_CONNECTION_DESCRIPTOR **ppOutputConnectionDescriptors);
    STDMETHOD(UnlockForProcess)();

    STDMETHOD(GetLatency)(_Out_ HNSTIME *pLatency);
    STDMETHOD(Reset)();
    STDMETHOD(GetRegistrationProperties)(_Outptr_result_maybenull_ APO_REG_PROPERTIES **ppRegProps);
    STDMETHOD(IsInputFormatSupported)(
        _In_opt_ IAudioMediaType *pOutputFormat,
        _In_ IAudioMediaType *pRequestedInputFormat,
        _Outptr_ IAudioMediaType **ppSupportedInputFormat);
    STDMETHOD(IsOutputFormatSupported)(
        _In_opt_ IAudioMediaType *pInputFormat,
        _In_ IAudioMediaType *pRequestedOutputFormat,
        _Outptr_ IAudioMediaType **ppSupportedOutputFormat);
    STDMETHOD(GetInputChannelCount)(_Out_ UINT32 *pu32ChannelCount);
    STDMETHOD(GetOutputChannelCount)(_Out_ UINT32 *pu32ChannelCount);

    // ===== IAudioProcessingObjectRT =====
    STDMETHOD_(void, APOProcess)(
        UINT32 u32NumInputConnections,
        _Inout_updates_(u32NumInputConnections) APO_CONNECTION_PROPERTY **ppInputConnections,
        UINT32 u32NumOutputConnections,
        _Inout_updates_(u32NumOutputConnections) APO_CONNECTION_PROPERTY **ppOutputConnections);

    STDMETHOD_(UINT32, CalcInputFrames)(UINT32 u32OutputFrameCount);
    STDMETHOD_(UINT32, CalcOutputFrames)(UINT32 u32InputFrameCount);

    // ===== IAudioSystemEffects (v1) =====
    // 你这套 SDK 常见签名：第三个参数为 LPWSTR**（名字数组，可为 null）
    STDMETHOD(GetEffectsList)(
        _Outptr_result_buffer_(*pcEffects) GUID **ppEffects,
        _Out_ UINT *pcEffects,
        _Outptr_result_maybenull_ LPWSTR **ppwstrEffectName);

    // ===== IPropertyStore =====
    STDMETHOD(GetCount)(DWORD *cProps);
    STDMETHOD(GetAt)(DWORD iProp, PROPERTYKEY *pkey);
    STDMETHOD(GetValue)(REFPROPERTYKEY key, PROPVARIANT *pv);
    STDMETHOD(SetValue)(REFPROPERTYKEY key, REFPROPVARIANT propvar);
    STDMETHOD(Commit)();

private:
    // 你的成员保持原样
    std::atomic<ULONG> m_ref{1};

    void *m_dspCtx = nullptr;
    UINT32 m_sr = 48000;
    UINT32 m_ch = 2;

    MyDspParams m_paramsActive{};
    MyDspParams m_paramsPending{};
    volatile LONG m_paramsSeq = 0;

    HANDLE m_hPipeThread = nullptr;
    HANDLE m_hStopEvt = nullptr;
    static DWORD WINAPI PipeThreadMain(LPVOID self);
    void ApplyParams_NoLock(const MyDspParams &p);
};
