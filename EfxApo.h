// EfxApo.h
#pragma once

// Windows & COM
#include <windows.h>
#include <unknwn.h>
#include <mmreg.h>

// Property Store
#include <propsys.h>
#include <propvarutil.h>
#include <functiondiscoverykeys_devpkey.h>

// APO 基类
#include <audioenginebaseapo.h>    // 定义 APOInit / APO_CONNECTION_*

// 如果你的环境没有 extension 头，这行删掉也可以
// #include <audioengineextensionapo.h>

// 你的 GUID/参数与 DSP
#include "MyApoGuids.h"
#include "MyApoParams.h"
#include "dsp_wrapper.h"

// —— 为了避免 IntelliSense 报红，再做一份前向声明（只用到指针时是安全的）——
struct APOInit;
struct APO_CONNECTION_DESCRIPTOR;
struct APO_CONNECTION_PROPERTY;

class CMyCompanyEfxApo :
    public IAudioProcessingObjectRT,
    public IAudioSystemEffects,
    public IPropertyStore
{
public:
    CMyCompanyEfxApo();
    ~CMyCompanyEfxApo();

    // IUnknown
    STDMETHOD(QueryInterface)(REFIID riid, void** ppv);
    STDMETHOD_(ULONG, AddRef)();
    STDMETHOD_(ULONG, Release)();

    // IAudioProcessingObject
    STDMETHOD(Initialize)(_In_ APOInit* pAPOInit);   // ★ 不是 APO_INIT_PARAM*
    STDMETHOD(LockForProcess)(
        UINT32 u32NumInputConnections,
        _In_reads_(u32NumInputConnections) APO_CONNECTION_DESCRIPTOR** ppInputConnectionDescriptors,
        UINT32 u32NumOutputConnections,
        _In_reads_(u32NumOutputConnections) APO_CONNECTION_DESCRIPTOR** ppOutputConnectionDescriptors);
    STDMETHOD(UnlockForProcess)();

    // IAudioProcessingObjectRT —— ★ 返回 void，参数是**双重指针**
    STDMETHOD_(void, APOProcess)(
        UINT32 u32NumInputConnections,
        _Inout_updates_(u32NumInputConnections) APO_CONNECTION_PROPERTY** ppInputConnections,
        UINT32 u32NumOutputConnections,
        _Inout_updates_(u32NumOutputConnections) APO_CONNECTION_PROPERTY** ppOutputConnections);

    // IPropertyStore
    STDMETHOD(GetCount)(DWORD* cProps);
    STDMETHOD(GetAt)(DWORD iProp, PROPERTYKEY* pkey);
    STDMETHOD(GetValue)(REFPROPERTYKEY key, PROPVARIANT* pv);
    STDMETHOD(SetValue)(REFPROPERTYKEY key, REFPROPVARIANT propvar);
    STDMETHOD(Commit)();

private:
    void*   m_dspCtx = nullptr;
    UINT32  m_sr = 48000;
    UINT32  m_ch = 2;

    MyDspParams   m_paramsActive{};
    MyDspParams   m_paramsPending{};
    volatile LONG m_paramsSeq = 0;

    HANDLE  m_hPipeThread = nullptr;
    HANDLE  m_hStopEvt    = nullptr;
    static DWORD WINAPI PipeThreadMain(LPVOID self);
    void ApplyParams_NoLock(const MyDspParams& p);
};
