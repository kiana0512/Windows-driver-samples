// EfxApo.cpp（关键片段）
#include "EfxApo.h"
#include <functiondiscoverykeys_devpkey.h>
#include <propvarutil.h>
#include <vector>

// === 工具：把 blob/简单值写入 pending，再 bump seq ===
static void CopyParams(MyDspParams& dst, const MyDspParams& src) { memcpy(&dst,&src,sizeof(dst)); }
void CMyCompanyEfxApo::ApplyParams_NoLock(const MyDspParams& p) {
    // 1) 落到 dsp_wrapper
    dsp_set_gain(m_dspCtx, p.gain);
    for (int i=0;i<MY_EQ_BANDS;i++) {
        auto& b = p.eq[i];
        int bandType = b.type; // 0/1/2：你可在 dsp_wrapper 里做分派
        dsp_set_eq_enabled(m_dspCtx, i, b.enabled);
        dsp_set_eq_params (m_dspCtx, i, b.freq, b.q, b.gain_db);
    }
    if (p.reverb.enabled) {
        dsp_set_reverb_enabled(m_dspCtx, 1);
        dsp_set_reverb_params (m_dspCtx, p.reverb.wet, p.reverb.room, p.reverb.damp, p.reverb.pre_ms);
    } else {
        dsp_set_reverb_enabled(m_dspCtx, 0);
    }
    dsp_set_limiter_enabled(m_dspCtx, p.limiterEnabled);

    // 2) 预留：如果有 opcode，就交给你的解释器/JIT
    // if (p.opcodeSize > 0) MyOpcodeExec::Load(p.opcode, p.opcodeSize);
}

// === Initialize：保存采样率/通道、创建 ctx、启动管道线程 ===
STDMETHODIMP CMyCompanyEfxApo::Initialize(APOInit* pAPOInit) {
    // ……你的原有初始化……
    // 取当前格式（示例：48k/2ch；真实从 conn props / APOInitSystemEffects2 获取）
    m_sr = 48000; m_ch = 2;
    m_dspCtx = dsp_create_context(m_sr, m_ch);

    // 默认参数
    ZeroMemory(&m_paramsActive, sizeof(m_paramsActive));
    m_paramsActive.gain = 1.0f;
    m_paramsActive.limiterEnabled = 1;
    CopyParams(m_paramsPending, m_paramsActive);

    // 启动命名管道监听（方案C预留：立即生效通道）
    m_hStopEvt = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    m_hPipeThread = CreateThread(nullptr, 0, &CMyCompanyEfxApo::PipeThreadMain, this, 0, nullptr);
    return S_OK;
}

// === 实时处理：检查参数是否更新，然后处理 ===
// 新（注意返回类型 void、参数是双重指针 **）：
STDMETHODIMP_(void) CMyCompanyEfxApo::APOProcess(
    UINT32 inC, APO_CONNECTION_PROPERTY** inP,
    UINT32 outC, APO_CONNECTION_PROPERTY** outP)
{
    if (inC && outC && inP && outP && inP[0] && outP[0]) {
        float* in  = (float*)inP[0]->pBuffer;
        float* out = (float*)outP[0]->pBuffer;
        UINT32 frames = outP[0]->u32ValidFrameCount;
        dsp_process_block(m_dspCtx, in, out, frames, m_ch);
    }
}

// === IPropertyStore：应用通过 IMMDevice→IPropertyStore→SetValue 写到我们 ===
STDMETHODIMP CMyCompanyEfxApo::SetValue(REFPROPERTYKEY key, REFPROPVARIANT pv) {
    if (IsEqualPropertyKey(key, PKEY_MyCompany_ParamsBlob)) {
        if (pv.vt != VT_BLOB || pv.blob.cbSize != sizeof(MyDspParams)) return E_INVALIDARG;
        const MyDspParams* p = (const MyDspParams*)pv.blob.pBlobData;
        CopyParams(m_paramsPending, *p);
        InterlockedIncrement(&m_paramsSeq); // 通知处理线程
        return S_OK;
    }
    if (IsEqualPropertyKey(key, PKEY_MyCompany_Gain)) {
        if (pv.vt == VT_R4 || pv.vt == VT_R8) {
            m_paramsPending.gain = (pv.vt==VT_R4) ? pv.fltVal : (float)pv.dblVal;
            InterlockedIncrement(&m_paramsSeq);
            return S_OK;
        }
        return E_INVALIDARG;
    }
    // 其他 key（EQ/Reverb/Limiter）同理：解析 BLOB 或标量，写 m_paramsPending + bump seq
    return E_NOTIMPL;
}

STDMETHODIMP CMyCompanyEfxApo::GetCount(DWORD* c){ if(!c) return E_POINTER; *c=1; return S_OK; }
STDMETHODIMP CMyCompanyEfxApo::GetAt(DWORD i, PROPERTYKEY* k){ if(!k) return E_POINTER; if(i) return E_BOUNDS; *k=PKEY_MyCompany_ParamsBlob; return S_OK; }
STDMETHODIMP CMyCompanyEfxApo::GetValue(REFPROPERTYKEY key, PROPVARIANT* pv){ return E_NOTIMPL; }
STDMETHODIMP CMyCompanyEfxApo::Commit(){ return S_OK; }

// === 方案C：后台线程，监听命名管道，收 MyDspParams 结构体 ===
DWORD WINAPI CMyCompanyEfxApo::PipeThreadMain(LPVOID selfPtr) {
    auto self = (CMyCompanyEfxApo*)selfPtr;
    for (;;) {
        HANDLE hPipe = CreateNamedPipeW(MYCOMPANY_PIPE_NAME,
            PIPE_ACCESS_INBOUND, PIPE_TYPE_BYTE|PIPE_READMODE_BYTE|PIPE_WAIT,
            1, 0, sizeof(MyDspParams), 0, nullptr);
        if (hPipe == INVALID_HANDLE_VALUE) break;

        BOOL ok = ConnectNamedPipe(hPipe, nullptr) ? TRUE : (GetLastError() == ERROR_PIPE_CONNECTED);
        if (!ok) { CloseHandle(hPipe); break; }

        MyDspParams p{};
        DWORD got = 0;
        while (ReadFile(hPipe, &p, sizeof(p), &got, nullptr)) {
            if (got == sizeof(p)) {
                CopyParams(self->m_paramsPending, p);
                InterlockedIncrement(&self->m_paramsSeq);
            }
            got = 0;
            if (WaitForSingleObject(self->m_hStopEvt, 0) == WAIT_OBJECT_0) break;
        }
        DisconnectNamedPipe(hPipe);
        CloseHandle(hPipe);

        if (WaitForSingleObject(self->m_hStopEvt, 0) == WAIT_OBJECT_0) break;
    }
    return 0;
}
