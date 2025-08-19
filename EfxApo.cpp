#include "EfxApo.h"
#include "dsp_wrapper.h"
#include <cstring>

MyEfxAPORT::MyEfxAPORT() : m_ctx(nullptr), m_channels(2), m_sampleRate(48000), m_gainLocal(1.0f) {}

MyEfxAPORT::~MyEfxAPORT() {
    if (m_ctx) dsp_destroy_context(m_ctx);
}

bool MyEfxAPORT::Initialize(uint32_t sampleRate, uint32_t channels) {
    m_sampleRate = sampleRate;
    m_channels = channels;
    m_ctx = dsp_create_context(sampleRate, channels);
    return (m_ctx != nullptr);
}

void MyEfxAPORT::Process(const float* in, float* out, uint32_t frames) {
    // 这里直接调用 dsp inner-loop（实时安全：无 malloc、无锁、无IO）
    dsp_process_block(m_ctx, in, out, frames, m_channels);
}

void MyEfxAPORT::SetGain(float g) {
    // 非实时线程调用：直接写到 ctx。真实项目要用 lock-free 更新或参数双缓冲。
    m_gainLocal = g;
    if (m_ctx) {
        // 直接调整上下文里 gain 字段（知道 ctx 的内部结构）
        // 这里强制转换，配合我们 dsp_wrapper 的实现：
        typedef struct { float gain; unsigned sr; unsigned ch; } _CTX;
        _CTX* c = (_CTX*)m_ctx;
        c->gain = g;
    }
}
