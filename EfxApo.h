#pragma once
#include <cstdint>

// 这是一个轻量 C++ 类，封装了上下文与实时处理入口。
// 我们**不**在此使用 MS 的 IAudioProcessingObjectRT，先把处理逻辑做好。
class MyEfxAPORT {
public:
    MyEfxAPORT();
    ~MyEfxAPORT();

    // 初始化（非实时线程）
    bool Initialize(uint32_t sampleRate, uint32_t channels);

    // 实时安全处理函数：in/out 为 interleaved float，frames = 每通道样本数
    void Process(const float* in, float* out, uint32_t frames);

    // 非实时设置参数（示例：gain）
    void SetGain(float g);

private:
    void* m_ctx; // 绑定到 dsp_wrapper 的上下文
    uint32_t m_channels;
    uint32_t m_sampleRate;
    float m_gainLocal;
};
