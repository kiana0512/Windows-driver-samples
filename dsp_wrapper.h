#pragma once
#include <stddef.h>

#ifdef __cplusplus
extern "C" {
#endif

// 创建/销毁上下文
void* dsp_create_context(unsigned sampleRate, unsigned channels);
void  dsp_destroy_context(void* ctx);

// in/out: interleaved float32 PCM, frames = 每声道 samples 个数
void  dsp_process_block(void* ctx, const float* in, float* out, size_t frames, unsigned channels);

#ifdef __cplusplus
}
#endif
