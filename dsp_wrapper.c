#include "dsp_wrapper.h"
#include <stdlib.h>
#include <string.h>
#include <math.h>

typedef struct {
    float gain; // 简单参数：增益，便于一开始测试能听出来
    unsigned sr;
    unsigned ch;
} DSP_CTX;

void* dsp_create_context(unsigned sampleRate, unsigned channels) {
    DSP_CTX* c = (DSP_CTX*)malloc(sizeof(DSP_CTX));
    if (!c) return NULL;
    c->gain = 1.0f; // 初始 unity
    c->sr = sampleRate;
    c->ch = channels;
    return c;
}

void dsp_destroy_context(void* ctx) {
    if (ctx) free(ctx);
}

void dsp_process_block(void* ctx, const float* in, float* out, size_t frames, unsigned channels) {
    DSP_CTX* c = (DSP_CTX*)ctx;
    size_t total = frames * channels;
    float g = (c ? c->gain : 1.0f);

    // 简单示例：乘以 gain，并防止 denormals（非常小的浮点）
    for (size_t i = 0; i < total; ++i) {
        float v = in[i] * g;
        if (v > -1e-30f && v < 1e-30f) v = 0.0f;
        out[i] = v;
    }
}
