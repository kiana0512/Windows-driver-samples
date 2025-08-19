#include "dsp_wrapper.h"
#include <stdlib.h>
#include <string.h>
#include <math.h>

//======================================================
// 实用宏与辅助函数
//======================================================
#ifndef M_PI
#define M_PI 3.14159265358979323846
#endif

static inline float dB_to_linear(float db) {
    return (float)pow(10.0, db / 20.0);
}

static inline float clampf(float x, float lo, float hi) {
    return x < lo ? lo : (x > hi ? hi : x);
}

static inline float softclip(float x) {
    // 简单软限幅：tanh 风格
    const float k = 1.5f;
    return tanhf(k * x);
}

//======================================================
// Biquad（双二阶）滤波器：低搁架 / 峰值 / 高搁架
// 每个通道各自一套状态
//======================================================
typedef struct {
    // 系数
    float b0, b1, b2, a1, a2;
    // 状态（Direct Form I 或 II 均可，这里用 DF1）
    float x1, x2, y1, y2;
    // 目标系数（参数更新时写入这组，处理时原子切换）
    volatile float t_b0, t_b1, t_b2, t_a1, t_a2;
    volatile int   update_pending;
    int enabled;
} Biquad;

static void biquad_reset(Biquad* s) {
    s->x1 = s->x2 = s->y1 = s->y2 = 0.0f;
}

static float biquad_process(Biquad* s, float x) {
    if (s->update_pending) {
        // 无锁切换（尽量保持简单、近似原子）
        s->b0 = s->t_b0; s->b1 = s->t_b1; s->b2 = s->t_b2; s->a1 = s->t_a1; s->a2 = s->t_a2;
        s->update_pending = 0;
        // 不重置状态，以防参数步进造成突变；若需要可插入系数光滑
    }
    if (!s->enabled) return x;

    float y = s->b0 * x + s->b1 * s->x1 + s->b2 * s->x2
                        - s->a1 * s->y1 - s->a2 * s->y2;
    s->x2 = s->x1; s->x1 = x;
    s->y2 = s->y1; s->y1 = y;
    return y;
}

// 设计函数：低搁架 / 峰值 / 高搁架（Audio EQ Cookbook）
static void biquad_design_lowshelf(Biquad* s, float fs, float f0, float gain_db, float Q) {
    float A  = dB_to_linear(gain_db);
    float w0 = 2.f * (float)M_PI * (f0 / fs);
    float alpha = sinf(w0) / (2.f * Q);
    float cosw0 = cosf(w0);

    float sqrtA = sqrtf(A);
    float b0 =     A*( (A+1) - (A-1)*cosw0 + 2*sqrtA*alpha );
    float b1 =  2*A*( (A-1) - (A+1)*cosw0 );
    float b2 =     A*( (A+1) - (A-1)*cosw0 - 2*sqrtA*alpha );
    float a0 =         (A+1) + (A-1)*cosw0 + 2*sqrtA*alpha;
    float a1 =   -2*( (A-1) + (A+1)*cosw0 );
    float a2 =         (A+1) + (A-1)*cosw0 - 2*sqrtA*alpha;

    s->t_b0 = b0/a0; s->t_b1 = b1/a0; s->t_b2 = b2/a0;
    s->t_a1 = a1/a0; s->t_a2 = a2/a0;
    s->update_pending = 1;
}

static void biquad_design_peaking(Biquad* s, float fs, float f0, float gain_db, float Q) {
    float A  = dB_to_linear(gain_db);
    float w0 = 2.f * (float)M_PI * (f0 / fs);
    float alpha = sinf(w0) / (2.f * Q);
    float cosw0 = cosf(w0);

    float b0 = 1 + alpha*A;
    float b1 = -2*cosw0;
    float b2 = 1 - alpha*A;
    float a0 = 1 + alpha/A;
    float a1 = -2*cosw0;
    float a2 = 1 - alpha/A;

    s->t_b0 = b0/a0; s->t_b1 = b1/a0; s->t_b2 = b2/a0;
    s->t_a1 = a1/a0; s->t_a2 = a2/a0;
    s->update_pending = 1;
}

static void biquad_design_highshelf(Biquad* s, float fs, float f0, float gain_db, float Q) {
    float A  = dB_to_linear(gain_db);
    float w0 = 2.f * (float)M_PI * (f0 / fs);
    float alpha = sinf(w0) / (2.f * Q);
    float cosw0 = cosf(w0);

    float sqrtA = sqrtf(A);
    float b0 =     A*( (A+1) + (A-1)*cosw0 + 2*sqrtA*alpha );
    float b1 = -2*A*( (A-1) + (A+1)*cosw0 );
    float b2 =     A*( (A+1) + (A-1)*cosw0 - 2*sqrtA*alpha );
    float a0 =         (A+1) - (A-1)*cosw0 + 2*sqrtA*alpha;
    float a1 =    2*( (A-1) - (A+1)*cosw0 );
    float a2 =         (A+1) - (A-1)*cosw0 - 2*sqrtA*alpha;

    s->t_b0 = b0/a0; s->t_b1 = b1/a0; s->t_b2 = b2/a0;
    s->t_a1 = a1/a0; s->t_a2 = a2/a0;
    s->update_pending = 1;
}

//======================================================
// 混响：简化 Schroeder（每声道：4 梳状 + 2 全通），带 pre-delay
// 为简洁与实时安全，使用 malloc 于 create 阶段，process 不分配
//======================================================
typedef struct {
    float* buf;
    int    len;
    int    idx;
    float  feedback;
} Comb;

typedef struct {
    float* buf;
    int    len;
    int    idx;
    float  feedback;
} Allpass;

typedef struct {
    // 预延迟
    float* predelay;
    int    pd_len;
    int    pd_w, pd_r;

    // 梳状 & 全通
    Comb   comb[4];
    Allpass ap[2];

    float  wet;       // 0..1
    float  room_size; // 0.2..0.9
    float  damp;      // 0..0.7
    int    enabled;
} ReverbChan;

// 便于不同采样率缩放（基于 48k 的典型取值）
static int ms_to_samples(float ms, unsigned sr) {
    int n = (int)(ms * 0.001f * (float)sr + 0.5f);
    if (n < 1) n = 1;
    return n;
}

static void comb_init(Comb* c, int len, float fb) {
    c->buf = (float*)calloc((size_t)len, sizeof(float));
    c->len = len;
    c->idx = 0;
    c->feedback = fb;
}

static void allpass_init(Allpass* a, int len, float fb) {
    a->buf = (float*)calloc((size_t)len, sizeof(float));
    a->len = len;
    a->idx = 0;
    a->feedback = fb;
}

static inline float comb_process(Comb* c, float x, float damp) {
    // simple damp via one-pole lowpass on the feedback
    float y = c->buf[c->idx];
    float z = (1.f - damp) * y + damp * 0.f; // 可扩展为对 y 的低通，这里简化
    float v = x + z * c->feedback;
    c->buf[c->idx] = v;
    if (++c->idx >= c->len) c->idx = 0;
    return y;
}

static inline float allpass_process(Allpass* a, float x) {
    float y = a->buf[a->idx];
    float v = x + (-a->feedback) * y;
    a->buf[a->idx] = v;
    float out = y + a->feedback * v;
    if (++a->idx >= a->len) a->idx = 0;
    return out;
}

static void reverb_free(ReverbChan* r) {
    if (!r) return;
    free(r->predelay);
    for (int i=0;i<4;i++) free(r->comb[i].buf);
    for (int i=0;i<2;i++) free(r->ap[i].buf);
    memset(r, 0, sizeof(*r));
}

static void reverb_init(ReverbChan* r, unsigned sr, float wet, float room_size, float damp, float pre_delay_ms) {
    memset(r, 0, sizeof(*r));
    r->wet = clampf(wet, 0.f, 1.f);
    r->room_size = clampf(room_size, 0.2f, 0.95f);
    r->damp = clampf(damp, 0.f, 0.7f);
    r->enabled = 0;

    // 预延迟
    r->pd_len = ms_to_samples(pre_delay_ms, sr);
    if (r->pd_len < 1) r->pd_len = 1;
    r->predelay = (float*)calloc((size_t)r->pd_len, sizeof(float));
    r->pd_w = r->pd_r = 0;

    // comb 与 allpass 的长度（基于 48k，做采样率缩放）
    // 典型值（ms）：comb: 29.7, 37.1, 41.1, 43.7; allpass: 5.0, 1.7
    int c0 = ms_to_samples(29.7f * 48000.f / (float)sr, sr);
    int c1 = ms_to_samples(37.1f * 48000.f / (float)sr, sr);
    int c2 = ms_to_samples(41.1f * 48000.f / (float)sr, sr);
    int c3 = ms_to_samples(43.7f * 48000.f / (float)sr, sr);
    int a0 = ms_to_samples(5.0f  * 48000.f / (float)sr, sr);
    int a1 = ms_to_samples(1.7f  * 48000.f / (float)sr, sr);

    comb_init(&r->comb[0], c0, 0.77f * r->room_size);
    comb_init(&r->comb[1], c1, 0.80f * r->room_size);
    comb_init(&r->comb[2], c2, 0.84f * r->room_size);
    comb_init(&r->comb[3], c3, 0.88f * r->room_size);

    allpass_init(&r->ap[0], a0, 0.5f);
    allpass_init(&r->ap[1], a1, 0.5f);
}

static inline float reverb_process_sample(ReverbChan* r, float x) {
    // 预延迟
    float pd_out = r->predelay[r->pd_r];
    r->predelay[r->pd_w] = x;
    if (++r->pd_w >= r->pd_len) r->pd_w = 0;
    if (++r->pd_r >= r->pd_len) r->pd_r = 0;

    // 四个梳状求和
    float s = 0.f;
    for (int i=0;i<4;i++) s += comb_process(&r->comb[i], pd_out, r->damp);
    s *= 0.25f;

    // 两个全通
    s = allpass_process(&r->ap[0], s);
    s = allpass_process(&r->ap[1], s);
    return s;
}

//======================================================
// 总上下文
//======================================================
typedef struct {
    unsigned sr;
    unsigned ch;

    // 参数（非实时写，实时读需要近似无锁）
    volatile float gain;

    // 3 段 EQ（每通道串联：低搁架 / 峰值 / 高搁架）
    Biquad* eqBands[3];   // [band] 指向长度为 ch 的数组
    volatile float eq_freq[3];
    volatile float eq_q[3];
    volatile float eq_gain_db[3];
    volatile int   eq_enabled[3];

    // 混响
    ReverbChan* reverb; // 每通道一个
    volatile int   reverb_enabled;
    volatile float reverb_wet;
    volatile float reverb_room;
    volatile float reverb_damp;
    volatile float reverb_pre_ms;

    // 软限幅器
    volatile int   limiter_enabled;

} DSP_CTX;

//======================================================
// 创建/销毁/复位
//======================================================
void* dsp_create_context(unsigned sampleRate, unsigned channels) {
    if (channels < 1) channels = 1;
    DSP_CTX* c = (DSP_CTX*)calloc(1, sizeof(DSP_CTX));
    if (!c) return NULL;
    c->sr = sampleRate;
    c->ch = channels;

    c->gain = 1.0f;

    // 默认 EQ 参数（可调）：低搁架 100Hz +3dB, 峰值 1kHz +0dB, 高搁架 8kHz +0dB, Q=0.707
    c->eq_freq[0] = 100.f;  c->eq_gain_db[0] = 0.f;   c->eq_q[0] = 0.707f; c->eq_enabled[0] = 0;
    c->eq_freq[1] = 1000.f; c->eq_gain_db[1] = 0.f;   c->eq_q[1] = 1.0f;   c->eq_enabled[1] = 0;
    c->eq_freq[2] = 8000.f; c->eq_gain_db[2] = 0.f;   c->eq_q[2] = 0.707f; c->eq_enabled[2] = 0;

    for (int b=0;b<3;b++) {
        c->eqBands[b] = (Biquad*)calloc(channels, sizeof(Biquad));
        if (!c->eqBands[b]) { dsp_destroy_context(c); return NULL; }
        for (unsigned ch=0; ch<channels; ++ch) {
            biquad_reset(&c->eqBands[b][ch]);
            c->eqBands[b][ch].enabled = 0;
        }
    }
    // 初次设计（虽然默认禁用）
    biquad_design_lowshelf (&c->eqBands[0][0], (float)c->sr, c->eq_freq[0], c->eq_gain_db[0], c->eq_q[0]);
    biquad_design_peaking  (&c->eqBands[1][0], (float)c->sr, c->eq_freq[1], c->eq_gain_db[1], c->eq_q[1]);
    biquad_design_highshelf(&c->eqBands[2][0], (float)c->sr, c->eq_freq[2], c->eq_gain_db[2], c->eq_q[2]);
    // 把 band[0] 的系数复制到每个通道（避免重复计算；参数一变会统一更新）
    for (int b=0;b<3;b++) {
        for (unsigned ch=1; ch<channels; ++ch) {
            c->eqBands[b][ch] = c->eqBands[b][0];
            biquad_reset(&c->eqBands[b][ch]); // 状态独立
        }
    }

    // 混响（默认禁用）
    c->reverb_enabled = 0;
    c->reverb_wet  = 0.2f;
    c->reverb_room = 0.7f;
    c->reverb_damp = 0.3f;
    c->reverb_pre_ms = 20.f;
    c->reverb = (ReverbChan*)calloc(channels, sizeof(ReverbChan));
    if (!c->reverb) { dsp_destroy_context(c); return NULL; }
    for (unsigned chn=0; chn<channels; ++chn) {
        reverb_init(&c->reverb[chn], c->sr, c->reverb_wet, c->reverb_room, c->reverb_damp, c->reverb_pre_ms);
    }

    c->limiter_enabled = 1; // 默认开启软限幅，防止测试时爆音
    return c;
}

void dsp_reset(void* ctx) {
    if (!ctx) return;
    DSP_CTX* c = (DSP_CTX*)ctx;
    for (int b=0;b<3;b++) {
        for (unsigned ch=0; ch<c->ch; ++ch) biquad_reset(&c->eqBands[b][ch]);
    }
    for (unsigned ch=0; ch<c->ch; ++ch) {
        reverb_free(&c->reverb[ch]);
        reverb_init(&c->reverb[ch], c->sr, c->reverb_wet, c->reverb_room, c->reverb_damp, c->reverb_pre_ms);
    }
}

void dsp_destroy_context(void* ctx) {
    if (!ctx) return;
    DSP_CTX* c = (DSP_CTX*)ctx;
    if (c->reverb) {
        for (unsigned ch=0; ch<c->ch; ++ch) reverb_free(&c->reverb[ch]);
        free(c->reverb);
    }
    for (int b=0;b<3;b++) free(c->eqBands[b]);
    free(c);
}

//======================================================
// 参数设置（非实时线程调用）
//======================================================
void dsp_set_gain(void* ctx, float linear_gain) {
    if (!ctx) return;
    DSP_CTX* c = (DSP_CTX*)ctx;
    c->gain = linear_gain;
}

void dsp_set_eq_enabled(void* ctx, int band, int enabled) {
    if (!ctx) return;
    if (band < 0 || band > 2) return;
    DSP_CTX* c = (DSP_CTX*)ctx;
    c->eq_enabled[band] = enabled ? 1 : 0;
    for (unsigned ch=0; ch<c->ch; ++ch) c->eqBands[band][ch].enabled = c->eq_enabled[band];
}

void dsp_set_eq_params(void* ctx, int band, float freq_hz, float q, float gain_db) {
    if (!ctx) return;
    if (band < 0 || band > 2) return;
    DSP_CTX* c = (DSP_CTX*)ctx;
    c->eq_freq[band] = clampf(freq_hz, 20.f, 20000.f);
    c->eq_q[band]    = clampf(q, 0.3f, 8.f);
    c->eq_gain_db[band] = clampf(gain_db, -24.f, 24.f);

    // 重新设计一组系数（拷给每通道的 t_*），实时线程会自动切换
    for (unsigned ch=0; ch<c->ch; ++ch) {
        switch (band) {
            case 0: biquad_design_lowshelf (&c->eqBands[0][ch], (float)c->sr, c->eq_freq[0], c->eq_gain_db[0], c->eq_q[0]); break;
            case 1: biquad_design_peaking  (&c->eqBands[1][ch], (float)c->sr, c->eq_freq[1], c->eq_gain_db[1], c->eq_q[1]); break;
            case 2: biquad_design_highshelf(&c->eqBands[2][ch], (float)c->sr, c->eq_freq[2], c->eq_gain_db[2], c->eq_q[2]); break;
        }
    }
}

void dsp_set_reverb_enabled(void* ctx, int enabled) {
    if (!ctx) return;
    DSP_CTX* c = (DSP_CTX*)ctx;
    c->reverb_enabled = enabled ? 1 : 0;
}

void dsp_set_reverb_params(void* ctx, float wet, float room_size, float damp, float pre_delay_ms) {
    if (!ctx) return;
    DSP_CTX* c = (DSP_CTX*)ctx;
    c->reverb_wet   = clampf(wet, 0.f, 1.f);
    c->reverb_room  = clampf(room_size, 0.2f, 0.95f);
    c->reverb_damp  = clampf(damp, 0.f, 0.7f);
    c->reverb_pre_ms= clampf(pre_delay_ms, 0.f, 100.f);

    // 重新初始化混响通道（非实时调用）
    for (unsigned ch=0; ch<c->ch; ++ch) {
        reverb_free(&c->reverb[ch]);
        reverb_init(&c->reverb[ch], c->sr, c->reverb_wet, c->reverb_room, c->reverb_damp, c->reverb_pre_ms);
        c->reverb[ch].enabled = c->reverb_enabled;
    }
}

void dsp_set_limiter_enabled(void* ctx, int enabled) {
    if (!ctx) return;
    DSP_CTX* c = (DSP_CTX*)ctx;
    c->limiter_enabled = enabled ? 1 : 0;
}

//======================================================
// 实时处理
// 流程：PreGain → EQ(LS→Peak→HS) → Reverb(湿干混合) → Limiter → 输出
//======================================================
void dsp_process_block(void* ctx, const float* in, float* out, size_t frames, unsigned channels) {
    DSP_CTX* c = (DSP_CTX*)ctx;
    if (!c || channels != c->ch || frames == 0) {
        // 兜底：直通
        if (in != out) memcpy(out, in, sizeof(float)*frames*channels);
        return;
    }

    const unsigned ch = c->ch;
    const float G = c->gain;
    const int eqEn0 = c->eq_enabled[0];
    const int eqEn1 = c->eq_enabled[1];
    const int eqEn2 = c->eq_enabled[2];
    const int rvEn  = c->reverb_enabled;
    const float wet = c->reverb_wet;
    const int limitEn = c->limiter_enabled;

    // 逐帧逐通道处理（保持简单，可后续矢量化）
    for (size_t n=0; n<frames; ++n) {
        for (unsigned cc=0; cc<ch; ++cc) {
            float x = in[n*ch + cc];

            // PreGain
            x *= G;

            // EQ 串联
            if (eqEn0) x = biquad_process(&c->eqBands[0][cc], x);
            if (eqEn1) x = biquad_process(&c->eqBands[1][cc], x);
            if (eqEn2) x = biquad_process(&c->eqBands[2][cc], x);

            float y = x;

            // Reverb（湿干）
            if (rvEn && c->reverb && c->reverb[cc].enabled) {
                float rv = reverb_process_sample(&c->reverb[cc], x);
                y = (1.0f - wet) * x + wet * rv;
            }

            // 软限幅（防止爆音）
            if (limitEn) y = softclip(y);

            out[n*ch + cc] = y;
        }
    }
}
