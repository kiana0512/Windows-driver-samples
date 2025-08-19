#pragma once
#include <stddef.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

// ================== 新增：统一常量与 EQ 类型（保持向后兼容） ==================
#ifndef MY_EQ_BANDS
#define MY_EQ_BANDS 12      // 将 EQ 段数扩展为 12 段（0..11）
#endif

// EQ 类型：0=峰值(Peak)，1=低搁架(LowShelf)，2=高搁架(HighShelf)
typedef enum { DSP_EQ_PEAK = 0, DSP_EQ_LOWSHELF = 1, DSP_EQ_HIGHSHELF = 2 } DSP_EQ_TYPE;

// ================== 公开 API（原有接口保留） ==================

// 创建/销毁（在非实时线程调用）
void* dsp_create_context(unsigned sampleRate, unsigned channels);
void  dsp_destroy_context(void* ctx);
void  dsp_reset(void* ctx);

// 处理（实时线程调用）
// in/out: interleaved float32, frames = 每声道样本数, channels = 实际通道数（与创建时一致）
void  dsp_process_block(void* ctx, const float* in, float* out, size_t frames, unsigned channels);

// -------- 参数设置（非实时线程调用，内部做无锁更新） --------

// 增益（线性倍数，例如 1.0 原音量，1.5 约 +3.52 dB）
void  dsp_set_gain(void* ctx, float linear_gain);

// 3 段 EQ（单位：Hz / dB / 无量纲 Q）
// 兼容保留：band=0 低搁架、band=1 峰值、band=2 高搁架；
// 扩展后：band 可以取 0..11，3..11 默认按“峰值”处理（如需指定类型看下面增强函数）
void  dsp_set_eq_enabled(void* ctx, int band, int enabled);
void  dsp_set_eq_params(void* ctx, int band, float freq_hz, float q, float gain_db);

// ========== 新增：增强型 EQ 设置（可选，不影响旧代码） ==========

// 设置某一段的 EQ 类型（0=Peak,1=LowShelf,2=HighShelf）
void  dsp_set_eq_type(void* ctx, int band, DSP_EQ_TYPE type);

// 一次性设置该段的所有参数（包含类型）
void  dsp_set_eq_params_ex(void* ctx, int band, float freq_hz, float q, float gain_db, DSP_EQ_TYPE type);

// 混响（Schroeder/简化 FDN 结构）
// wet: 0~1（湿声比例），pre_delay_ms: 0~100，可选，room_size/damp 范围见实现注释
void  dsp_set_reverb_enabled(void* ctx, int enabled);
void  dsp_set_reverb_params(void* ctx, float wet, float room_size, float damp, float pre_delay_ms);

// 软限幅器（防爆音，可选）
void  dsp_set_limiter_enabled(void* ctx, int enabled);

#ifdef __cplusplus
}
#endif
