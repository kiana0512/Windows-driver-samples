// main.cpp — EFX 离线验证批量测试
// 作用：生成输入扫频(48k/float/立体声)；分别用不同参数/块大小/采样率/就地处理跑一遍 DSP；
//      导出多份 out_*.wav，便于在 Audacity 中同时对比波形/频谱/响度变化。

#define _USE_MATH_DEFINES
#include <cmath>
#include <cstdint>
#include <vector>
#include <string>
#include <iostream>
#include <algorithm>
#include <chrono>

#include "dsp_wrapper.h" // 你刚换好的“增益+3段EQ+混响+限幅”版本
#include "wav_writer.h"  // 前面我给你的 32-bit float WAV 写入器

#ifndef M_PI
#define M_PI 3.14159265358979323846
#endif

using clock_type = std::chrono::high_resolution_clock;
struct Timing
{
    uint64_t blocks = 0;
    uint64_t total_us = 0;
    uint64_t max_us = 0;
};

// 生成对数扫频（便于看 EQ/频响）：50Hz → 18kHz
static void gen_log_sweep(std::vector<float> &interleaved,
                          uint32_t sampleRate, uint16_t channels,
                          float seconds, float startHz = 50.0f, float endHz = 18000.0f, float gain = 0.5f)
{
    const uint32_t frames = static_cast<uint32_t>(seconds * sampleRate);
    interleaved.assign(frames * channels, 0.0f);

    const double T = seconds;
    const double f1 = startHz, f2 = endHz;
    const double K = T / std::log(f2 / f1);

    double t = 0.0;
    for (uint32_t n = 0; n < frames; ++n)
    {
        const double phi = 2.0 * M_PI * f1 * K * (std::exp(t / K) - 1.0);
        const float v = static_cast<float>(std::sin(phi)) * gain; // 约 -6 dBFS 余量
        for (uint16_t ch = 0; ch < channels; ++ch)
        {
            interleaved[n * channels + ch] = v;
        }
        t += 1.0 / sampleRate;
    }
}

// 分块处理（支持 in-place），并统计每块处理耗时（微秒）
static void process_blocked(void *ctx,
                            float *inout,       // 当 inPlace=true 时：输入输出同一缓冲区
                            float *outSeparate, // 当 inPlace=false 时：输出缓冲区
                            uint32_t frames,
                            uint32_t sampleRate,
                            uint16_t channels,
                            uint32_t blockFrames,
                            bool inPlace,
                            Timing &timing)
{
    (void)sampleRate; // 这里不用，但保留以便你扩展
    timing = {};

    if (inPlace)
    {
        // 直接就地处理同一块（每次传递一个连续窗口）
        uint32_t done = 0;
        while (done < frames)
        {
            uint32_t todo = std::min(blockFrames, frames - done);
            auto *ptr = inout + done * channels;

            auto t0 = clock_type::now();
            dsp_process_block(ctx, ptr, ptr, todo, channels);
            auto t1 = clock_type::now();

            uint64_t us = std::chrono::duration_cast<std::chrono::microseconds>(t1 - t0).count();
            timing.blocks++;
            timing.total_us += us;
            timing.max_us = std::max(timing.max_us, us);

            done += todo;
        }
    }
    else
    {
        // 用独立的 in/out 小块缓冲
        std::vector<float> inBlk(blockFrames * channels);
        std::vector<float> outBlk(blockFrames * channels);

        uint32_t done = 0;
        while (done < frames)
        {
            uint32_t todo = std::min(blockFrames, frames - done);
            const float *src = inout + done * channels;
            std::copy_n(src, todo * channels, inBlk.data());

            auto t0 = clock_type::now();
            dsp_process_block(ctx, inBlk.data(), outBlk.data(), todo, channels);
            auto t1 = clock_type::now();

            std::copy_n(outBlk.data(), todo * channels, outSeparate + done * channels);

            uint64_t us = std::chrono::duration_cast<std::chrono::microseconds>(t1 - t0).count();
            timing.blocks++;
            timing.total_us += us;
            timing.max_us = std::max(timing.max_us, us);

            done += todo;
        }
    }
}

// 打印计时结果
static void print_timing(const char *name, const Timing &t, uint32_t blockFrames)
{
    const double avg = t.blocks ? (double)t.total_us / (double)t.blocks : 0.0;
    std::cout << "[TIMING] " << name
              << " | blocks=" << t.blocks
              << " | block=" << blockFrames << " frames"
              << " | avg=" << avg << " us/blk"
              << " | max=" << t.max_us << " us\n";
}

int main()
{
    // ---- 全局基础：Win11 典型音频流格式 ----
    const uint32_t SR48k = 48000;
    const uint32_t SR441k = 44100; // 备用测试
    const uint16_t CH_ST = 2;
    const float DURATION = 8.0f; // 8 秒扫频
    // 1) 生成 48k/立体声/float 的输入扫频，并写入 in_float.wav
    std::vector<float> in48;
    gen_log_sweep(in48, SR48k, CH_ST, DURATION, 50.0f, 18000.0f, 0.5f);
    if (!write_wav_float32("in_float.wav", in48, SR48k, CH_ST))
    {
        std::cerr << "写入 in_float.wav 失败\n";
        return -1;
    }
    std::cout << "生成 in_float.wav (48k/stereo/float)\n";

    // 为各用例准备通用变量
    Timing tim;
    const uint32_t BLOCK_10MS = 480; // 10ms @48k
    const uint32_t BLOCK_128 = 128;  // 小块测试
    const uint32_t frames48 = static_cast<uint32_t>(in48.size() / CH_ST);

    // -------------------------
    // 用例 A：空处理（Null Test）
    // 目标：验证“直通”—— out 应 ≈ in
    // -------------------------
    {
        std::vector<float> out(in48.size());
        void *ctx = dsp_create_context(SR48k, CH_ST);
        if (!ctx)
        {
            std::cerr << "ctx fail\n";
            return -2;
        }

        dsp_set_gain(ctx, 1.0f);
        dsp_set_eq_enabled(ctx, 0, 0);
        dsp_set_eq_enabled(ctx, 1, 0);
        dsp_set_eq_enabled(ctx, 2, 0);
        dsp_set_reverb_enabled(ctx, 0);
        dsp_set_limiter_enabled(ctx, 0);

        process_blocked(ctx, in48.data(), out.data(), frames48, SR48k, CH_ST, BLOCK_10MS, /*inPlace=*/false, tim);
        dsp_destroy_context(ctx);
        write_wav_float32("out_null.wav", out, SR48k, CH_ST);
        print_timing("null", tim, BLOCK_10MS);
    }

    // -------------------------
    // 用例 B：增益线性（+3.52 dB）
    // -------------------------
    {
        std::vector<float> out(in48.size());
        void *ctx = dsp_create_context(SR48k, CH_ST);
        dsp_set_gain(ctx, 1.5f);
        dsp_set_eq_enabled(ctx, 0, 0);
        dsp_set_eq_enabled(ctx, 1, 0);
        dsp_set_eq_enabled(ctx, 2, 0);
        dsp_set_reverb_enabled(ctx, 0);
        dsp_set_limiter_enabled(ctx, 0); // 为了对比“纯增益”，先关限幅

        process_blocked(ctx, in48.data(), out.data(), frames48, SR48k, CH_ST, BLOCK_10MS, false, tim);
        dsp_destroy_context(ctx);
        write_wav_float32("out_gain.wav", out, SR48k, CH_ST);
        print_timing("gain", tim, BLOCK_10MS);
    }

    // -------------------------
    // 用例 C：3 段 EQ（低/峰/高）
    // -------------------------
    {
        std::vector<float> out(in48.size());
        void *ctx = dsp_create_context(SR48k, CH_ST);
        dsp_set_gain(ctx, 1.0f);
        dsp_set_eq_enabled(ctx, 0, 1);
        dsp_set_eq_params(ctx, 0, 120.f, 0.707f, +6.f);
        dsp_set_eq_enabled(ctx, 1, 1);
        dsp_set_eq_params(ctx, 1, 1200.f, 1.2f, -6.f);
        dsp_set_eq_enabled(ctx, 2, 1);
        dsp_set_eq_params(ctx, 2, 8000.f, 0.707f, +6.f);
        dsp_set_reverb_enabled(ctx, 0);
        dsp_set_limiter_enabled(ctx, 1); // 避免 EQ 抬升导致削顶

        process_blocked(ctx, in48.data(), out.data(), frames48, SR48k, CH_ST, BLOCK_10MS, false, tim);
        dsp_destroy_context(ctx);
        write_wav_float32("out_eq.wav", out, SR48k, CH_ST);
        print_timing("eq", tim, BLOCK_10MS);
    }

    // -------------------------
    // 用例 D：混响（Schroeder）
    // -------------------------
    {
        std::vector<float> out(in48.size());
        void *ctx = dsp_create_context(SR48k, CH_ST);
        dsp_set_gain(ctx, 1.0f);
        dsp_set_reverb_enabled(ctx, 1);
        dsp_set_reverb_params(ctx, 0.25f, 0.7f, 0.3f, 20.f); // 湿比/房间/阻尼/预延迟
        dsp_set_eq_enabled(ctx, 0, 0);
        dsp_set_eq_enabled(ctx, 1, 0);
        dsp_set_eq_enabled(ctx, 2, 0);
        dsp_set_limiter_enabled(ctx, 1);

        process_blocked(ctx, in48.data(), out.data(), frames48, SR48k, CH_ST, BLOCK_10MS, false, tim);
        dsp_destroy_context(ctx);
        write_wav_float32("out_reverb.wav", out, SR48k, CH_ST);
        print_timing("reverb", tim, BLOCK_10MS);
    }

    // -------------------------
    // 用例 E：限幅（比较 on/off）
    // -------------------------
    {
        // off
        {
            std::vector<float> out(in48.size());
            void *ctx = dsp_create_context(SR48k, CH_ST);
            dsp_set_gain(ctx, 2.0f); // 有意拉大，利于观察削顶
            dsp_set_limiter_enabled(ctx, 0);
            dsp_set_eq_enabled(ctx, 0, 0);
            dsp_set_eq_enabled(ctx, 1, 0);
            dsp_set_eq_enabled(ctx, 2, 0);
            dsp_set_reverb_enabled(ctx, 0);

            process_blocked(ctx, in48.data(), out.data(), frames48, SR48k, CH_ST, BLOCK_10MS, false, tim);
            dsp_destroy_context(ctx);
            write_wav_float32("out_limiter_off.wav", out, SR48k, CH_ST);
            print_timing("limiter_off", tim, BLOCK_10MS);
        }
        // on
        {
            std::vector<float> out(in48.size());
            void *ctx = dsp_create_context(SR48k, CH_ST);
            dsp_set_gain(ctx, 2.0f);
            dsp_set_limiter_enabled(ctx, 1);
            dsp_set_eq_enabled(ctx, 0, 0);
            dsp_set_eq_enabled(ctx, 1, 0);
            dsp_set_eq_enabled(ctx, 2, 0);
            dsp_set_reverb_enabled(ctx, 0);

            process_blocked(ctx, in48.data(), out.data(), frames48, SR48k, CH_ST, BLOCK_10MS, false, tim);
            dsp_destroy_context(ctx);
            write_wav_float32("out_limiter_on.wav", out, SR48k, CH_ST);
            print_timing("limiter_on", tim, BLOCK_10MS);
        }
    }

    // -------------------------
    // 用例 F：块大小 128 帧（稳定性）
    // -------------------------
    {
        std::vector<float> out(in48.size());
        void *ctx = dsp_create_context(SR48k, CH_ST);
        dsp_set_gain(ctx, 1.2f);
        dsp_set_eq_enabled(ctx, 1, 1);
        dsp_set_eq_params(ctx, 1, 1500.f, 1.0f, +3.f);
        dsp_set_limiter_enabled(ctx, 1);

        process_blocked(ctx, in48.data(), out.data(), frames48, SR48k, CH_ST, BLOCK_128, false, tim);
        dsp_destroy_context(ctx);
        write_wav_float32("out_block128.wav", out, SR48k, CH_ST);
        print_timing("block128", tim, BLOCK_128);
    }

    // -------------------------
    // 用例 G：就地处理 in-place（验证 in==out 是否安全）
    // -------------------------
    {
        std::vector<float> inout = in48; // 复制一份用作 in-place 处理
        void *ctx = dsp_create_context(SR48k, CH_ST);
        dsp_set_gain(ctx, 1.1f);
        dsp_set_eq_enabled(ctx, 2, 1);
        dsp_set_eq_params(ctx, 2, 9000.f, 0.707f, +4.f);
        dsp_set_limiter_enabled(ctx, 1);

        process_blocked(ctx, inout.data(), /*outSeparate*/ nullptr, frames48, SR48k, CH_ST, BLOCK_10MS, true, tim);
        dsp_destroy_context(ctx);
        write_wav_float32("out_inplace.wav", inout, SR48k, CH_ST);
        print_timing("inplace", tim, BLOCK_10MS);
    }

    // -------------------------
    // 用例 H：采样率 44.1k（重生输入并处理）
    // -------------------------
    {
        std::vector<float> in441;
        gen_log_sweep(in441, SR441k, CH_ST, DURATION, 50.0f, 18000.0f, 0.5f);
        write_wav_float32("in_44100_float.wav", in441, SR441k, CH_ST);

        const uint32_t frames441 = static_cast<uint32_t>(in441.size() / CH_ST);
        std::vector<float> out(in441.size());

        void *ctx = dsp_create_context(SR441k, CH_ST);
        dsp_set_gain(ctx, 1.3f);
        dsp_set_eq_enabled(ctx, 0, 1);
        dsp_set_eq_params(ctx, 0, 100.f, 0.707f, +3.f);
        dsp_set_eq_enabled(ctx, 1, 1);
        dsp_set_eq_params(ctx, 1, 1200.f, 1.0f, -3.f);
        dsp_set_limiter_enabled(ctx, 1);

        const uint32_t BLOCK_10MS_441 = 441; // 约 10ms @44.1k
        process_blocked(ctx, in441.data(), out.data(), frames441, SR441k, CH_ST, BLOCK_10MS_441, false, tim);
        dsp_destroy_context(ctx);
        write_wav_float32("out_sr44100.wav", out, SR441k, CH_ST);
        print_timing("sr44100", tim, BLOCK_10MS_441);
    }

    std::cout << "\n已生成这些文件（当前工作目录）:\n"
              << "  in_float.wav, in_44100_float.wav(可选)\n"
              << "  out_null.wav, out_gain.wav, out_eq.wav, out_reverb.wav,\n"
              << "  out_limiter_off.wav, out_limiter_on.wav,\n"
              << "  out_block128.wav, out_inplace.wav, out_sr44100.wav\n\n"
              << "建议打开 Audacity:\n"
              << "  1) 同时导入 in_float.wav 与各 out_*.wav，比对波形振幅、频谱(分析→绘制频谱)。\n"
              << "  2) Null Test: 把 out_null 反相后与 in_float 混音，应接近静音。\n"
              << "  3) 限幅 on/off：观察波峰“圆角” vs “削顶”。\n"
              << "  4) reverb：观察尾音拉长与 ~20ms 预延迟。\n"
              << "  5) 通过上面 [TIMING] 行查看各用例的平均/最大耗时（μs/块）。\n";
    return 0;
}
