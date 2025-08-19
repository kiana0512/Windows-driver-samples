#pragma once
#include <cstdint>
#include <vector>

// 写 32-bit IEEE Float 的 WAVEFORMATEXTENSIBLE (常见于 Win11 共享模式内部格式)
// interleaved: 交错的 float32 缓冲（LRLR...）
// frames: 每声道样本数
// sampleRate: 48000 最常见
// channels: 2（立体声）
bool write_wav_float32(const char* path,
                       const float* interleaved,
                       uint32_t frames,
                       uint32_t sampleRate,
                       uint16_t channels);

// 方便传 vector 的重载
inline bool write_wav_float32(const char* path,
                              const std::vector<float>& interleaved,
                              uint32_t sampleRate,
                              uint16_t channels) {
    if (interleaved.empty()) return false;
    uint32_t frames = static_cast<uint32_t>(interleaved.size() / channels);
    return write_wav_float32(path, interleaved.data(), frames, sampleRate, channels);
}
