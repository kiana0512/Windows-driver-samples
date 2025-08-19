#include <iostream>
#include <fstream>
#include <cmath>
#include <cstdint>
#include <vector>
#include "dsp_wrapper.h"
#define _USE_MATH_DEFINES
#include <math.h>

static void write_wav_header(std::ofstream &ofs, uint32_t sampleRate, uint16_t channels, uint32_t samples) {
    uint32_t byteRate = sampleRate * channels * 2; // 16-bit PCM
    uint32_t dataSize = samples * channels * 2;
    ofs.seekp(0);
    ofs.write("RIFF", 4);
    uint32_t chunkSize = 36 + dataSize;
    ofs.write(reinterpret_cast<const char*>(&chunkSize), 4);
    ofs.write("WAVE", 4);
    ofs.write("fmt ", 4);
    uint32_t subchunk1Size = 16;
    ofs.write(reinterpret_cast<const char*>(&subchunk1Size), 4);
    uint16_t audioFormat = 1; // PCM
    ofs.write(reinterpret_cast<const char*>(&audioFormat), 2);
    ofs.write(reinterpret_cast<const char*>(&channels), 2);
    ofs.write(reinterpret_cast<const char*>(&sampleRate), 4);
    ofs.write(reinterpret_cast<const char*>(&byteRate), 4);
    uint16_t blockAlign = channels * 2;
    ofs.write(reinterpret_cast<const char*>(&blockAlign), 2);
    uint16_t bitsPerSample = 16;
    ofs.write(reinterpret_cast<const char*>(&bitsPerSample), 2);
    ofs.write("data", 4);
    ofs.write(reinterpret_cast<const char*>(&dataSize), 4);
}

// clamp float [-1,1] to int16
static inline int16_t f_to_i16(float v) {
    if (v > 1.0f) v = 1.0f;
    if (v < -1.0f) v = -1.0f;
    return static_cast<int16_t>(v * 32767.0f);
}

int main() {
    const uint32_t sampleRate = 48000;
    const uint32_t channels = 2;
    const float freq = 440.0f; // test tone A4
    const float durationSec = 5.0f;
    const uint32_t totalFrames = static_cast<uint32_t>(sampleRate * durationSec);
    const uint32_t blockFrames = 256;

    // create dsp context
    void* ctx = dsp_create_context(sampleRate, channels);
    if (!ctx) {
        std::cerr << "dsp_create_context failed\n";
        return -1;
    }
    // optional: tweak gain for clear audible effect
    // assuming context layout as in dsp_wrapper.c:
    typedef struct { float gain; unsigned sr; unsigned ch; } _CTX;
    _CTX* c = (_CTX*)ctx;
    c->gain = 1.5f; // +50% amplitude (测试用，注意音量)

    std::vector<int16_t> outSamples(totalFrames * channels);
    std::vector<float> inBlock(blockFrames * channels);
    std::vector<float> outBlock(blockFrames * channels);

    double phase = 0.0;
    double phaseInc = 2.0 * M_PI * freq / sampleRate;

    uint32_t framesDone = 0;
    while (framesDone < totalFrames) {
        uint32_t toDo = std::min<uint32_t>(blockFrames, totalFrames - framesDone);
        // generate interleaved stereo sine
        for (uint32_t f = 0; f < toDo; ++f) {
            float v = static_cast<float>(sin(phase));
            phase += phaseInc;
            if (phase > 2.0*M_PI) phase -= 2.0*M_PI;
            for (uint32_t ch = 0; ch < channels; ++ch) {
                inBlock[f*channels + ch] = v;
            }
        }

        // process block
        dsp_process_block(ctx, inBlock.data(), outBlock.data(), toDo, channels);

        // convert to int16 and store
        for (uint32_t f = 0; f < toDo; ++f) {
            for (uint32_t ch = 0; ch < channels; ++ch) {
                float vf = outBlock[f*channels + ch];
                outSamples[(framesDone + f) * channels + ch] = f_to_i16(vf);
            }
        }

        framesDone += toDo;
    }

    // write wav
    std::ofstream ofs("out_test.wav", std::ios::binary);
    write_wav_header(ofs, sampleRate, channels, totalFrames);
    ofs.write(reinterpret_cast<const char*>(outSamples.data()), outSamples.size() * sizeof(int16_t));
    ofs.close();

    dsp_destroy_context(ctx);

    std::cout << "WAV written: out_test.wav\n";
    return 0;
}
