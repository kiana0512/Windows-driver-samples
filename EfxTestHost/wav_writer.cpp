#include "wav_writer.h"
#include <fstream>
#include <algorithm>

static const uint8_t SUBTYPE_IEEE_FLOAT[16] = {
    0x03,0x00,0x00,0x00, 0x00,0x00, 0x10,0x00, 0x80,0x00, 0x00,0xAA,0x00,0x38,0x9B,0x71
};

#pragma pack(push, 1)
struct WAVFMTEXT {
    uint16_t wFormatTag;      // 0xFFFE (WAVE_FORMAT_EXTENSIBLE)
    uint16_t nChannels;
    uint32_t nSamplesPerSec;
    uint32_t nAvgBytesPerSec;
    uint16_t nBlockAlign;
    uint16_t wBitsPerSample;  // 32（float32）
    uint16_t cbSize;          // 22 (size of extension)
    uint16_t wValidBitsPerSample; // 32
    uint32_t dwChannelMask;   // 0x3 = FL | FR
    uint8_t  SubFormat[16];   // KSDATAFORMAT_SUBTYPE_IEEE_FLOAT
};
#pragma pack(pop)

bool write_wav_float32(const char* path,
                       const float* interleaved,
                       uint32_t frames,
                       uint32_t sampleRate,
                       uint16_t channels)
{
    if (!path || !interleaved || frames == 0 || channels == 0) return false;

    uint32_t bytesData = frames * channels * sizeof(float);
    uint32_t fmtSize   = 40; // WAVEFORMATEXTENSIBLE size
    uint32_t riffSize  = 4 + (8 + fmtSize) + (8 + bytesData); // "WAVE" + fmt + data

    WAVFMTEXT fmt{};
    fmt.wFormatTag           = 0xFFFE; // WAVE_FORMAT_EXTENSIBLE
    fmt.nChannels            = channels;
    fmt.nSamplesPerSec       = sampleRate;
    fmt.wBitsPerSample       = 32;
    fmt.nBlockAlign          = channels * (fmt.wBitsPerSample / 8);
    fmt.nAvgBytesPerSec      = sampleRate * fmt.nBlockAlign;
    fmt.cbSize               = 22;
    fmt.wValidBitsPerSample  = 32;
    // 这里简单设置立体声掩码；如果是多通道可根据需要扩展
    fmt.dwChannelMask        = (channels == 1) ? 0x4 /*FC*/ : 0x3 /*FL|FR*/;
    std::copy(std::begin(SUBTYPE_IEEE_FLOAT), std::end(SUBTYPE_IEEE_FLOAT), fmt.SubFormat);

    std::ofstream ofs(path, std::ios::binary);
    if (!ofs) return false;

    // RIFF header
    ofs.write("RIFF", 4);
    ofs.write(reinterpret_cast<const char*>(&riffSize), 4);
    ofs.write("WAVE", 4);

    // fmt chunk
    ofs.write("fmt ", 4);
    ofs.write(reinterpret_cast<const char*>(&fmtSize), 4);
    ofs.write(reinterpret_cast<const char*>(&fmt), sizeof(fmt));

    // data chunk
    ofs.write("data", 4);
    ofs.write(reinterpret_cast<const char*>(&bytesData), 4);
    ofs.write(reinterpret_cast<const char*>(interleaved), bytesData);

    return ofs.good();
}
