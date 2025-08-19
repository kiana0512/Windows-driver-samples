#pragma once
#include <stdint.h>
#define MY_EQ_BANDS 12

#pragma pack(push,1)
struct MyEqBand {
    int32_t enabled;   // 0/1
    float   freq;      // Hz
    float   q;         // 0.3..8
    float   gain_db;   // -24..+24
    int32_t type;      // 0=peak,1=lowShelf,2=highShelf （你可扩展）
};

struct MyReverb {
    int32_t enabled;   // 0/1
    float   wet;       // 0..1
    float   room;      // 0.2..0.95
    float   damp;      // 0..0.7
    float   pre_ms;    // 0..100
};

struct MyDspParams {
    float   gain;                // 线性
    MyEqBand eq[MY_EQ_BANDS];    // 12 段
    MyReverb reverb;
    int32_t limiterEnabled;      // 0/1

    // 预留：后续下发“指令集/字节码”
    uint32_t opcodeSize;         // <= sizeof(opcode)
    uint8_t  opcode[512];
};
#pragma pack(pop)
