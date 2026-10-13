#pragma once

struct StreamFormatCandidate
{
    unsigned short containerBits;
    unsigned short validBits;
    bool floatingPoint;
};

template<class TryFormat>
bool SelectStreamFormat(TryFormat& tryFormat)
{
    static const StreamFormatCandidate candidates[] = {
        {32, 32, true}, {32, 32, false}, {32, 24, false},
        {24, 24, false}, {16, 16, false}
    };
    for (unsigned i = 0; i < sizeof(candidates) / sizeof(candidates[0]); ++i)
        if (tryFormat(candidates[i])) return true;
    return false;
}
