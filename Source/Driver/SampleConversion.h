/*  ASIO2WASAPI Universal ASIO Driver   
    Copyright (C) Zoltan Bacsko (Falcosoft) 2023-2026

    This file is part of ASIO2WASAPI.

    ASIO2WASAPI is free software; you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation; either version 2 of the License, or
    (at your option) any later version.

    ASIO2WASAPI is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with ASIO2WASAPI; if not, write to the Free Software
    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA
*/
#pragma once

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <cmath>

// WASAPI PCM is signed, little endian and left aligned in its container.
// Preserve the playback endpoint's ASIO type; convert only captured samples.
inline void ConvertCaptureSample(unsigned char* dest, const unsigned char* source,
    size_t sourceBytes, bool sourceFloat, size_t destBytes, bool destFloat)
{
    if (sourceBytes == destBytes && sourceFloat == destFloat)
    {
        std::memcpy(dest, source, sourceBytes);
        return;
    }
    if (!sourceFloat && !destFloat)
    {
        std::uint32_t raw = 0;
        for (size_t byte = 0; byte < sourceBytes; ++byte)
            raw |= (std::uint32_t)source[byte] << (8 * byte);
        if (destBytes > sourceBytes)
        {
            // Normalize amplitude by shifting unsigned bits. Signed negative
            // left shifts would be undefined; unsigned shifts preserve PCM bits.
            raw <<= 8 * (destBytes - sourceBytes);
        }
        else
        {
            unsigned shift = (unsigned)(8 * (sourceBytes - destBytes));
            std::uint32_t sign = (std::uint32_t)1 << (8 * sourceBytes - 1);
            bool negative = (raw & sign) != 0;
            std::uint32_t sourceMask = sourceBytes == 4 ? 0xffffffffu :
                ((std::uint32_t)1 << (8 * sourceBytes)) - 1;
            std::uint32_t magnitude = negative ? ((~raw + 1) & sourceMask) : raw;
            // Preserve Test 2's nearest rounding, ties away from zero, and
            // saturation at the positive full-scale limit.
            magnitude = (magnitude + ((std::uint32_t)1 << (shift - 1))) >> shift;
            std::uint32_t limit = ((std::uint32_t)1 << (8 * destBytes - 1)) - 1;
            if (!negative && magnitude > limit) magnitude = limit;
            raw = negative ? (std::uint32_t)0 - magnitude : magnitude;
        }
        for (size_t byte = 0; byte < destBytes; ++byte)
            dest[byte] = (unsigned char)(raw >> (8 * byte));
        return;
    }
    double value;
    if (sourceFloat)
    {
        // Inspect non-finite IEEE bits explicitly: the original release
        // project uses /fp:fast, which can optimize away NaN comparisons.
        std::uint32_t bits;
        std::memcpy(&bits, source, sizeof(bits));
        if ((bits & 0x7f800000u) == 0x7f800000u)
            value = (bits & 0x007fffffu) ? 0.0 : (bits & 0x80000000u) ? -1.0 : 1.0;
        else
        {
            float sample;
            std::memcpy(&sample, source, sizeof(sample));
            value = sample;
        }
    }
    else
    {
        std::uint32_t raw = 0;
        for (size_t byte = 0; byte < sourceBytes; ++byte)
            raw |= (std::uint32_t)source[byte] << (8 * byte);
        double scale = sourceBytes == 2 ? 32768.0 : sourceBytes == 3 ? 8388608.0 : 2147483648.0;
        value = (raw >= scale ? (double)raw - 2.0 * scale : (double)raw) / scale;
    }
    if (destFloat)
    {
        float sample = (float)value;
        std::memcpy(dest, &sample, sizeof(sample));
        return;
    }
    double scale = destBytes == 2 ? 32768.0 : destBytes == 3 ? 8388608.0 : 2147483648.0;
    double scaled;
    if (value <= -1.0) scaled = -scale;
    else if (value >= 1.0) scaled = scale - 1.0;
    else
    {
        scaled = value * scale;
        scaled = scaled < 0 ? std::ceil(scaled - 0.5) : std::floor(scaled + 0.5);
        if (scaled > scale - 1.0) scaled = scale - 1.0;
    }
    std::uint32_t raw = (std::uint32_t)(std::int64_t)scaled;
    for (size_t byte = 0; byte < destBytes; ++byte)
        dest[byte] = (unsigned char)(raw >> (8 * byte));
}
