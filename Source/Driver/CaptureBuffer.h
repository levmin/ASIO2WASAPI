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

#include <vector>
#include <cstring>
#include <cstddef>
#include "SampleConversion.h"

// Single audio-thread FIFO. Store whole interleaved frames; the ASIO host
// receives planar channels. No allocation takes place while streaming.
class CaptureBuffer
{
    std::vector<unsigned char> m_data;
    size_t m_frameBytes, m_read, m_count, m_capacity;
public:
    CaptureBuffer() : m_frameBytes(0), m_read(0), m_count(0), m_capacity(0) {}
    void configure(size_t capacity, size_t frameBytes)
    {
        m_data.assign(capacity * frameBytes, 0);
        m_frameBytes = frameBytes;
        m_capacity = capacity;
        clear();
    }
    void clear() { m_read = m_count = 0; }
    size_t available() const { return m_count; }
    void push(const unsigned char* data, size_t frames, bool silent)
    {
        if (!m_capacity) return;
        for (size_t i = 0; i < frames; ++i)
        {
            if (m_count == m_capacity)
            {
                m_read = (m_read + 1) % m_capacity;
                --m_count; // bounded latency if the capture clock runs faster
            }
            unsigned char* dest = &m_data[((m_read + m_count) % m_capacity) * m_frameBytes];
            if (silent) std::memset(dest, 0, m_frameBytes);
            else std::memcpy(dest, data + i * m_frameBytes, m_frameBytes);
            ++m_count;
        }
    }
    void readPlanar(std::vector<std::vector<unsigned char> >& channels,
                    size_t frames, size_t sampleBytes)
    {
        readPlanarConverted(channels, frames, sampleBytes, false, sampleBytes, false);
    }
    void readPlanarConverted(std::vector<std::vector<unsigned char> >& channels,
        size_t frames, size_t sourceBytes, bool sourceFloat, size_t destBytes, bool destFloat)
    {
        const size_t copied = frames < m_count ? frames : m_count;
        for (size_t channel = 0; channel < channels.size(); ++channel)
        {
            if (channels[channel].empty()) continue;
            unsigned char* dest = &channels[channel][0];
            for (size_t i = 0; i < copied; ++i)
                ConvertCaptureSample(dest + i * destBytes,
                    &m_data[((m_read + i) % m_capacity) * m_frameBytes + channel * sourceBytes],
                    sourceBytes, sourceFloat, destBytes, destFloat);
            // An underrun (including startup) is defined silence, never stale data.
            std::memset(dest + copied * destBytes, 0, (frames - copied) * destBytes);
        }
        if (m_capacity) m_read = (m_read + copied) % m_capacity;
        m_count -= copied;
    }
};
