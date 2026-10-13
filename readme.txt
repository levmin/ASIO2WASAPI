Documents enclosed:
"ASIO2WASAPI.pdf" - installation and configuration instructions
gpl.txt - GNU Public License

Change log:
--------------

Falcosoft:
Version 1.3.0 

1. Added WASAPI recording/input support in shared and exclusive modes, including full-duplex and input-only operation.
2. Added independent recording-device selection, with automatic input-channel detection.
3. Added automatic sample-format conversion when input and output devices use different native formats, keeping ASIO sample types compatible.
4. Added an optional Prefer common input/output format setting for shared and exclusive modes, with automatic fallback when no common format is available.
5. Added control-panel diagnostics showing the negotiated native input/output formats, the ASIO sample format, and whether conversion is required.

Falcosoft:
Version 1.2.3

1. Low latency shared mode buffer size can be configured (and needed update period is derived from buffer size).
2. Added workaround for hosts that do not set control panel dialog's parent window correctly.
3. Fixed problems with non-standard sample rates (e.g. 49716). When shared mode format converter is used actually any sample rates can be used now.
4. Other minor fixes.

Falcosoft:
Version 1.2.2

1. Fixed speaker order in case of some frequently used speaker setups.
2. Added proper latency report in case of shared mode.
3. Other minor fixes.

Falcosoft:
Version 1.2.1

1. Added Win10+ special low latency shared mode.
This mode needs proper drivers. Usually proprietary vendor drivers from Realtek, NVIDIA etc. do not support low latency shared mode. But generic High Definition Audio driver from Microsoft does support it on Windows 10/11.
2. Added version info to ASIO control panel.

Falcosoft:
Version 1.2

1. Added support for WASAPI shared mode
2. Modified versioning so all projects get the version info from driver's version.h

Falcosoft:
Version 1.1

1.Fixed supported format query.
2.Fixed non modifiable 100 ms latency problem.
3.Fixed too quiet rendering in case of supported 24-bit PCM format.
4.Count of supported channels is calculated automatically.
5.Supported sample rates are calculated automatically.
6.Added support for default device. Driver can restore itself when default device or audio properties are changed in Windows.

Version 1.0

First stable release
Visual Studio 2012 port

Version 0.8

First beta release
