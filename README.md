# MicLockTray

Windows tray app that keeps your microphone input volume fixed at a user-chosen percentage. It targets the **default recording device** and re-applies the volume every 10 seconds so other software can’t quietly change it.

> Some hardware/app combos keep “helpfully” adjusting mic gain. MicLockTray pins it where you want it.

---

## Features

- Locks mic input volume to a chosen target (1–100%).
- Simple tray UI: Force now, Pause/Resume, Set target volume, Install/Remove autorun, Exit.
- No admin required. Per-user autorun.
- Lightweight single EXE. No dependencies. No telemetry. No log files.
- Uses Windows Core Audio APIs (IAudioEndpointVolume).
- [NirCmd](https://www.nirsoft.net/utils/nircmd.html) NOT required! - This is a standalone application developed by me!
