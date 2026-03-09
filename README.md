# MicLockTray

####  See [Release Page](https://github.com/dillacorn/MicLockTray/releases) for install directions!

Windows tray app that keeps your microphone input volume fixed at a user-chosen percentage. It targets the **default recording device** and **corrects instantly** whenever something tries to change it (event-driven).

> Some hardware/app combos keep “helpfully” adjusting mic gain. MicLockTray pins it where you want it.

---

## Features

- Locks mic input volume to a chosen target (1–100%).
- Simple tray UI: Pause/Resume, Set target volume, Install/Remove autorun, Exit.
- No admin required. Per-user autorun.
- Lightweight single EXE. No dependencies. No telemetry. No log files.
- Uses Windows Core Audio APIs (`IAudioEndpointVolume`).
- [NirCmd](https://www.nirsoft.net/utils/nircmd.html) NOT required — this is a standalone application developed by me!

---

## Icon Credit
Icon: **“Audio input microphone Icon”** by [Papirus Dev Team](https://www.iconarchive.com/artist/papirus-team.html) on [iconarchive.com](https://www.iconarchive.com/show/papirus-devices-icons-by-papirus-team/audio-input-microphone-icon.html)

---

Need a Mic Lock Linux 🐧 solution?
- Check out my script [miclock.sh](https://github.com/dillacorn/arch-hypr-dots/blob/main/config/hypr/scripts/miclock.sh)

---

Want to lock your audio output volume instead?
- Check out [VolLockTray](https://github.com/dillacorn/VolLockTray)

## License
This project is licensed under the [MIT License](https://github.com/dillacorn/MicLockTray/blob/main/LICENSE).

### Legal Notice
This project is a general-purpose open-source utility that runs locally on the
user’s system. It does not provide a hosted service and does not collect user
data. Users are responsible for complying with laws and regulations in their
own jurisdiction when using this software.

## ☕ Donate

Built and maintained out of passion. Always FOSS. Donations appreciated.  
[Donate via PayPal](https://www.paypal.com/donate/?business=XSNV4QP8JFY9Y&no_recurring=0&item_name=Built+and+maintained+out+of+passion.+Always+FOSS.+Donations+appreciated.+%28smtty%2C+MicLockTray%2C+awtarchy%29&currency_code=USD)
