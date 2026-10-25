# MicLockTray

####  See [Release Page](https://github.com/dillacorn/MicLockTray/releases) for install directions!

Windows tray app that keeps microphone input volumes fixed at your chosen percentages. Lock one or multiple recording devices independently, even across disconnects and restarts. Volume corrections use Windows Core Audio events rather than periodic polling.

> Some hardware/app combos keep “helpfully” adjusting mic gain. MicLockTray pins it where you want it.

---

## Features

- Lock multiple microphones separately, each with its own target (1–100%).
- Remember locked and unlocked devices after they are unplugged, reconnected, or Windows restarts.
- Keep the previous default-recording-device lock as a fallback for microphones without individual settings.
- Right-click the tray icon to view microphones, set targets, pause/resume enforcement, disable notifications, configure per-user autorun, or exit.
- A light-grey tray menu and a one-time first-launch right-click hint.
- Automatically follows the Windows display language (English and Simplified Chinese).
- No admin required, extra utilities, telemetry, or log files. Standalone .NET 8 Windows EXE.
- Event-driven Windows Core Audio enforcement (`IAudioEndpointVolume`).

---

## Using MicLockTray

Right-click the microphone icon in the Windows notification area. Each connected recording device shows a lock symbol and its saved target volume. Open a microphone's submenu to **Lock/Unlock** it or **Set target volume**. Setting a device volume automatically locks it.

Previously configured microphones appear under **Disconnected microphones** when unplugged. Their locks and target percentages are restored when Windows recognizes the same audio endpoint again. Changing USB ports, reinstalling drivers, or replacing hardware may produce a different Windows endpoint ID.

**Default mic fallback** preserves the behavior of older releases: it locks default recording devices that have no explicit per-device setting. Unlocking a specific microphone overrides the fallback for that device. This fallback can also be turned off independently.

All settings are stored per user under `%LOCALAPPDATA%\\MicLockTray\\config.json`.

---

## Localization

MicLockTray detects the Windows user-interface culture at startup. English and Simplified Chinese are currently included; unsupported languages safely fall back to English. Translation strings remain embedded in the main executable, so localization does not add runtime dependencies or extra language files.

Simplified Chinese translation adapted from [BaoZiFly-233/MicLockTray](https://github.com/BaoZiFly-233/MicLockTray).

---

## Development note

MicLockTray began as my own standalone FOSS utility. Large language models have since been used as development tools to help write, review, refactor, document, and troubleshoot enhancements. LLM-generated changes are not treated as inherently correct; maintenance and release decisions remain my responsibility.

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
