# MicLockTray

A lightweight Windows tray app that keeps your microphone input volume where you set it. Lock multiple microphones independently without constant polling.

**[Download MicLockTray](https://github.com/dillacorn/MicLockTray/releases)**

## Features

- Lock any number of microphones at separate volumes (1–100%).
- Remember each lock and volume after unplugging or restarting.
- Right-click the tray icon; hover a microphone to lock/unlock it, edit its target volume inline, or open properties.
- Follow Windows light/dark mode or choose a menu theme. Pause enforcement, notifications, and autorun are optional.
- Retain the original default-microphone lock for unconfigured devices.
- No administrator access, extra utilities, telemetry, or log files. English and Simplified Chinese.

## Usage

Run `MicLockTray.exe`, **right-click its tray icon**, and hover over a microphone to lock/unlock it, adjust the volume, or open its properties. Other settings and the GitHub About link are in the main menu. Disconnected microphones keep their saved settings.

## About

Originally developed by me. Later enhancements use LLM assistance for development and documentation; changes still require review and validation.

- **Translation:** Simplified Chinese adapted from [BaoZiFly-233](https://github.com/BaoZiFly-233/MicLockTray).
- **Icon:** [Papirus Dev Team](https://www.iconarchive.com/show/papirus-devices-icons-by-papirus-team/audio-input-microphone-icon.html).
- **Related:** [Linux mic lock](https://github.com/dillacorn/arch-hypr-dots/blob/main/config/hypr/scripts/miclock.sh) · [VolLockTray (output volume)](https://github.com/dillacorn/VolLockTray)
- **License:** [MIT](LICENSE). Runs locally and collects no user data; use is subject to applicable laws.

Built and maintained as a FOSS passion project. [Donate via PayPal](https://www.paypal.com/donate/?business=XSNV4QP8JFY9Y&no_recurring=0&item_name=Built+and+maintained+out+of+passion.+Always+FOSS.+Donations+appreciated.+%28smtty%2C+MicLockTray%2C+awtarchy%29&currency_code=USD).
