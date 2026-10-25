using System;
using System.IO;
using System.Linq;
using MicLockTray;

internal static class Program
{
    private static int _assertions;

    [STAThread]
    private static void Main()
    {
        string directory = Path.Combine(Path.GetTempPath(),
            "MicLockTray.RegressionTests." + Guid.NewGuid().ToString("N"));

        Directory.CreateDirectory(directory);
        try
        {
            Settings.UseTestDirectory(directory);
            Equal(100, Settings.TargetPercent, "first run starts with default 100%");
            Equal(true, Settings.DefaultLocked, "first run locks default microphone");
            Equal(true, Settings.NotificationsEnabled, "existing notification default");
            Equal(false, Settings.OnboardingShown, "onboarding starts unshown");
            Equal("system", Settings.MenuTheme, "menu follows Windows by default");

            // Simulate an unmodified v1.5 config that contains only target_percent.
            File.WriteAllText(Path.Combine(directory, "config.json"), """
                {
                  "target_percent": 71
                }
                """);
            Settings.Load();
            Equal(71, Settings.TargetPercent, "legacy target remains in effect");
            Equal(71, Settings.EffectiveTarget("default-mic", true), "legacy default mic locked");
            Equal(null, Settings.EffectiveTarget("other-mic", false), "non-default not implicitly locked");
            Equal(true, Settings.DefaultLocked, "legacy fallback stays enabled");
            Equal(true, Settings.NotificationsEnabled, "legacy notifications stay enabled");
            Equal("system", Settings.MenuTheme, "legacy config migrates to system theme");

            Settings.SetDeviceLock("USB-MIC", "USB microphone", true, 37);
            Settings.SetDeviceLock("HEADSET", "Headset microphone", true, 64);
            Equal(37, Settings.EffectiveTarget("USB-MIC", false), "USB mic locked independently");
            Equal(64, Settings.EffectiveTarget("HEADSET", false), "headset locked independently");
            Equal(37, Settings.EffectiveTarget("usb-mic", true), "device IDs compare without case");
            Equal(37, Settings.EffectiveTarget("USB-MIC", false), "unplugging preserves saved lock");

            Settings.SetDeviceLock("USB-MIC", "USB microphone", false, 99);
            Equal(null, Settings.EffectiveTarget("USB-MIC", true),
                "explicit unlock overrides default-microphone fallback");
            Equal(37, Settings.GetDevice("USB-MIC")?.TargetPercent,
                "unlock preserves prior microphone target");
            Settings.SetDeviceLock("USB-MIC", "USB microphone", true, 99);
            Equal(37, Settings.EffectiveTarget("USB-MIC", false), "relock restores previous target");
            Settings.SetDeviceTarget("USB-MIC", "USB microphone", 45);
            Equal(45, Settings.EffectiveTarget("USB-MIC", false), "per-device target updates");

            Settings.SetDefaultLocked(false);
            Equal(null, Settings.EffectiveTarget("unconfigured", true), "fallback can be disabled");
            Equal(45, Settings.EffectiveTarget("USB-MIC", false),
                "disabling fallback does not disable explicit device locks");
            Settings.SetNotificationsEnabled(false);
            Settings.SetMenuTheme("dark");
            Settings.SetOnboardingShown();
            Settings.SetTarget(28);

            // Simulate restarting and reconstructing all in-memory settings.
            Settings.UseTestDirectory(directory);
            Settings.Load();
            Equal(28, Settings.TargetPercent, "legacy target round-trips through JSON");
            Equal(false, Settings.DefaultLocked, "fallback setting persists");
            Equal(false, Settings.NotificationsEnabled, "notification setting persists");
            Equal("dark", Settings.MenuTheme, "manual dark theme persists");
            Equal(true, Settings.OnboardingShown, "onboarding state persists");
            Equal(45, Settings.EffectiveTarget("USB-MIC", false), "USB target survives restart");
            Equal(64, Settings.EffectiveTarget("HEADSET", false), "headset target survives restart");
            Equal("USB microphone", Settings.GetDevice("USB-MIC")?.Name,
                "friendly name persists for disconnected microphone");
            Equal(2, Settings.GetSavedDevices().Length, "disconnected devices remain saved");

            Settings.ForgetDevice("USB-MIC");
            Equal(null, Settings.EffectiveTarget("USB-MIC", false), "forgotten mic no longer enforced");
            Settings.SetDefaultLocked(true);
            Equal(28, Settings.EffectiveTarget("USB-MIC", true),
                "forgotten mic can use default fallback again");

            Settings.SetDeviceTarget("HEADSET", "Headset microphone", -10);
            Equal(1, Settings.EffectiveTarget("HEADSET", false), "volume clamps below 1%");
            Settings.SetDeviceTarget("HEADSET", "Headset microphone", 500);
            Equal(100, Settings.EffectiveTarget("HEADSET", false), "volume clamps above 100%");

            Equal(new Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E"),
                typeof(CoreAudio.IMMDeviceCollection).GUID,
                "Windows capture enumeration interface must use the correct IID");

            Equal(true, MenuAppearance.IsDark("dark", false), "manual dark overrides Windows light");
            Equal(false, MenuAppearance.IsDark("light", true), "manual light overrides Windows dark");
            Equal(true, MenuAppearance.IsDark("system", true), "system mode follows Windows dark");
            Equal(false, MenuAppearance.IsDark("system", false), "system mode follows Windows light");
            Settings.SetMenuTheme("unknown");
            Equal("system", Settings.MenuTheme, "invalid menu theme falls back to system");

            // Match the native Core Audio ABI: these methods take LPCGUID, not a
            // by-value GUID. Reflection verifies the declaration without hardware.
            foreach (string methodName in new[]
            {
                "SetMasterVolumeLevel", "SetMasterVolumeLevelScalar",
                "SetChannelVolumeLevel", "SetChannelVolumeLevelScalar",
                "SetMute", "VolumeStepUp", "VolumeStepDown"
            })
            {
                var method = typeof(CoreAudio.IAudioEndpointVolume).GetMethod(methodName)
                    ?? throw new InvalidOperationException("Missing Core Audio method: " + methodName);
                var parameter = method.GetParameters().Last();
                Equal(typeof(Guid).MakeByRefType(), parameter.ParameterType,
                    methodName + " must pass the event-context GUID by reference");
            }

            // Exercise the embedded WinForms editor without displaying a tray icon.
            int changes = 0;
            int latest = 0;
            using (var dark = new InlineVolumeEditor(75, true, "Volume (1–100):",
                       value => { changes++; latest = value; }))
            {
                Equal(75, dark.Value, "inline editor begins at current target");
                Equal(System.Drawing.Color.FromArgb(27, 27, 27), dark.BackColor,
                    "inline editor follows near-black menu theme");
                Equal(0, changes, "building a submenu must not modify saved volume");
                dark.SetValue(83);
                Equal(83, latest, "editing target invokes the volume callback");
                Equal(1, changes, "editing target triggers a single callback");
                Equal(83, dark.Value, "edited target is displayed in same submenu");
            }

            using (var light = new InlineVolumeEditor(36, false, "Volume (1–100):",
                       _ => { }))
            {
                Equal(System.Drawing.Color.FromArgb(211, 216, 222), light.BackColor,
                    "inline editor follows light menu theme");
                Equal(36, light.Value, "light editor preserves initial target");
            }

            Console.WriteLine($"PASS: {_assertions} settings, interop and UI regression assertions");
        }
        finally
        {
            // Always clean only this test's unique temporary directory.
            Directory.Delete(directory, recursive: true);
        }
    }

    private static void Equal<T>(T expected, T actual, string description)
    {
        if (!Equals(expected, actual))
            throw new InvalidOperationException(
                $"FAIL: {description}: expected {expected}, got {actual}");
        _assertions++;
    }
}
