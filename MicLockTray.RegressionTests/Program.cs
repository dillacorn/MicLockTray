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
            using (var dark = new InlineVolumeEditor(75, true, "Volume",
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

            using (var light = new InlineVolumeEditor(36, false, "Volume",
                       _ => { }))
            {
                Equal(System.Drawing.Color.FromArgb(211, 216, 222), light.BackColor,
                    "inline editor follows light menu theme");
                Equal(36, light.Value, "light editor preserves initial target");
                Equal(27, light.Height, "volume editor uses compact row height");
                Equal(System.Windows.Forms.BorderStyle.None, light.Input.BorderStyle,
                    "flat editor removes native spinner chrome");
                Equal(3, light.Input.MaxLength, "editor only accepts three-digit percentages");
                Equal("36", light.Input.Text, "inline text matches current target");
                light.Step(1);
                Equal(37, light.Value, "plus step increases target by one");
                light.Step(-1);
                Equal(36, light.Value, "minus step decreases target by one");
                light.Input.Text = "85";
                light.CommitText();
                Equal(85, light.Value, "typed percentage commits within the menu");
                light.Input.Text = "999";
                light.CommitText();
                Equal(100, light.Value, "typed percentage clamps at 100");
                light.Input.Text = "invalid";
                light.CommitText();
                Equal("100", light.Input.Text, "invalid input restores saved target");
                light.SetValue(1);
                light.Step(-1);
                Equal(1, light.Value, "minus button never falls below 1");
            }

            // The tray may sit on the bottom edge of an auto-hidden taskbar.
            // The menu must open completely above the edge, not beneath it.
            var desktop = new System.Drawing.Rectangle(0, 0, 1920, 1080);
            var nearTaskbar = MenuPlacement.Calculate(
                new System.Drawing.Point(1900, 1079), desktop,
                new System.Drawing.Size(280, 280));
            Equal(true, nearTaskbar.Y + 280 <= desktop.Bottom - 4,
                "tray menu leaves clearance for auto-hidden taskbar");
            Equal(true, desktop.Bottom - (nearTaskbar.Y + 280) <= 12,
                "tray menu remains near the tray icon with no large cursor gap");
            Equal(true, nearTaskbar.X >= desktop.Left &&
                nearTaskbar.X + 280 <= desktop.Right,
                "tray menu stays within horizontal display bounds");

            var secondaryScreen = new System.Drawing.Rectangle(-2560, 0, 2560, 1440);
            var secondaryMenu = MenuPlacement.Calculate(
                new System.Drawing.Point(-10, 1400), secondaryScreen,
                new System.Drawing.Size(300, 230));
            Equal(true, secondaryMenu.X >= secondaryScreen.Left &&
                secondaryMenu.X + 300 <= secondaryScreen.Right,
                "tray menu stays on its secondary monitor");
            Equal(true, secondaryMenu.Y + 230 <= secondaryScreen.Bottom - 4,
                "secondary monitor respects bottom-edge clearance");

            var upperTaskbar = MenuPlacement.Calculate(
                new System.Drawing.Point(600, 15),
                new System.Drawing.Rectangle(0, 50, 1200, 700),
                new System.Drawing.Size(300, 240));
            Equal(true, upperTaskbar.Y >= 58,
                "top taskbar keeps context menu within work area");

            // Only microphone lock actions should keep their submenu open.
            using (var lockItem = new PersistentLockMenuItem { Text = "Unlock microphone" })
            {
                Equal(true, lockItem.KeepsMenuOpen,
                    "microphone lock action does not dismiss the context menu");
                int toggles = 0;
                lockItem.Click += (_, _) =>
                {
                    toggles++;
                    lockItem.Text = toggles % 2 == 0
                        ? "Unlock microphone" : "Lock microphone";
                };
                lockItem.PerformClick();
                Equal("Lock microphone", lockItem.Text,
                    "lock label updates immediately after a click");
                lockItem.PerformClick();
                Equal("Unlock microphone", lockItem.Text,
                    "second lock click restores the label");
                Equal(2, toggles, "successive lock clicks are accepted in one menu session");
            }

            // Native WinForms dismissal must remain enabled on both the root and
            // microphone submenus. Regression from setting AutoClose=false.
            using (var root = new System.Windows.Forms.ContextMenuStrip())
            using (var microphone = new System.Windows.Forms.ToolStripMenuItem("Microphone"))
            {
                root.Items.Add(microphone);
                microphone.DropDownItems.Add(new PersistentLockMenuItem
                {
                    Text = "Unlock microphone"
                });
                using var trayIcon = new System.Windows.Forms.NotifyIcon
                {
                    ContextMenuStrip = root
                };
                Equal(root, trayIcon.ContextMenuStrip,
                    "tray icon owns popup so Windows can manage activation");
                Equal(true, root.AutoClose, "root uses native outside-click dismissal");
                Equal(true, microphone.DropDown.AutoClose,
                    "microphone flyout uses native outside-click dismissal");
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
