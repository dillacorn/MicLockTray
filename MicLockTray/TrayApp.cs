using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MicLockTray;

internal sealed class TrayApp : ApplicationContext
{
    private const string IconFileName = "Papirus-Team-Papirus-Devices-Audio-input-microphone.ico";
    private Color _menuBackground = Color.FromArgb(211, 216, 222);
    private Color _menuText = Color.FromArgb(32, 42, 54);
    private bool _menuDark;

    private readonly NotifyIcon _icon;
    private readonly Icon _trayIcon;
    private readonly ContextMenuStrip _menu;
    private readonly MicEnforcer _enforcer;
    private readonly System.Windows.Forms.Timer? _startupTipTimer;
    private bool _cleanedUp;

    public TrayApp(bool startedHidden = false)
    {
        var text = UiText.Current;
        _trayIcon = LoadTrayIcon();
        _menu = new ContextMenuStrip
        {
            ShowImageMargin = false,
            ShowCheckMargin = false,
            ShowItemToolTips = false
        };
        ApplyMenuAppearance();
        _menu.Closed += (_, _) => CloseSubmenus();

        _icon = new NotifyIcon
        {
            Icon = _trayIcon,
            Text = text.AppName,
            Visible = true
        };

        _icon.MouseUp += OnTrayMouseUp;

        _enforcer = new MicEnforcer();
        _enforcer.Enable();
        _enforcer.ForceToTarget();
        PopulateMenu();

        // Defer notifications until the Windows message loop is running. The shell
        // controls tray overflow visibility, so it cannot be forcibly expanded.
        // Autorun uses --hidden; do not consume the onboarding hint in that mode.
        if (!startedHidden)
        {
            _startupTipTimer = new System.Windows.Forms.Timer { Interval = 1500 };
            _startupTipTimer.Tick += (_, _) =>
            {
                _startupTipTimer.Stop();
                if (!Settings.NotificationsEnabled) return;
                if (!Settings.OnboardingShown)
                {
                    Notify(text.FirstLaunchHint, 5000);
                    Settings.SetOnboardingShown();
                }
                else
                {
                    Notify(UiText.Format(text.StartupBalloonFormat, Settings.TargetPercent), 1200);
                }
            };
            _startupTipTimer.Start();
        }

        Application.ApplicationExit += (_, _) => Cleanup();
    }

    private void ApplyMenuAppearance()
    {
        _menuDark = MenuAppearance.IsDark(Settings.MenuTheme, MenuAppearance.WindowsDarkMode());
        _menuBackground = _menuDark ? Color.FromArgb(27, 27, 27) : Color.FromArgb(211, 216, 222);
        _menuText = _menuDark ? Color.FromArgb(233, 237, 241) : Color.FromArgb(32, 42, 54);
        _menu.BackColor = _menuBackground;
        _menu.ForeColor = _menuText;
        _menu.Renderer = new ToolStripProfessionalRenderer(new MicColors(_menuDark));
    }

    private void OnTrayMouseUp(object? sender, MouseEventArgs e)
    {
        if (e.Button != MouseButtons.Right || _cleanedUp) return;

        // NotifyIcon's automatic context placement can put nested menus below the
        // visible desktop when the taskbar auto-hides. Measure the full menu first,
        // then place it above the cursor within the monitor's usable bounds.
        ApplyMenuAppearance();
        PopulateMenu();
        var cursor = Cursor.Position;
        var screen = Screen.FromPoint(cursor);
        var bounds = Rectangle.Intersect(screen.WorkingArea, screen.Bounds);
        var placement = MenuPlacement.Calculate(cursor, bounds,
            _menu.GetPreferredSize(Size.Empty), bottomMargin: 24);
        ConfigureDropDownDirections(placement, bounds);
        _menu.Show(placement);
    }

    private void ConfigureDropDownDirections(Point location, Rectangle bounds)
    {
        foreach (ToolStripMenuItem item in _menu.Items.OfType<ToolStripMenuItem>())
            ConfigureDropDownDirection(item, location.X, bounds);
    }

    private static void ConfigureDropDownDirection(
        ToolStripMenuItem item, int menuLeft, Rectangle bounds)
    {
        if (!item.HasDropDownItems) return;

        // Use a horizontal cascade, never a below-parent direction.
        // This prevents the first microphone submenu from dropping over an
        // auto-hidden taskbar. WinForms can still correct the placement at edges.
        var width = item.DropDown.GetPreferredSize(Size.Empty).Width;
        var rootWidth = item.Owner?.Width ?? 0;
        item.DropDownDirection =
            bounds.Right - (menuLeft + rootWidth) >= width
                ? ToolStripDropDownDirection.Right
                : ToolStripDropDownDirection.Left;
        item.DropDown.ShowItemToolTips = false;

        foreach (var nested in item.DropDownItems.OfType<ToolStripMenuItem>())
            ConfigureDropDownDirection(nested, menuLeft, bounds);
    }

    private void CloseSubmenus()
    {
        foreach (var item in _menu.Items.OfType<ToolStripMenuItem>())
            CloseSubmenu(item);
    }

    private static void CloseSubmenu(ToolStripMenuItem item)
    {
        if (!item.HasDropDownItems) return;
        foreach (var child in item.DropDownItems.OfType<ToolStripMenuItem>())
            CloseSubmenu(child);
        if (item.DropDown.Visible) item.DropDown.Close();
    }

    private void KeepSubmenuOpenOnPointerLeave(ToolStripMenuItem item)
    {
        // Keep the capture submenu visible while moving the pointer outside it.
        // Actual clicks, keyboard dismissal, and an explicit root close still work.
        item.DropDown.AutoClose = false;
        item.DropDown.Closing += (_, e) =>
        {
            e.Cancel = e.CloseReason == ToolStripDropDownCloseReason.AppFocusChange
                && _menu.Visible;
        };
        item.DropDownOpening += (_, _) =>
        {
            foreach (var other in _menu.Items.OfType<ToolStripMenuItem>())
            {
                if (!ReferenceEquals(other, item)) CloseSubmenu(other);
            }
        };
    }

    private void PopulateMenu()
    {
        var text = UiText.Current;
        var previous = _menu.Items.Cast<ToolStripItem>().ToArray();
        _menu.Items.Clear();
        foreach (var item in previous) item.Dispose();

        var devices = _enforcer.GetDevices();
        var connectedIds = new HashSet<string>(
            devices.Select(device => device.Id), StringComparer.OrdinalIgnoreCase);
        UpdateTooltip(devices);

        _menu.Items.Add(NewItem(text.AppName, enabled: false));
        if (devices.Length == 0)
            _menu.Items.Add(NewItem(_enforcer.DeviceEnumerationFailed
                ? text.DetectionFailed : text.NoMicrophones, enabled: false));

        foreach (var device in devices)
        {
            var saved = Settings.GetDevice(device.Id);
            bool locked = Settings.EffectiveTarget(device.Id, device.IsDefault).HasValue;
            int target = saved?.TargetPercent ??
                (locked ? Settings.TargetPercent : device.CurrentPercent);

            var row = NewItem(DeviceCaption(device.Name, target, locked, device.IsBound));
            var lockAction = NewItem(
                $"{LockMark(!locked)} {(locked ? text.UnlockMicrophone : text.LockMicrophone)}",
                keepOpen: true);
            lockAction.Font = new Font(_menu.Font, FontStyle.Bold);
            lockAction.Click += (_, _) =>
            {
                bool next = !Settings.EffectiveTarget(device.Id, device.IsDefault).HasValue;
                int initialTarget = Settings.GetDevice(device.Id)?.TargetPercent ??
                    (device.IsDefault ? Settings.TargetPercent : device.CurrentPercent);
                Settings.SetDeviceLock(device.Id, device.Name, next, initialTarget);
                _enforcer.ForceToTarget();

                // Update the same open submenu so the new state is unambiguous.
                bool nowLocked = Settings.EffectiveTarget(device.Id, device.IsDefault).HasValue;
                int actualTarget = Settings.GetDevice(device.Id)?.TargetPercent ?? initialTarget;
                lockAction.Text =
                    $"{LockMark(!nowLocked)} {(nowLocked ? text.UnlockMicrophone : text.LockMicrophone)}";
                row.Text = DeviceCaption(device.Name, actualTarget, nowLocked, device.IsBound);
                UpdateTooltip(_enforcer.GetDevices());
            };

            // Keep the single lock control prominent, with consistent spacing.
            row.DropDownItems.Add(lockAction);
            row.DropDownItems.Add(new ToolStripSeparator());
            row.DropDownItems.Add(NewVolumeControl(target, newValue =>
            {
                Settings.SetDeviceTarget(device.Id, device.Name, newValue);
                _enforcer.ForceToTarget();
                lockAction.Text = $"{LockMark(false)} {text.UnlockMicrophone}";
                row.Text = DeviceCaption(device.Name, newValue, true, device.IsBound);
                UpdateTooltip(_enforcer.GetDevices());
            }));
            row.DropDownItems.Add(NewItem(text.MicrophoneProperties, () =>
            {
                _menu.Close();
                OpenExternalTarget("ms-settings:sound-properties?endpointId=" +
                    Uri.EscapeDataString(device.Id));
            }));
            // Windows does not expose a reliable direct link to a specific legacy
            // recording-device dialog. This opens its containing Recording tab.
            row.DropDownItems.Add(NewItem(text.ClassicSoundPanel, () =>
            {
                _menu.Close();
                OpenClassicRecordingPanel();
            }));
            KeepSubmenuOpenOnPointerLeave(row);
            _menu.Items.Add(row);
        }

        var disconnected = Settings.GetSavedDevices()
            .Where(saved => !connectedIds.Contains(saved.Id)).ToArray();
        if (disconnected.Length > 0)
        {
            var offlineMenu = NewItem(UiText.Format(
                text.DisconnectedMicrophonesFormat, disconnected.Length));
            foreach (var device in disconnected)
            {
                var row = NewItem(DeviceCaption(device.Name, device.TargetPercent,
                    device.Locked, bound: true));
                var lockAction = NewItem(
                    $"{LockMark(!device.Locked)} {(device.Locked ? text.UnlockMicrophone : text.LockMicrophone)}",
                    keepOpen: true);
                lockAction.Font = new Font(_menu.Font, FontStyle.Bold);
                lockAction.Click += (_, _) =>
                {
                    bool next = Settings.GetDevice(device.Id)?.Locked != true;
                    Settings.SetDeviceLock(device.Id, device.Name, next, device.TargetPercent);
                    var saved = Settings.GetDevice(device.Id)!;
                    lockAction.Text =
                        $"{LockMark(!saved.Locked)} {(saved.Locked ? text.UnlockMicrophone : text.LockMicrophone)}";
                    row.Text = DeviceCaption(device.Name, saved.TargetPercent,
                        saved.Locked, bound: true);
                };

                row.DropDownItems.Add(lockAction);
                row.DropDownItems.Add(new ToolStripSeparator());
                row.DropDownItems.Add(NewVolumeControl(device.TargetPercent, newValue =>
                {
                    Settings.SetDeviceTarget(device.Id, device.Name, newValue);
                    lockAction.Text = $"{LockMark(false)} {text.UnlockMicrophone}";
                    row.Text = DeviceCaption(device.Name, newValue, locked: true, bound: true);
                }));
                row.DropDownItems.Add(NewItem(text.ForgetMicrophone, () =>
                {
                    Settings.ForgetDevice(device.Id);
                    _menu.Close();
                }));
                KeepSubmenuOpenOnPointerLeave(row);
                offlineMenu.DropDownItems.Add(row);
            }
            _menu.Items.Add(offlineMenu);
        }

        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add(NewItem(
            _enforcer.IsEnabled ? text.PauseEnforcement : text.ResumeEnforcement,
            ToggleEnforcement));
        _menu.Items.Add(NewItem(
            Settings.NotificationsEnabled ? text.DisableNotifications : text.EnableNotifications,
            () => Settings.SetNotificationsEnabled(!Settings.NotificationsEnabled)));

        _menu.Items.Add(new ToolStripSeparator());

        // One level for theme choice. Everything else stays visible at the root.
        var themes = NewItem(text.MenuTheme);
        foreach (var option in new[]
        {
            (Code: "system", Label: text.ThemeSystem),
            (Code: "light", Label: text.ThemeLight),
            (Code: "dark", Label: text.ThemeDark)
        })
        {
            bool selected = Settings.MenuTheme == option.Code;
            themes.DropDownItems.Add(NewItem(
                (selected ? "✓ " : "   ") + option.Label, () =>
                {
                    Settings.SetMenuTheme(option.Code);
                    ApplyMenuAppearance();
                }));
        }
        _menu.Items.Add(themes);

        bool autorun = Installer.IsInstalled();
        _menu.Items.Add(NewItem(
            autorun ? text.RemoveAutorun : text.InstallAutorun,
            () =>
            {
                if (autorun) Installer.Uninstall();
                else Installer.Install();
            }));

        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add(NewItem(text.About, () =>
            OpenExternalTarget("https://github.com/dillacorn/MicLockTray")));
        _menu.Items.Add(NewItem(text.Exit, () => ExitThreadCore()));
    }

    private string DeviceCaption(string name, int percent, bool locked, bool bound)
    {
        string warning = locked && _enforcer.IsEnabled && !bound
            ? $" ({UiText.Current.NotEnforcing})" : "";
        return $"{LockMark(locked)} {name} — {percent}%{warning}";
    }

    private ToolStripControlHost NewVolumeControl(int percent, Action<int> changed)
    {
        var editor = new InlineVolumeEditor(percent, _menuDark,
            UiText.Current.VolumeLabel, changed);
        return new ToolStripControlHost(editor)
        {
            AutoSize = false,
            Size = editor.Size,
            Margin = new Padding(0, 0, 0, 1),
            Padding = Padding.Empty,
            BackColor = _menuBackground
        };
    }

    private void UpdateTooltip(MicEnforcer.DeviceInfo[] devices)
    {
        var text = UiText.Current;
        if (!_enforcer.IsEnabled)
        {
            _icon.Text = text.PausedTooltip;
            return;
        }
        int lockedCount = devices.Count(device =>
            device.IsBound && Settings.EffectiveTarget(device.Id, device.IsDefault).HasValue);
        _icon.Text = UiText.Format(text.TrayTooltipFormat, lockedCount);
    }

    private ToolStripMenuItem NewItem(string label, Action? onClick = null,
        bool enabled = true, bool keepOpen = false)
    {
        var item = keepOpen
            ? new PersistentLockMenuItem { Text = label }
            : new ToolStripMenuItem(label);
        item.BackColor = _menuBackground;
        item.ForeColor = _menuText;
        item.Enabled = enabled;
        item.AutoToolTip = false;
        if (onClick is not null) item.Click += (_, _) => onClick();
        return item;
    }

    private static string LockMark(bool locked) => locked ? "🔒" : "🔓";

    private void ToggleEnforcement()
    {
        var text = UiText.Current;
        if (_enforcer.IsEnabled)
        {
            _enforcer.Disable();
            Notify(text.PausedBalloon, 800);
        }
        else
        {
            _enforcer.Enable();
            _enforcer.ForceToTarget();
            Notify(UiText.Format(text.ResumedBalloonFormat, Settings.TargetPercent), 800);
        }

    }

    private void OpenClassicRecordingPanel()
    {
        try
        {
            var startInfo = new ProcessStartInfo("control.exe") { UseShellExecute = true };
            startInfo.ArgumentList.Add("mmsys.cpl,,1");
            Process.Start(startInfo);
        }
        catch (Exception ex) { ReportOpenError(ex); }
    }

    private void OpenExternalTarget(string uri)
    {
        try { Process.Start(new ProcessStartInfo(uri) { UseShellExecute = true }); }
        catch (Exception ex) { ReportOpenError(ex); }
    }

    private static void ReportOpenError(Exception ex)
    {
        var text = UiText.Current;
        MessageBox.Show(UiText.Format(text.AudioSettingsErrorFormat, ex.Message),
            text.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
    }

    private void Notify(string message, int duration)
    {
        if (!Settings.NotificationsEnabled) return;
        try { _icon.ShowBalloonTip(duration, UiText.Current.AppName, message, ToolTipIcon.Info); }
        catch { }
    }

    private static Icon LoadTrayIcon()
    {
        try
        {
            var exePath = Application.ExecutablePath;
            var exeDirectory = Path.GetDirectoryName(exePath) ?? AppContext.BaseDirectory;
            var baseDirectoryIcon = Path.Combine(AppContext.BaseDirectory, IconFileName);
            var executableDirectoryIcon = Path.Combine(exeDirectory, IconFileName);

            if (File.Exists(baseDirectoryIcon)) return new Icon(baseDirectoryIcon);
            if (File.Exists(executableDirectoryIcon)) return new Icon(executableDirectoryIcon);

            return Icon.ExtractAssociatedIcon(exePath) ?? SystemIcons.Application;
        }
        catch
        {
            try { return Icon.ExtractAssociatedIcon(Application.ExecutablePath) ?? SystemIcons.Application; }
            catch { return SystemIcons.Application; }
        }
    }

    private void Cleanup()
    {
        if (_cleanedUp) return;
        _cleanedUp = true;
        try { _startupTipTimer?.Stop(); _startupTipTimer?.Dispose(); } catch { }
        try { _enforcer.Dispose(); } catch { }
        try { _icon.Visible = false; } catch { }
        try { _menu.Dispose(); } catch { }
        try { _icon.Dispose(); } catch { }
        try { _trayIcon.Dispose(); } catch { }
    }

    protected override void ExitThreadCore()
    {
        Cleanup();
        base.ExitThreadCore();
    }

    private sealed class MicColors : ProfessionalColorTable
    {
        private readonly Color _background;
        private readonly Color _selected;
        private readonly Color _border;
        private readonly Color _separator;

        public MicColors(bool dark)
        {
            _background = dark ? Color.FromArgb(27, 27, 27) : Color.FromArgb(211, 216, 222);
            _selected = dark ? Color.FromArgb(53, 55, 58) : Color.FromArgb(170, 184, 198);
            _border = dark ? Color.FromArgb(66, 66, 66) : Color.FromArgb(140, 153, 166);
            _separator = dark ? Color.FromArgb(69, 69, 69) : Color.FromArgb(158, 167, 178);
        }

        public override Color ToolStripDropDownBackground => _background;
        public override Color ImageMarginGradientBegin => _background;
        public override Color ImageMarginGradientMiddle => _background;
        public override Color ImageMarginGradientEnd => _background;
        public override Color MenuBorder => _border;
        public override Color MenuItemBorder => _border;
        public override Color MenuItemSelected => _selected;
        public override Color MenuItemSelectedGradientBegin => _selected;
        public override Color MenuItemSelectedGradientEnd => _selected;
        public override Color MenuItemPressedGradientBegin => _selected;
        public override Color MenuItemPressedGradientMiddle => _selected;
        public override Color MenuItemPressedGradientEnd => _selected;
        public override Color SeparatorDark => _separator;
        public override Color SeparatorLight => _background;
    }
}

