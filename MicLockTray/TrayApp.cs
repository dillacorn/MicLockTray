using System;
using System.Collections.Generic;
using System.ComponentModel;
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
    private VolumePrompt? _activeVolumePrompt;
    private bool _cleanedUp;

    public TrayApp(bool startedHidden = false)
    {
        var text = UiText.Current;
        _trayIcon = LoadTrayIcon();
        _menu = new ContextMenuStrip
        {
            ShowImageMargin = false,
            ShowCheckMargin = false
        };
        ApplyMenuAppearance();
        _menu.Opening += RebuildMenu;

        _icon = new NotifyIcon
        {
            Icon = _trayIcon,
            Text = text.AppName,
            ContextMenuStrip = _menu,
            Visible = true
        };

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

    private void RebuildMenu(object? sender, CancelEventArgs e)
    {
        ApplyMenuAppearance();
        PopulateMenu();
    }

    private void ApplyMenuAppearance()
    {
        _menuDark = MenuAppearance.IsDark(Settings.MenuTheme, MenuAppearance.WindowsDarkMode());
        _menuBackground = _menuDark ? Color.FromArgb(39, 44, 51) : Color.FromArgb(211, 216, 222);
        _menuText = _menuDark ? Color.FromArgb(233, 237, 241) : Color.FromArgb(32, 42, 54);
        _menu.BackColor = _menuBackground;
        _menu.ForeColor = _menuText;
        _menu.Renderer = new ToolStripProfessionalRenderer(new MicColors(_menuDark));
    }

    private void PopulateMenu()
    {
        var text = UiText.Current;
        var oldItems = _menu.Items.Cast<ToolStripItem>().ToArray();
        _menu.Items.Clear();
        foreach (var item in oldItems) item.Dispose();

        var devices = _enforcer.GetDevices();
        var activeIds = new HashSet<string>(devices.Select(d => d.Id), StringComparer.OrdinalIgnoreCase);
        UpdateTooltip(devices);

        var heading = NewItem(text.Microphones, enabled: false);
        _menu.Items.Add(heading);

        if (devices.Length == 0)
            _menu.Items.Add(NewItem(text.NoMicrophones, enabled: false));

        // One click on the microphone row toggles its lock.
        foreach (var device in devices)
        {
            var saved = Settings.GetDevice(device.Id);
            bool locked = saved?.Locked ?? (device.IsDefault && Settings.DefaultLocked);
            int target = saved?.TargetPercent ?? (device.IsDefault && Settings.DefaultLocked
                ? Settings.TargetPercent : device.CurrentPercent);
            var item = NewItem($"{LockMark(locked)} {device.Name} — {target}%", () =>
            {
                Settings.SetDeviceLock(device.Id, device.Name, !locked,
                    saved?.TargetPercent ?? (device.IsDefault ? Settings.TargetPercent : device.CurrentPercent));
                _enforcer.ForceToTarget();
            });
            item.ToolTipText = locked ? text.UnlockMicrophone : text.LockMicrophone;
            _menu.Items.Add(item);
        }

        if (devices.Length > 0)
        {
            var volumeMenu = NewItem(text.SetMicrophoneVolume);
            foreach (var device in devices)
            {
                var saved = Settings.GetDevice(device.Id);
                int target = saved?.TargetPercent ?? device.CurrentPercent;
                volumeMenu.DropDownItems.Add(NewItem($"{device.Name} — {target}%", () =>
                    PromptDeviceVolume(device.Id, device.Name, target)));
            }
            _menu.Items.Add(volumeMenu);

            var propertiesMenu = NewItem(text.MicrophoneProperties);
            foreach (var device in devices)
            {
                string endpointId = device.Id;
                propertiesMenu.DropDownItems.Add(NewItem(device.Name, () =>
                    OpenSoundSettings("ms-settings:sound-properties?endpointId=" +
                        Uri.EscapeDataString(endpointId))));
            }
            _menu.Items.Add(propertiesMenu);
        }

        var offline = Settings.GetSavedDevices().Where(device => !activeIds.Contains(device.Id)).ToArray();
        if (offline.Length > 0)
        {
            var offlineMenu = NewItem(UiText.Format(text.DisconnectedMicrophonesFormat, offline.Length));
            foreach (var device in offline)
            {
                var item = NewItem($"{LockMark(device.Locked)} {device.Name} — {device.TargetPercent}%");
                item.DropDownItems.Add(NewItem(device.Locked ? text.UnlockMicrophone : text.LockMicrophone,
                    () =>
                    {
                        Settings.SetDeviceLock(device.Id, device.Name, !device.Locked, device.TargetPercent);

                    }));
                item.DropDownItems.Add(NewItem(text.SetTargetVolume, () =>
                    PromptDeviceVolume(device.Id, device.Name, device.TargetPercent)));
                item.DropDownItems.Add(new ToolStripSeparator());
                item.DropDownItems.Add(NewItem(text.ForgetMicrophone, () =>
                {
                    Settings.ForgetDevice(device.Id);

                }));
                offlineMenu.DropDownItems.Add(item);
            }
            _menu.Items.Add(offlineMenu);
        }

        var defaultItem = NewItem(
            $"{LockMark(Settings.DefaultLocked)} {text.DefaultFallback} — {Settings.TargetPercent}%");
        defaultItem.ToolTipText = text.DefaultFallbackDescription;
        defaultItem.DropDownItems.Add(NewItem(
            Settings.DefaultLocked ? text.UnlockMicrophone : text.LockMicrophone, () =>
            {
                Settings.SetDefaultLocked(!Settings.DefaultLocked);
                _enforcer.ForceToTarget();

            }));
        defaultItem.DropDownItems.Add(NewItem(text.SetTargetVolume, PromptDefaultVolume));
        _menu.Items.Add(defaultItem);

        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add(NewItem(_enforcer.IsEnabled ? text.PauseEnforcement : text.ResumeEnforcement,
            ToggleEnforcement));
        _menu.Items.Add(NewItem(Settings.NotificationsEnabled ? text.DisableNotifications : text.EnableNotifications,
            () =>
            {
                Settings.SetNotificationsEnabled(!Settings.NotificationsEnabled);

            }));
        _menu.Items.Add(new ToolStripSeparator());

        bool autorun = Installer.IsInstalled();
        _menu.Items.Add(NewItem(text.InstallAutorun, () =>
        {
            Installer.Install();

        }, enabled: !autorun));
        _menu.Items.Add(NewItem(text.RemoveAutorun, () =>
        {
            Installer.Uninstall();

        }, enabled: autorun));
        _menu.Items.Add(new ToolStripSeparator());

        var soundMenu = NewItem(text.SoundPanels);
        soundMenu.DropDownItems.Add(NewItem(text.WindowsSoundSettings,
            () => OpenSoundSettings("ms-settings:sound")));
        soundMenu.DropDownItems.Add(NewItem(text.ClassicSoundPanel, OpenClassicRecordingPanel));
        _menu.Items.Add(soundMenu);

        var themeMenu = NewItem(text.MenuTheme);
        foreach (var option in new[]
        {
            (Code: "system", Label: text.ThemeSystem),
            (Code: "light", Label: text.ThemeLight),
            (Code: "dark", Label: text.ThemeDark)
        })
        {
            bool selected = Settings.MenuTheme == option.Code;
            themeMenu.DropDownItems.Add(NewItem(
                (selected ? "✓ " : "   ") + option.Label, () =>
                {
                    Settings.SetMenuTheme(option.Code);
                    ApplyMenuAppearance();
                }));
        }
        _menu.Items.Add(themeMenu);

        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add(NewItem(text.Exit, () => ExitThreadCore()));
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
            Settings.EffectiveTarget(device.Id, device.IsDefault).HasValue);
        _icon.Text = UiText.Format(text.TrayTooltipFormat, lockedCount);
    }

    private ToolStripMenuItem NewItem(string label, Action? onClick = null, bool enabled = true)
    {
        var item = new ToolStripMenuItem(label)
        {
            BackColor = _menuBackground,
            ForeColor = _menuText,
            Enabled = enabled
        };
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
        catch (Exception ex) { ReportSoundSettingsError(ex); }
    }

    private void OpenSoundSettings(string uri)
    {
        try { Process.Start(new ProcessStartInfo(uri) { UseShellExecute = true }); }
        catch (Exception ex) { ReportSoundSettingsError(ex); }
    }

    private static void ReportSoundSettingsError(Exception ex)
    {
        var text = UiText.Current;
        MessageBox.Show(UiText.Format(text.AudioSettingsErrorFormat, ex.Message),
            text.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
    }

    private int? AskVolume(int initialPercent, string? name = null)
    {
        if (_activeVolumePrompt is { IsDisposed: false })
        {
            _activeVolumePrompt.Activate();
            _activeVolumePrompt.BringToFront();
            return null;
        }

        using var dialog = new VolumePrompt(initialPercent, name);
        _activeVolumePrompt = dialog;
        try
        {
            return dialog.ShowDialog() == DialogResult.OK ? dialog.Value : null;
        }
        finally
        {
            _activeVolumePrompt = null;
        }
    }

    private void PromptDeviceVolume(string id, string name, int initialPercent)
    {
        int? chosen = AskVolume(initialPercent, name);
        if (chosen is null) return;

        Settings.SetDeviceTarget(id, name, chosen.Value);
        _enforcer.ForceToTarget();

        Notify(UiText.Format(UiText.Current.TargetSetBalloonFormat, chosen.Value), 900);
    }

    private void PromptDefaultVolume()
    {
        int? chosen = AskVolume(Settings.TargetPercent);
        if (chosen is null) return;

        Settings.SetTarget(chosen.Value);
        _enforcer.ForceToTarget();

        Notify(UiText.Format(UiText.Current.TargetSetBalloonFormat, Settings.TargetPercent), 900);
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
            _background = dark ? Color.FromArgb(39, 44, 51) : Color.FromArgb(211, 216, 222);
            _selected = dark ? Color.FromArgb(66, 79, 92) : Color.FromArgb(170, 184, 198);
            _border = dark ? Color.FromArgb(89, 102, 116) : Color.FromArgb(140, 153, 166);
            _separator = dark ? Color.FromArgb(82, 92, 105) : Color.FromArgb(158, 167, 178);
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

internal sealed class VolumePrompt : Form
{
    private readonly NumericUpDown _volumeInput;
    private readonly Button _okButton;
    private readonly Button _cancelButton;

    public int Value => (int)_volumeInput.Value;

    public VolumePrompt(int current, string? microphoneName = null)
    {
        var text = UiText.Current;

        Text = microphoneName is null ? text.VolumeDialogTitle : $"{text.VolumeDialogTitle} — {microphoneName}";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterScreen;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        AutoScaleMode = AutoScaleMode.Font;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Padding = new Padding(12);

        var label = new Label
        {
            Text = text.VolumeLabel,
            AutoSize = true,
            Anchor = AnchorStyles.Left
        };

        _volumeInput = new NumericUpDown
        {
            Minimum = 1,
            Maximum = 100,
            Value = Math.Clamp(current, 1, 100),
            Anchor = AnchorStyles.Left,
            Width = 110
        };

        _okButton = new Button
        {
            Text = text.Ok,
            DialogResult = DialogResult.OK,
            AutoSize = true
        };

        _cancelButton = new Button
        {
            Text = text.Cancel,
            DialogResult = DialogResult.Cancel,
            AutoSize = true
        };

        var buttons = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Dock = DockStyle.Fill,
            WrapContents = false,
            Padding = new Padding(0, 8, 0, 0)
        };
        buttons.Controls.Add(_cancelButton);
        buttons.Controls.Add(_okButton);

        var grid = new TableLayoutPanel
        {
            AutoSize = true,
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 2
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        grid.Controls.Add(label, 0, 0);
        grid.Controls.Add(_volumeInput, 1, 0);
        grid.Controls.Add(buttons, 0, 1);
        grid.SetColumnSpan(buttons, 2);

        Controls.Add(grid);
        AcceptButton = _okButton;
        CancelButton = _cancelButton;

        Shown += (_, _) =>
        {
            try
            {
                _volumeInput.Focus();
                _volumeInput.Select(0, _volumeInput.Text.Length);
            }
            catch { }
        };
    }
}
