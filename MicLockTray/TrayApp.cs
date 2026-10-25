using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MicLockTray;

internal sealed class TrayApp : ApplicationContext
{
    private const string IconFileName = "Papirus-Team-Papirus-Devices-Audio-input-microphone.ico";
    private static readonly Color MenuBackground = Color.FromArgb(211, 216, 222);
    private static readonly Color MenuText = Color.FromArgb(32, 42, 54);

    private readonly NotifyIcon _icon;
    private readonly Icon _trayIcon;
    private readonly ContextMenuStrip _menu;
    private readonly MicEnforcer _enforcer;
    private bool _cleanedUp;

    public TrayApp()
    {
        var text = UiText.Current;
        _trayIcon = LoadTrayIcon();
        _menu = new ContextMenuStrip
        {
            BackColor = MenuBackground,
            ForeColor = MenuText,
            Renderer = new ToolStripProfessionalRenderer(new MicColors()),
            ShowImageMargin = false,
            ShowCheckMargin = false
        };
        _menu.Opening += RebuildMenu;

        _icon = new NotifyIcon
        {
            Icon = _trayIcon,
            Text = UiText.Format(text.TrayTooltipFormat, Settings.TargetPercent),
            ContextMenuStrip = _menu,
            Visible = true
        };

        _enforcer = new MicEnforcer();
        _enforcer.Enable();
        _enforcer.ForceToTarget();
        PopulateMenu();

        // The tray overflow itself is controlled by Windows. A balloon is the
        // supported way to point first-time users toward the app's icon.
        if (!Settings.OnboardingShown)
        {
            Settings.SetOnboardingShown();
            Notify(text.FirstLaunchHint, 5000);
        }
        else
        {
            Notify(UiText.Format(text.StartupBalloonFormat, Settings.TargetPercent), 1200);
        }

        Application.ApplicationExit += (_, _) => Cleanup();
    }

    private void RebuildMenu(object? sender, CancelEventArgs e)
    {
        PopulateMenu();
    }

    private void PopulateMenu()
    {
        var text = UiText.Current;
        var oldItems = _menu.Items.Cast<ToolStripItem>().ToArray();
        _menu.Items.Clear();
        foreach (var item in oldItems) item.Dispose();

        var devices = _enforcer.GetDevices();
        var activeIds = new HashSet<string>(devices.Select(d => d.Id), StringComparer.OrdinalIgnoreCase);

        var heading = new ToolStripMenuItem(text.Microphones) { Enabled = false };
        _menu.Items.Add(heading);

        if (devices.Length == 0)
            _menu.Items.Add(new ToolStripMenuItem(text.NoMicrophones) { Enabled = false });

        // One click on the microphone row toggles its lock.
        foreach (var device in devices)
        {
            var saved = Settings.GetDevice(device.Id);
            bool locked = saved?.Locked ?? (device.IsDefault && Settings.DefaultLocked);
            int target = saved?.TargetPercent ?? Settings.TargetPercent;
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
        _menu.Items.Add(NewItem(text.Exit, () => ExitThreadCore()));
    }

    private static ToolStripMenuItem NewItem(string label, Action? onClick = null, bool enabled = true)
    {
        var item = new ToolStripMenuItem(label)
        {
            BackColor = MenuBackground,
            ForeColor = MenuText,
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

    private void PromptDeviceVolume(string id, string name, int initialPercent)
    {
        using var dialog = new VolumePrompt(initialPercent, name);
        if (dialog.ShowDialog() != DialogResult.OK) return;

        Settings.SetDeviceTarget(id, name, dialog.Value);
        _enforcer.ForceToTarget();

        Notify(UiText.Format(UiText.Current.TargetSetBalloonFormat, dialog.Value), 900);
    }

    private void PromptDefaultVolume()
    {
        using var dialog = new VolumePrompt(Settings.TargetPercent);
        if (dialog.ShowDialog() != DialogResult.OK) return;

        Settings.SetTarget(dialog.Value);
        _icon.Text = UiText.Format(UiText.Current.TrayTooltipFormat, Settings.TargetPercent);
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
        private static readonly Color Selected = Color.FromArgb(170, 184, 198);
        public override Color ToolStripDropDownBackground => MenuBackground;
        public override Color ImageMarginGradientBegin => MenuBackground;
        public override Color ImageMarginGradientMiddle => MenuBackground;
        public override Color ImageMarginGradientEnd => MenuBackground;
        public override Color MenuBorder => Color.FromArgb(140, 153, 166);
        public override Color MenuItemBorder => Color.FromArgb(130, 145, 160);
        public override Color MenuItemSelected => Selected;
        public override Color MenuItemSelectedGradientBegin => Selected;
        public override Color MenuItemSelectedGradientEnd => Selected;
        public override Color MenuItemPressedGradientBegin => Selected;
        public override Color MenuItemPressedGradientMiddle => Selected;
        public override Color MenuItemPressedGradientEnd => Selected;
        public override Color SeparatorDark => Color.FromArgb(158, 167, 178);
        public override Color SeparatorLight => MenuBackground;
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
