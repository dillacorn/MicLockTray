using System;
using System.IO;
using System.Windows.Forms;

namespace MicLockTray;

internal sealed class TrayApp : ApplicationContext
{
    private const string IconFileName = "Papirus-Team-Papirus-Devices-Audio-input-microphone.ico";

    private readonly NotifyIcon _icon;
    private readonly System.Drawing.Icon _trayIcon;
    private readonly ToolStripMenuItem _toggleItem;
    private readonly ToolStripMenuItem _installItem;
    private readonly ToolStripMenuItem _uninstallItem;
    private readonly ToolStripMenuItem _setTargetItem;
    private readonly MicEnforcer _enforcer;
    private bool _cleanedUp;

    public TrayApp()
    {
        var text = UiText.Current;
        _trayIcon = LoadTrayIcon();

        _icon = new NotifyIcon
        {
            Icon = _trayIcon,
            Text = UiText.Format(text.TrayTooltipFormat, Settings.TargetPercent),
            Visible = true,
            ContextMenuStrip = new ContextMenuStrip()
        };

        _toggleItem = new ToolStripMenuItem(text.PauseEnforcement);
        _setTargetItem = new ToolStripMenuItem(text.SetTargetVolume);
        _icon.ContextMenuStrip!.Items.Add(_toggleItem);
        _icon.ContextMenuStrip.Items.Add(_setTargetItem);
        _icon.ContextMenuStrip.Items.Add(new ToolStripSeparator());

        _installItem = new ToolStripMenuItem(text.InstallAutorun) { Enabled = !Installer.IsInstalled() };
        _uninstallItem = new ToolStripMenuItem(text.RemoveAutorun) { Enabled = Installer.IsInstalled() };
        _icon.ContextMenuStrip.Items.Add(_installItem);
        _icon.ContextMenuStrip.Items.Add(_uninstallItem);
        _icon.ContextMenuStrip.Items.Add(new ToolStripSeparator());

        var exitItem = new ToolStripMenuItem(text.Exit);
        _icon.ContextMenuStrip.Items.Add(exitItem);

        _enforcer = new MicEnforcer(() => Settings.TargetPercent / 100f);
        _enforcer.Enable();
        _enforcer.ForceToTarget();

        _toggleItem.Click += (_, _) => ToggleEnforcement();
        _setTargetItem.Click += (_, _) => PromptAndSetTarget();
        _installItem.Click += (_, _) => { Installer.Install(); RefreshInstallMenu(); };
        _uninstallItem.Click += (_, _) => { Installer.Uninstall(); RefreshInstallMenu(); };
        exitItem.Click += (_, _) => ExitThreadCore();

        try
        {
            _icon.ShowBalloonTip(
                1200,
                text.AppName,
                UiText.Format(text.StartupBalloonFormat, Settings.TargetPercent),
                ToolTipIcon.Info);
        }
        catch { }

        Application.ApplicationExit += (_, _) => Cleanup();
    }

    private static System.Drawing.Icon LoadTrayIcon()
    {
        try
        {
            var exePath = Application.ExecutablePath;
            var exeDirectory = Path.GetDirectoryName(exePath) ?? AppContext.BaseDirectory;
            var baseDirectoryIcon = Path.Combine(AppContext.BaseDirectory, IconFileName);
            var executableDirectoryIcon = Path.Combine(exeDirectory, IconFileName);

            if (File.Exists(baseDirectoryIcon)) return new System.Drawing.Icon(baseDirectoryIcon);
            if (File.Exists(executableDirectoryIcon)) return new System.Drawing.Icon(executableDirectoryIcon);

            return System.Drawing.Icon.ExtractAssociatedIcon(exePath) ?? System.Drawing.SystemIcons.Application;
        }
        catch
        {
            try
            {
                return System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath)
                    ?? System.Drawing.SystemIcons.Application;
            }
            catch
            {
                return System.Drawing.SystemIcons.Application;
            }
        }
    }

    private void RefreshInstallMenu()
    {
        bool installed = Installer.IsInstalled();
        _installItem.Enabled = !installed;
        _uninstallItem.Enabled = installed;
    }

    private void ToggleEnforcement()
    {
        var text = UiText.Current;

        if (_enforcer.IsEnabled)
        {
            _enforcer.Disable();
            _toggleItem.Text = text.ResumeEnforcement;
            try { _icon.ShowBalloonTip(800, text.AppName, text.PausedBalloon, ToolTipIcon.None); } catch { }
            return;
        }

        _enforcer.Enable();
        _enforcer.ForceToTarget();
        _toggleItem.Text = text.PauseEnforcement;
        try
        {
            _icon.ShowBalloonTip(
                800,
                text.AppName,
                UiText.Format(text.ResumedBalloonFormat, Settings.TargetPercent),
                ToolTipIcon.Info);
        }
        catch { }
    }

    private void PromptAndSetTarget()
    {
        var text = UiText.Current;
        using var dialog = new VolumePrompt(Settings.TargetPercent);
        if (dialog.ShowDialog() != DialogResult.OK) return;

        Settings.SetTarget(dialog.Value);
        _icon.Text = UiText.Format(text.TrayTooltipFormat, Settings.TargetPercent);
        _enforcer.ForceToTarget();
        try
        {
            _icon.ShowBalloonTip(
                900,
                text.AppName,
                UiText.Format(text.TargetSetBalloonFormat, Settings.TargetPercent),
                ToolTipIcon.Info);
        }
        catch { }
    }

    private void Cleanup()
    {
        if (_cleanedUp) return;
        _cleanedUp = true;

        try { _enforcer.Dispose(); } catch { }
        try { _icon.Visible = false; } catch { }
        try { _icon.ContextMenuStrip?.Dispose(); } catch { }
        try { _icon.Dispose(); } catch { }
        try { _trayIcon.Dispose(); } catch { }
    }

    protected override void ExitThreadCore()
    {
        Cleanup();
        base.ExitThreadCore();
    }
}

internal sealed class VolumePrompt : Form
{
    private readonly NumericUpDown _volumeInput;
    private readonly Button _okButton;
    private readonly Button _cancelButton;

    public int Value => (int)_volumeInput.Value;

    public VolumePrompt(int current)
    {
        var text = UiText.Current;

        Text = text.VolumeDialogTitle;
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
