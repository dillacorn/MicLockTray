// Program.cs
// MicLockTray (standalone, no NirCmd)
// Tray app that forces the default capture device volume to a target % every 10s via Core Audio.
// Adds: user-configurable target percent (1–100). No log file writes.
// Menu: Force now (X%), Pause/Resume, Set target volume…, Install autorun, Remove autorun, Exit.

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Win32;
using FormsTimer = System.Windows.Forms.Timer;

namespace MicLockTray
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            // single instance
            using var mutex = new Mutex(true, $@"Local\MicLockTray-{Environment.UserName}", out bool createdNew);
            if (!createdNew) return;

            Settings.Load();

            var arg = args.Length > 0 ? args[0].Trim().ToLowerInvariant() : string.Empty;
            if (arg == "--install") { Installer.Install(); return; }
            if (arg == "--uninstall") { Installer.Uninstall(); return; }

            ApplicationConfiguration.Initialize(); // from WinForms template/generator
            Application.Run(new TrayApp());
        }
    }

    internal static class Installer
    {
        private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string RunValueName = "MicLockTray";
        private static string ExePath => Application.ExecutablePath;

        public static bool IsInstalled()
        {
            using var rk = Registry.CurrentUser.OpenSubKey(RunKeyPath, false);
            var val = rk?.GetValue(RunValueName) as string;
            return !string.IsNullOrWhiteSpace(val);
        }

        public static void Install()
        {
            using var rk = Registry.CurrentUser.CreateSubKey(RunKeyPath, true);
            var cmd = $"\"{ExePath}\" --hidden"; // hidden at logon
            rk.SetValue(RunValueName, cmd, RegistryValueKind.String);
            MessageBox.Show("Autorun installed.", "MicLockTray", MessageBoxButtons.OK, MessageBoxIcon.Information);
            // also start a hidden instance now
            try
            {
                var psi = new ProcessStartInfo(ExePath, "--hidden")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                Process.Start(psi);
            }
            catch { }
        }

        public static void Uninstall()
        {
            using var rk = Registry.CurrentUser.CreateSubKey(RunKeyPath, true);
            rk.DeleteValue(RunValueName, false);
            MessageBox.Show("Autorun removed.", "MicLockTray", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    internal static class Settings
    {
        private static readonly string Dir  = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MicLockTray");
        private static readonly string File = Path.Combine(Dir, "config.json");

        public static int TargetPercent { get; private set; } = 100; // 1–100 inclusive

        public static void Load()
        {
            try
            {
                if (!System.IO.File.Exists(File)) return;
                var json = System.IO.File.ReadAllText(File);
                var dto = JsonSerializer.Deserialize<ConfigDto>(json);
                if (dto != null) SetTarget(dto.target_percent);
            }
            catch { /* ignore */ }
        }

        public static void SetTarget(int percent)
        {
            if (percent < 1) percent = 1;
            if (percent > 100) percent = 100;
            TargetPercent = percent;
            Save();
        }

        private static void Save()
        {
            try
            {
                Directory.CreateDirectory(Dir);
                var dto = new ConfigDto { target_percent = TargetPercent };
                var json = JsonSerializer.Serialize(dto, new JsonSerializerOptions { WriteIndented = true });
                System.IO.File.WriteAllText(File, json);
            }
            catch { /* ignore */ }
        }

        private sealed class ConfigDto { public int target_percent { get; set; } = 100; }
    }

    internal sealed class TrayApp : ApplicationContext
    {
        private readonly NotifyIcon _icon;
        private readonly FormsTimer _timer;
        private readonly ToolStripMenuItem _miForce;
        private readonly ToolStripMenuItem _miToggle;
        private readonly ToolStripMenuItem _miInstall;
        private readonly ToolStripMenuItem _miUninstall;
        private readonly ToolStripMenuItem _miSetTarget;

        public TrayApp()
        {
            _icon = new NotifyIcon
            {
                Icon = System.Drawing.SystemIcons.Application,
                Text = $"MicLockTray: target {Settings.TargetPercent}%",
                Visible = true,
                ContextMenuStrip = new ContextMenuStrip()
            };

            _miForce     = new ToolStripMenuItem(ForceLabel());
            _miToggle    = new ToolStripMenuItem("Pause enforcement");
            _miSetTarget = new ToolStripMenuItem("Set target volume…");
            _icon.ContextMenuStrip!.Items.Add(_miForce);
            _icon.ContextMenuStrip.Items.Add(_miToggle);
            _icon.ContextMenuStrip.Items.Add(_miSetTarget);
            _icon.ContextMenuStrip.Items.Add(new ToolStripSeparator());

            _miInstall   = new ToolStripMenuItem("Install autorun") { Enabled = !Installer.IsInstalled() };
            _miUninstall = new ToolStripMenuItem("Remove autorun")  { Enabled =  Installer.IsInstalled() };
            _icon.ContextMenuStrip.Items.Add(_miInstall);
            _icon.ContextMenuStrip.Items.Add(_miUninstall);
            _icon.ContextMenuStrip.Items.Add(new ToolStripSeparator());

            var miExit = new ToolStripMenuItem("Exit");
            _icon.ContextMenuStrip.Items.Add(miExit);

            _miForce.Click += (_, _) => ForceNow();
            _miToggle.Click += (_, _) => ToggleTimer();
            _miSetTarget.Click += (_, _) => PromptAndSetTarget();
            _miInstall.Click += (_, _) => { Installer.Install(); RefreshInstallMenu(); };
            _miUninstall.Click += (_, _) => { Installer.Uninstall(); RefreshInstallMenu(); };
            miExit.Click += (_, _) => ExitThreadCore();

            _timer = new FormsTimer { Interval = 10_000 };
            _timer.Tick += (_, _) => Enforce();
            _timer.Start();

            // initial
            Enforce();
            try { _icon.ShowBalloonTip(1200, "MicLockTray", $"Microphone volume locked to {Settings.TargetPercent}%.", ToolTipIcon.Info); } catch { }

            Application.ApplicationExit += (_, _) => { try { _icon.Visible = false; _icon.Dispose(); } catch { } };
        }

        private string ForceLabel() => $"Force now ({Settings.TargetPercent}%)";

        private void RefreshInstallMenu()
        {
            bool installed = Installer.IsInstalled();
            _miInstall.Enabled = !installed;
            _miUninstall.Enabled = installed;
        }

        private void ToggleTimer()
        {
            if (_timer.Enabled)
            {
                _timer.Stop();
                _miToggle.Text = "Resume enforcement";
                try { _icon.ShowBalloonTip(800, "MicLockTray", "Paused.", ToolTipIcon.None); } catch { }
            }
            else
            {
                _timer.Start();
                _miToggle.Text = "Pause enforcement";
                try { _icon.ShowBalloonTip(800, "MicLockTray", $"Resumed. Mic will be forced to {Settings.TargetPercent}% every 10s.", ToolTipIcon.Info); } catch { }
            }
        }

        private void PromptAndSetTarget()
        {
            using var dlg = new VolumePrompt(Settings.TargetPercent);
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                Settings.SetTarget(dlg.Value);
                _miForce.Text = ForceLabel();
                _icon.Text = $"MicLockTray: target {Settings.TargetPercent}%";
                Enforce();
                try { _icon.ShowBalloonTip(900, "MicLockTray", $"Target set to {Settings.TargetPercent}%.", ToolTipIcon.Info); } catch { }
            }
        }

        private void ForceNow() => Enforce();

        private static void Enforce()
        {
            try
            {
                float scalar = Settings.TargetPercent / 100f;
                CoreAudio.SetDefaultCaptureScalar(scalar);
            }
            catch { /* ignore */ }
        }

        protected override void ExitThreadCore()
        {
            try { _timer.Stop(); } catch { }
            try { _icon.Visible = false; _icon.Dispose(); } catch { }
            base.ExitThreadCore();
        }
    }

    // Small numeric dialog (1–100) so we don't depend on Microsoft.VisualBasic.InputBox
    internal sealed class VolumePrompt : Form
    {
        private readonly NumericUpDown _num;
        private readonly Button _ok;
        private readonly Button _cancel;

        public int Value => (int)_num.Value;

        public VolumePrompt(int current)
        {
            Text = "Set target volume (%)";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false; MaximizeBox = false;
            ClientSize = new System.Drawing.Size(260, 110);

            var lbl = new Label { Text = "Volume (1–100):", Left = 12, Top = 15, AutoSize = true };
            _num = new NumericUpDown { Left = 130, Top = 12, Width = 100, Minimum = 1, Maximum = 100, Value = Math.Min(100, Math.Max(1, current)) };

            _ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 70, Width = 80, Top = 60 };
            _cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 160, Width = 80, Top = 60 };

            Controls.AddRange(new Control[] { lbl, _num, _ok, _cancel });
            AcceptButton = _ok;
            CancelButton = _cancel;
        }
    }

    internal static class CoreAudio
    {
        public enum EDataFlow : int { eRender = 0, eCapture = 1, eAll = 2 }
        public enum ERole     : int { eConsole = 0, eMultimedia = 1, eCommunications = 2 }

        // Correct vtable order
        [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDeviceEnumerator
        {
            [PreserveSig] int EnumAudioEndpoints(EDataFlow dataFlow, int dwStateMask, out IMMDeviceCollection ppDevices);
            [PreserveSig] int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);
            [PreserveSig] int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string pwstrId, out IMMDevice ppDevice);
            [PreserveSig] int RegisterEndpointNotificationCallback(IntPtr pClient);
            [PreserveSig] int UnregisterEndpointNotificationCallback(IntPtr pClient);
        }

        [ComImport, Guid("0BD7A1BE-7A1A-44DB-8397-C0F926C399A4"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDeviceCollection
        {
            [PreserveSig] int GetCount(out uint pcDevices);
            [PreserveSig] int Item(uint nDevice, out IMMDevice ppDevice);
        }

        [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        private class MMDeviceEnumerator { }

        [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDevice
        {
            [PreserveSig] int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
            [PreserveSig] int OpenPropertyStore(int stgmAccess, out IntPtr ppProperties);
            [PreserveSig] int GetId(out IntPtr ppstrId);
            [PreserveSig] int GetState(out int pdwState);
        }

        [ComImport, Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAudioEndpointVolume
        {
            [PreserveSig] int RegisterControlChangeNotify(IntPtr pNotify);
            [PreserveSig] int UnregisterControlChangeNotify(IntPtr pNotify);
            [PreserveSig] int GetChannelCount(out uint pnChannelCount);
            [PreserveSig] int SetMasterVolumeLevel(float fLevelDB, Guid pguidEventContext);
            [PreserveSig] int SetMasterVolumeLevelScalar(float fLevel, Guid pguidEventContext);
            [PreserveSig] int GetMasterVolumeLevel(out float pfLevelDB);
            [PreserveSig] int GetMasterVolumeLevelScalar(out float pfLevel);
            [PreserveSig] int SetChannelVolumeLevel(uint nChannel, float fLevelDB, Guid pguidEventContext);
            [PreserveSig] int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, Guid pguidEventContext);
            [PreserveSig] int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);
            [PreserveSig] int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel);
            [PreserveSig] int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, Guid pguidEventContext);
            [PreserveSig] int GetMute(out bool pbMute);
            [PreserveSig] int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount);
            [PreserveSig] int VolumeStepUp(Guid pguidEventContext);
            [PreserveSig] int VolumeStepDown(Guid pguidEventContext);
            [PreserveSig] int QueryHardwareSupport(out uint pdwHardwareSupportMask);
            [PreserveSig] int GetVolumeRange(out float mindB, out float maxdB, out float incrementdB);
        }

        private static Guid IID_IAudioEndpointVolume = new("5CDF2C82-841E-4546-9722-0CF74078229A");
        private const int CLSCTX_INPROC_SERVER = 0x1;

        private static IAudioEndpointVolume? TryGetEndpoint(EDataFlow flow, ERole role)
        {
            try
            {
                var enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
                int hr = enumerator.GetDefaultAudioEndpoint(flow, role, out var dev);
                if (hr != 0 || dev is null) return null;

                var iid = IID_IAudioEndpointVolume;
                hr = dev.Activate(ref iid, CLSCTX_INPROC_SERVER, IntPtr.Zero, out var obj);
                if (hr != 0 || obj is null) return null;

                return (IAudioEndpointVolume)obj;
            }
            catch { return null; }
        }

        public static void SetDefaultCaptureScalar(float scalar)
        {
            if (scalar < 0f) scalar = 0f;
            if (scalar > 1f) scalar = 1f;

            foreach (ERole role in new[] { ERole.eConsole, ERole.eMultimedia, ERole.eCommunications })
            {
                var ep = TryGetEndpoint(EDataFlow.eCapture, role);
                if (ep == null) continue;
                _ = ep.SetMasterVolumeLevelScalar(scalar, Guid.Empty);
            }
        }
    }
}
