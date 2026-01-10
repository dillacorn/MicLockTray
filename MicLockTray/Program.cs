// Program.cs
// MicLockTray — event-driven mic volume lock (no polling).
// Locks default capture device to a target % by listening to Core Audio callbacks.
// Removes: "Force now" menu item/action.
// Adds: periodic working-set trim (very light) to mitigate visible memory growth under stress tests.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace MicLockTray
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            using var mutex = new Mutex(true, $@"Local\MicLockTray-{Environment.UserName}", out bool createdNew);
            if (!createdNew) return;

            Settings.Load();

            var arg = args.Length > 0 ? args[0].Trim().ToLowerInvariant() : string.Empty;
            if (arg == "--install") { Installer.Install(); return; }
            if (arg == "--uninstall") { Installer.Uninstall(); return; }

            ApplicationConfiguration.Initialize();
            Application.Run(new TrayApp());
        }
    }

    internal static class Installer
    {
        private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string RunValueName = "MicLockTray";

        public static void Install()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true) ??
                                Registry.CurrentUser.CreateSubKey(RunKeyPath, writable: true);

                if (key == null) return;
                var exe = Application.ExecutablePath;
                key.SetValue(RunValueName, $"\"{exe}\"", RegistryValueKind.String);
                MessageBox.Show("Autorun installed.", "MicLockTray", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to install autorun:\n{ex.Message}", "MicLockTray", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void Uninstall()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true);
                key?.DeleteValue(RunValueName, throwOnMissingValue: false);
                MessageBox.Show("Autorun removed.", "MicLockTray", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to remove autorun:\n{ex.Message}", "MicLockTray", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static bool IsInstalled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false);
                var v = key?.GetValue(RunValueName) as string;
                return !string.IsNullOrWhiteSpace(v);
            }
            catch { return false; }
        }
    }

    internal static class Settings
    {
        private static readonly string Dir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MicLockTray");
        private static readonly string PathJson = System.IO.Path.Combine(Dir, "settings.json");

        public static int TargetPercent { get; set; } = 65;
        public static bool Enabled { get; set; } = true;

        public static void Load()
        {
            try
            {
                if (!File.Exists(PathJson)) return;
                var json = File.ReadAllText(PathJson);
                var dto = JsonSerializer.Deserialize<SettingsDto>(json);
                if (dto == null) return;

                TargetPercent = Clamp(dto.TargetPercent, 1, 100);
                Enabled = dto.Enabled;
            }
            catch { }
        }

        public static void Save()
        {
            try
            {
                Directory.CreateDirectory(Dir);
                var dto = new SettingsDto
                {
                    TargetPercent = Clamp(TargetPercent, 1, 100),
                    Enabled = Enabled
                };
                var json = JsonSerializer.Serialize(dto, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(PathJson, json);
            }
            catch { }
        }

        private static int Clamp(int v, int lo, int hi)
        {
            if (v < lo) return lo;
            if (v > hi) return hi;
            return v;
        }

        private sealed class SettingsDto
        {
            public int TargetPercent { get; set; } = 65;
            public bool Enabled { get; set; } = true;
        }
    }

    internal sealed class TrayApp : ApplicationContext
    {
        private readonly NotifyIcon _icon = new();
        private readonly ToolStripMenuItem _miToggle;
        private readonly ToolStripMenuItem _miSetTarget;
        private readonly ToolStripMenuItem _miInstall;
        private readonly ToolStripMenuItem _miUninstall;

        private readonly MicEnforcer _enforcer;

        private readonly System.Windows.Forms.Timer _trimTimer = new() { Interval = 60_000 };

        public TrayApp()
        {
            _icon.Text = "MicLockTray";
            _icon.Icon = System.Drawing.SystemIcons.Information;
            _icon.Visible = true;

            _icon.ContextMenuStrip = new ContextMenuStrip();

            _miToggle = new ToolStripMenuItem(Settings.Enabled ? "Pause enforcement" : "Resume enforcement");
            _miSetTarget = new ToolStripMenuItem($"Set target volume… ({Settings.TargetPercent}%)");
            _miInstall = new ToolStripMenuItem("Install autorun");
            _miUninstall = new ToolStripMenuItem("Remove autorun");

            _icon.ContextMenuStrip.Items.Add(_miToggle);
            _icon.ContextMenuStrip.Items.Add(_miSetTarget);
            _icon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
            _icon.ContextMenuStrip.Items.Add(_miInstall);
            _icon.ContextMenuStrip.Items.Add(_miUninstall);
            _icon.ContextMenuStrip.Items.Add(new ToolStripSeparator());

            var miExit = new ToolStripMenuItem("Exit");
            _icon.ContextMenuStrip.Items.Add(miExit);

            _enforcer = new MicEnforcer(() => Settings.TargetPercent / 100f);
            _enforcer.Enable();
            _enforcer.ForceToTarget();

            _miToggle.Click += (_, _) => ToggleEnforcement();
            _miSetTarget.Click += (_, _) => PromptAndSetTarget();
            _miInstall.Click += (_, _) => { Installer.Install(); RefreshInstallMenu(); };
            _miUninstall.Click += (_, _) => { Installer.Uninstall(); RefreshInstallMenu(); };
            miExit.Click += (_, _) => ExitThreadCore();

            RefreshInstallMenu();

            _trimTimer.Tick += (_, _) => MemoryTrimmer.Trim();
            _trimTimer.Start();

            Application.ApplicationExit += (_, _) =>
            {
                try
                {
                    _icon.Visible = false;
                    _icon.Dispose();
                    _enforcer.Dispose();
                    _trimTimer.Stop();
                }
                catch { }
            };
        }

        private void RefreshInstallMenu()
        {
            bool installed = Installer.IsInstalled();
            _miInstall.Enabled = !installed;
            _miUninstall.Enabled = installed;
        }

        private void ToggleEnforcement()
        {
            if (_enforcer.IsEnabled)
            {
                _enforcer.Disable();
                _miToggle.Text = "Resume enforcement";
                try { _icon.ShowBalloonTip(800, "MicLockTray", "Paused.", ToolTipIcon.None); } catch { }
            }
            else
            {
                _enforcer.Enable();
                _enforcer.ForceToTarget();
                _miToggle.Text = "Pause enforcement";
                try { _icon.ShowBalloonTip(800, "MicLockTray", "Resumed.", ToolTipIcon.None); } catch { }
            }

            Settings.Enabled = _enforcer.IsEnabled;
            Settings.Save();
        }

        private void PromptAndSetTarget()
        {
            using var dlg = new VolumePrompt(Settings.TargetPercent);
            if (dlg.ShowDialog() != DialogResult.OK) return;

            int v = dlg.Value;
            Settings.TargetPercent = v;
            Settings.Save();

            _miSetTarget.Text = $"Set target volume… ({Settings.TargetPercent}%)";
            _enforcer.ForceToTarget();
        }

        protected override void ExitThreadCore()
        {
            try { _enforcer.Dispose(); } catch { }
            try { _trimTimer.Stop(); } catch { }
            try { _icon.Visible = false; _icon.Dispose(); } catch { }
            base.ExitThreadCore();
        }
    }

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
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;

            // DPI/text-scaling safe layout (no hard-coded pixel positioning).
            AutoScaleMode = AutoScaleMode.Font;
            Padding = new Padding(12);
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;

            var lbl = new Label
            {
                Text = "Volume (1–100):",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 0, 12, 0)
            };

            _num = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 100,
                Value = Math.Min(100, Math.Max(1, current)),
                Width = 110,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0)
            };

            _ok = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                AutoSize = true,
                MinimumSize = new System.Drawing.Size(80, 0),
                Margin = new Padding(0)
            };

            _cancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                AutoSize = true,
                MinimumSize = new System.Drawing.Size(80, 0),
                Margin = new Padding(12, 0, 0, 0)
            };

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                WrapContents = false,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 12, 0, 0)
            };

            // RightToLeft flow: add OK first so it ends up on the far right.
            buttons.Controls.Add(_ok);
            buttons.Controls.Add(_cancel);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                RowCount = 2,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            layout.Controls.Add(lbl, 0, 0);
            layout.Controls.Add(_num, 1, 0);
            layout.Controls.Add(buttons, 0, 1);
            layout.SetColumnSpan(buttons, 2);

            Controls.Add(layout);

            AcceptButton = _ok;
            CancelButton = _cancel;
        }
    }

    internal static class MemoryTrimmer
    {
        [DllImport("psapi.dll")]
        private static extern int EmptyWorkingSet(IntPtr hProcess);

        public static void Trim()
        {
            try
            {
                using var p = Process.GetCurrentProcess();
                _ = EmptyWorkingSet(p.Handle);
            }
            catch { }
        }
    }

    internal sealed class MicEnforcer : IDisposable
    {
        private readonly Func<float> _targetScalar;
        private bool _enabled;

        private CoreAudio.IMMDeviceEnumerator? _enumerator;
        private CoreAudio.IMMDevice? _device;
        private CoreAudio.IAudioEndpointVolume? _endpoint;
        private VolumeCallback? _callback;

        public bool IsEnabled => _enabled;

        public MicEnforcer(Func<float> targetScalar)
        {
            _targetScalar = targetScalar;
        }

        public void Enable()
        {
            if (_enabled) return;
            _enabled = true;

            Bind();
        }

        public void Disable()
        {
            if (!_enabled) return;
            _enabled = false;

            Unbind();
        }

        public void ForceToTarget()
        {
            try
            {
                if (!_enabled) return;
                var ep = _endpoint;
                if (ep == null) return;

                float s = Clamp01(_targetScalar());
                _ = ep.SetMasterVolumeLevelScalar(s, Guid.Empty);
            }
            catch { }
        }

        private void Bind()
        {
            try
            {
                _enumerator = (CoreAudio.IMMDeviceEnumerator)new CoreAudio.MMDeviceEnumerator();

                _ = _enumerator.GetDefaultAudioEndpoint(CoreAudio.EDataFlow.eCapture, CoreAudio.ERole.eCommunications, out _device);

                var iid = typeof(CoreAudio.IAudioEndpointVolume).GUID;
                _ = _device.Activate(ref iid, 0, IntPtr.Zero, out object obj);
                _endpoint = (CoreAudio.IAudioEndpointVolume)obj;

                _callback = new VolumeCallback(this);
                _ = _endpoint.RegisterControlChangeNotify(_callback);
            }
            catch
            {
                Unbind();
            }
        }

        private void Unbind()
        {
            try
            {
                if (_endpoint != null && _callback != null)
                {
                    try { _ = _endpoint.UnregisterControlChangeNotify(_callback); } catch { }
                }
            }
            catch { }

            try { if (_endpoint != null) Marshal.ReleaseComObject(_endpoint); } catch { }
            try { if (_device != null) Marshal.ReleaseComObject(_device); } catch { }
            try { if (_enumerator != null) Marshal.ReleaseComObject(_enumerator); } catch { }

            _endpoint = null;
            _device = null;
            _enumerator = null;
            _callback = null;
        }

        public void Dispose()
        {
            Unbind();
        }

        private void OnNotify(float newScalar)
        {
            if (!_enabled) return;

            float target = Clamp01(_targetScalar());

            // deadband to avoid chatter
            if (Math.Abs(newScalar - target) < 0.005f) return;

            try { _ = _endpoint?.SetMasterVolumeLevelScalar(target, Guid.Empty); } catch { }
        }

        private static float Clamp01(float v)
        {
            if (v < 0f) return 0f;
            if (v > 1f) return 1f;
            return v;
        }

        private sealed class VolumeCallback : CoreAudio.IAudioEndpointVolumeCallback
        {
            private readonly MicEnforcer _p;
            public VolumeCallback(MicEnforcer p) => _p = p;

            public void OnNotify(IntPtr pNotify)
            {
                try
                {
                    var data = Marshal.PtrToStructure<CoreAudio.AUDIO_VOLUME_NOTIFICATION_DATA>(pNotify);
                    _p.OnNotify(data.fMasterVolume);
                }
                catch { }
            }
        }
    }

    internal static class CoreAudio
    {
        internal enum EDataFlow
        {
            eRender = 0,
            eCapture = 1,
            eAll = 2
        }

        internal enum ERole
        {
            eConsole = 0,
            eMultimedia = 1,
            eCommunications = 2
        }

        [ComImport]
        [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        internal class MMDeviceEnumerator
        {
        }

        [ComImport]
        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IMMDeviceEnumerator
        {
            [PreserveSig] int EnumAudioEndpoints(EDataFlow dataFlow, int dwStateMask, out object ppDevices);
            [PreserveSig] int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);
            [PreserveSig] int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string pwstrId, out IMMDevice ppDevice);
            [PreserveSig] int RegisterEndpointNotificationCallback(IntPtr pClient);
            [PreserveSig] int UnregisterEndpointNotificationCallback(IntPtr pClient);
        }

        [ComImport]
        [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IMMDevice
        {
            [PreserveSig] int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
            [PreserveSig] int OpenPropertyStore(int stgmAccess, out object ppProperties);
            [PreserveSig] int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);
            [PreserveSig] int GetState(out int pdwState);
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct AUDIO_VOLUME_NOTIFICATION_DATA
        {
            public Guid guidEventContext;
            [MarshalAs(UnmanagedType.Bool)]
            public bool bMuted;
            public float fMasterVolume;
            public uint nChannels;
        }

        [ComImport]
        [Guid("657804FA-D6AD-4496-8A60-352752AF4F89")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IAudioEndpointVolumeCallback
        {
            void OnNotify(IntPtr pNotify);
        }

        [ComImport]
        [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IAudioEndpointVolume
        {
            [PreserveSig] int RegisterControlChangeNotify([MarshalAs(UnmanagedType.Interface)] IAudioEndpointVolumeCallback pNotify);
            [PreserveSig] int UnregisterControlChangeNotify([MarshalAs(UnmanagedType.Interface)] IAudioEndpointVolumeCallback pNotify);
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
            [PreserveSig] int GetVolumeRange(out float pflVolumeMindB, out float pflVolumeMaxdB, out float pflVolumeIncrementdB);
        }
    }
}
