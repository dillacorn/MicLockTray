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
        private static void Main(string[] args)
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
            var cmd = $"\"{ExePath}\" --hidden";
            rk.SetValue(RunValueName, cmd, RegistryValueKind.String);
            MessageBox.Show("Autorun installed.", "MicLockTray", MessageBoxButtons.OK, MessageBoxIcon.Information);

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

        public static int TargetPercent { get; private set; } = 100; // 1–100

        public static void Load()
        {
            try
            {
                if (!System.IO.File.Exists(File)) return;
                var json = System.IO.File.ReadAllText(File);
                var dto = JsonSerializer.Deserialize<ConfigDto>(json);
                if (dto != null) SetTarget(dto.target_percent);
            }
            catch { }
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
            catch { }
        }

        private sealed class ConfigDto { public int target_percent { get; set; } = 100; }
    }

    internal sealed class TrayApp : ApplicationContext
    {
        private readonly NotifyIcon _icon;
        private readonly ToolStripMenuItem _miToggle;
        private readonly ToolStripMenuItem _miInstall;
        private readonly ToolStripMenuItem _miUninstall;
        private readonly ToolStripMenuItem _miSetTarget;
        private readonly MicEnforcer _enforcer;

        // very light periodic working-set trim to reduce visible memory creep under stress
        private readonly System.Windows.Forms.Timer _trimTimer = new() { Interval = 60_000 };

        public TrayApp()
        {
            _icon = new NotifyIcon
            {
                Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath) ?? System.Drawing.SystemIcons.Application,
                Text = $"MicLockTray: target {Settings.TargetPercent}%",
                Visible = true,
                ContextMenuStrip = new ContextMenuStrip()
            };

            _miToggle    = new ToolStripMenuItem("Pause enforcement");
            _miSetTarget = new ToolStripMenuItem("Set target volume…");
            _icon.ContextMenuStrip!.Items.Add(_miToggle);
            _icon.ContextMenuStrip.Items.Add(_miSetTarget);
            _icon.ContextMenuStrip.Items.Add(new ToolStripSeparator());

            _miInstall   = new ToolStripMenuItem("Install autorun") { Enabled = !Installer.IsInstalled() };
            _miUninstall = new ToolStripMenuItem("Remove autorun")  { Enabled =  Installer.IsInstalled() };
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

            // start periodic trim (cheap; no allocations)
            _trimTimer.Tick += (_, _) => MemoryTrimmer.Trim();
            _trimTimer.Start();

            try { _icon.ShowBalloonTip(1200, "MicLockTray", $"Microphone volume locked to {Settings.TargetPercent}%.", ToolTipIcon.Info); } catch { }

            Application.ApplicationExit += (_, _) => { try { _icon.Visible = false; _icon.Dispose(); _enforcer.Dispose(); _trimTimer.Stop(); } catch { } };
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
                try { _icon.ShowBalloonTip(800, "MicLockTray", $"Resumed. Locking at {Settings.TargetPercent}% on change.", ToolTipIcon.Info); } catch { }
            }
        }

        private void PromptAndSetTarget()
        {
            using var dlg = new VolumePrompt(Settings.TargetPercent);
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                Settings.SetTarget(dlg.Value);
                _icon.Text = $"MicLockTray: target {Settings.TargetPercent}%";
                _enforcer.ForceToTarget();
                try { _icon.ShowBalloonTip(900, "MicLockTray", $"Target set to {Settings.TargetPercent}%.", ToolTipIcon.Info); } catch { }
            }
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

    // -------- CoreAudio interop + event-driven enforcer --------

    internal static class CoreAudio
    {
        public enum EDataFlow : int { eRender = 0, eCapture = 1, eAll = 2 }
        public enum ERole     : int { eConsole = 0, eMultimedia = 1, eCommunications = 2 }

        [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IMMDeviceEnumerator
        {
            [PreserveSig] int EnumAudioEndpoints(EDataFlow dataFlow, int dwStateMask, out IMMDeviceCollection ppDevices);
            [PreserveSig] int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);
            [PreserveSig] int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string pwstrId, out IMMDevice ppDevice);
            [PreserveSig] int RegisterEndpointNotificationCallback([MarshalAs(UnmanagedType.Interface)] IMMNotificationClient pClient);
            [PreserveSig] int UnregisterEndpointNotificationCallback([MarshalAs(UnmanagedType.Interface)] IMMNotificationClient pClient);
        }

        [ComImport, Guid("7991EEC9-7E89-4D85-8390-6C703CEC60C0"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IMMNotificationClient
        {
            void OnDeviceStateChanged([MarshalAs(UnmanagedType.LPWStr)] string pwstrDeviceId, uint dwNewState);
            void OnDeviceAdded([MarshalAs(UnmanagedType.LPWStr)] string pwstrDeviceId);
            void OnDeviceRemoved([MarshalAs(UnmanagedType.LPWStr)] string pwstrDeviceId);
            void OnDefaultDeviceChanged(EDataFlow flow, ERole role, [MarshalAs(UnmanagedType.LPWStr)] string pwstrDefaultDeviceId);
            void OnPropertyValueChanged([MarshalAs(UnmanagedType.LPWStr)] string pwstrDeviceId, PROPERTYKEY key);
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct PROPERTYKEY { public Guid fmtid; public uint pid; }

        [ComImport, Guid("0BD7A1BE-7A1A-44DB-8397-C0F926C399A4"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IMMDeviceCollection
        {
            [PreserveSig] int GetCount(out uint pcDevices);
            [PreserveSig] int Item(uint nDevice, out IMMDevice ppDevice);
        }

        [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        public class MMDeviceEnumerator { }

        [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IMMDevice
        {
            [PreserveSig] int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
            [PreserveSig] int OpenPropertyStore(int stgmAccess, out IntPtr ppProperties);
            [PreserveSig] int GetId(out IntPtr ppstrId);
            [PreserveSig] int GetState(out int pdwState);
        }

        [ComImport, Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
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
            [PreserveSig] int GetChannelVolumeLevelScalar(out float pfLevel);
            [PreserveSig] int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, Guid pguidEventContext);
            [PreserveSig] int GetMute(out bool pbMute);
            [PreserveSig] int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount);
            [PreserveSig] int VolumeStepUp(Guid pguidEventContext);
            [PreserveSig] int VolumeStepDown(Guid pguidEventContext);
            [PreserveSig] int QueryHardwareSupport(out uint pdwHardwareSupportMask);
            [PreserveSig] int GetVolumeRange(out float mindB, out float maxdB, out float incrementdB);
        }
    }

    [ComImport, Guid("657804FA-D6AD-4496-8A60-352752AF4F89"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioEndpointVolumeCallback
    {
        [PreserveSig] int OnNotify(IntPtr pNotify);
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct AUDIO_VOLUME_NOTIFICATION_DATA
    {
        public Guid guidEventContext;
        [MarshalAs(UnmanagedType.Bool)] public bool bMuted;
        public float fMasterVolume;
        public uint nChannels;
        public IntPtr afChannelVolumes; // float*
    }

    internal sealed class MicEnforcer : IDisposable
    {
        private sealed class Binding
        {
            public CoreAudio.IAudioEndpointVolume Ep = default!;
            public VolumeCallback Cb = default!;
        }

        [ComVisible(true)]
        private sealed class VolumeCallback : IAudioEndpointVolumeCallback
        {
            private readonly CoreAudio.IAudioEndpointVolume _ep;
            private readonly Func<float> _getTarget;
            private readonly Guid _ctx;

            public VolumeCallback(CoreAudio.IAudioEndpointVolume ep, Func<float> getTarget, Guid ctx)
            { _ep = ep; _getTarget = getTarget; _ctx = ctx; }

            public int OnNotify(IntPtr pNotify)
            {
                try
                {
                    var data = Marshal.PtrToStructure<AUDIO_VOLUME_NOTIFICATION_DATA>(pNotify);
                    if (data.guidEventContext == _ctx) return 0; // ignore self
                    float target = Math.Clamp(_getTarget(), 0f, 1f);
                    if (Math.Abs(data.fMasterVolume - target) > 0.005f)
                        _ = _ep.SetMasterVolumeLevelScalar(target, _ctx);
                }
                catch { }
                return 0;
            }
        }

        [ComVisible(true)]
        private sealed class NotificationClient : CoreAudio.IMMNotificationClient
        {
            private readonly MicEnforcer _owner;
            public NotificationClient(MicEnforcer owner) { _owner = owner; }
            public void OnDefaultDeviceChanged(CoreAudio.EDataFlow flow, CoreAudio.ERole role, string id)
            { if (flow == CoreAudio.EDataFlow.eCapture) { try { _owner.Rebind(role); _owner.ForceToTarget(); } catch { } } }
            public void OnDeviceStateChanged(string id, uint state) { }
            public void OnDeviceAdded(string id) { }
            public void OnDeviceRemoved(string id) { }
            public void OnPropertyValueChanged(string id, CoreAudio.PROPERTYKEY key) { }
        }

        private readonly Func<float> _getTargetScalar;
        private readonly Guid _eventCtx = Guid.NewGuid();
        private readonly CoreAudio.IMMDeviceEnumerator _enum = (CoreAudio.IMMDeviceEnumerator)new CoreAudio.MMDeviceEnumerator();
        private readonly NotificationClient _notify;
        private readonly Dictionary<CoreAudio.ERole, Binding> _bindings = new();
        private bool _enabled;
        private bool _disposed;

        public MicEnforcer(Func<float> getTargetScalar)
        {
            _getTargetScalar = getTargetScalar;
            _notify = new NotificationClient(this);
        }

        public bool IsEnabled => _enabled;

        public void Enable()
        {
            if (_enabled) return;
            try { _ = _enum.RegisterEndpointNotificationCallback(_notify); } catch { }
            foreach (CoreAudio.ERole role in new[] { CoreAudio.ERole.eConsole, CoreAudio.ERole.eMultimedia, CoreAudio.ERole.eCommunications })
                Bind(role);
            _enabled = true;
        }

        public void Disable()
        {
            if (!_enabled) return;
            foreach (var kv in _bindings)
            {
                try { kv.Value.Ep.UnregisterControlChangeNotify(kv.Value.Cb); } catch { }
                SafeReleaseCom(kv.Value.Ep);
            }
            _bindings.Clear();
            try { _ = _enum.UnregisterEndpointNotificationCallback(_notify); } catch { }
            _enabled = false;
        }

        public void ForceToTarget()
        {
            float target = Math.Clamp(_getTargetScalar(), 0f, 1f);
            foreach (var kv in _bindings)
            {
                try { _ = kv.Value.Ep.SetMasterVolumeLevelScalar(target, _eventCtx); } catch { }
            }
        }

        private void Bind(CoreAudio.ERole role)
        {
            if (_bindings.TryGetValue(role, out var old))
            {
                try { old.Ep.UnregisterControlChangeNotify(old.Cb); } catch { }
                SafeReleaseCom(old.Ep);
                _bindings.Remove(role);
            }

            CoreAudio.IMMDevice? dev = null;
            try
            {
                int hr = _enum.GetDefaultAudioEndpoint(CoreAudio.EDataFlow.eCapture, role, out dev);
                if (hr != 0 || dev is null) return;

                Guid iid = new("5CDF2C82-841E-4546-9722-0CF74078229A"); // IAudioEndpointVolume
                hr = dev.Activate(ref iid, 0x1 /*CLSCTX_INPROC_SERVER*/, IntPtr.Zero, out var obj);
                if (hr != 0 || obj is null) return;

                var ep = (CoreAudio.IAudioEndpointVolume)obj;
                var cb = new VolumeCallback(ep, _getTargetScalar, _eventCtx);
                _ = ep.RegisterControlChangeNotify(cb);

                _bindings[role] = new Binding { Ep = ep, Cb = cb };
            }
            catch { }
            finally
            {
                if (dev != null) SafeReleaseCom(dev);
            }
        }

        private void Rebind(CoreAudio.ERole role) => Bind(role);

        private static void SafeReleaseCom(object com)
        {
            try
            {
                if (Marshal.IsComObject(com))
                {
                    while (Marshal.ReleaseComObject(com) > 0) { }
                }
            }
            catch { }
        }

        public void Dispose()
        {
            if (_disposed) return;
            try { Disable(); } catch { }
            SafeReleaseCom(_enum);
            _disposed = true;
        }
    }
}
