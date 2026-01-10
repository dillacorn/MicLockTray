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

    internal static class Settings
    {
        private static readonly string Dir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MicLockTray");
        private static readonly string PathJson = System.IO.Path.Combine(Dir, "settings.json");

        public static int TargetPercent { get; set; } = 65;
        public static bool Enabled { get; set; } = true;
        public static bool StartWithWindows { get; set; } = false;
        public static int WorkingSetTrimSeconds { get; set; } = 300;

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
                StartWithWindows = dto.StartWithWindows;
                WorkingSetTrimSeconds = Clamp(dto.WorkingSetTrimSeconds, 15, 24 * 60 * 60);
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
                    Enabled = Enabled,
                    StartWithWindows = StartWithWindows,
                    WorkingSetTrimSeconds = Clamp(WorkingSetTrimSeconds, 15, 24 * 60 * 60),
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
            public bool StartWithWindows { get; set; } = false;
            public int WorkingSetTrimSeconds { get; set; } = 300;
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
                Settings.StartWithWindows = true;
                WriteRunKey(true);
                Settings.Save();
            }
            catch { }
        }

        public static void Uninstall()
        {
            try
            {
                Settings.StartWithWindows = false;
                WriteRunKey(false);
                Settings.Save();
            }
            catch { }
        }

        public static void SyncRunKey()
        {
            try
            {
                WriteRunKey(Settings.StartWithWindows);
            }
            catch { }
        }

        private static void WriteRunKey(bool enabled)
        {
            using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true) ??
                            Registry.CurrentUser.CreateSubKey(RunKeyPath, writable: true);

            if (key == null) return;

            if (!enabled)
            {
                try { key.DeleteValue(RunValueName, throwOnMissingValue: false); } catch { }
                return;
            }

            var exe = Application.ExecutablePath;
            key.SetValue(RunValueName, $"\"{exe}\"");
        }
    }

    internal sealed class TrayApp : ApplicationContext
    {
        private readonly NotifyIcon _icon;
        private readonly ContextMenuStrip _menu;

        private readonly ToolStripMenuItem _enabledItem;
        private readonly ToolStripMenuItem _setTargetItem;
        private readonly ToolStripMenuItem _startWithWindowsItem;
        private readonly ToolStripMenuItem _exitItem;

        private CoreAudioWatcher? _watcher;
        private readonly Timer _debounceTimer;
        private int _pendingTargetPercent;
        private bool _volumeCallbackArmed;

        private readonly Timer _trimTimer;

        public TrayApp()
        {
            _pendingTargetPercent = Settings.TargetPercent;

            _menu = new ContextMenuStrip();

            _enabledItem = new ToolStripMenuItem("Enabled") { Checked = Settings.Enabled, CheckOnClick = true };
            _enabledItem.CheckedChanged += (_, __) =>
            {
                Settings.Enabled = _enabledItem.Checked;
                Settings.Save();
                UpdateWatcherEnabledState();
            };

            _setTargetItem = new ToolStripMenuItem($"Set target volume… ({Settings.TargetPercent}%)");
            _setTargetItem.Click += (_, __) => PromptTarget();

            _startWithWindowsItem = new ToolStripMenuItem("Start with Windows")
            {
                Checked = Settings.StartWithWindows,
                CheckOnClick = true
            };
            _startWithWindowsItem.CheckedChanged += (_, __) =>
            {
                Settings.StartWithWindows = _startWithWindowsItem.Checked;
                Installer.SyncRunKey();
                Settings.Save();
            };

            _exitItem = new ToolStripMenuItem("Exit");
            _exitItem.Click += (_, __) => ExitThread();

            _menu.Items.AddRange(new ToolStripItem[]
            {
                _enabledItem,
                new ToolStripSeparator(),
                _setTargetItem,
                new ToolStripSeparator(),
                _startWithWindowsItem,
                new ToolStripSeparator(),
                _exitItem
            });

            _icon = new NotifyIcon
            {
                Visible = true,
                Text = "MicLockTray",
                Icon = System.Drawing.SystemIcons.Information,
                ContextMenuStrip = _menu
            };

            _icon.DoubleClick += (_, __) => PromptTarget();

            _debounceTimer = new Timer { Interval = 250 };
            _debounceTimer.Tick += (_, __) =>
            {
                _debounceTimer.Stop();
                ApplyTargetPercent(_pendingTargetPercent);
            };

            _trimTimer = new Timer();
            _trimTimer.Interval = Math.Max(15000, Settings.WorkingSetTrimSeconds * 1000);
            _trimTimer.Tick += (_, __) => TryTrimWorkingSet();
            _trimTimer.Start();

            UpdateWatcherEnabledState();
        }

        private void PromptTarget()
        {
            using var dlg = new VolumePrompt(Settings.TargetPercent);
            if (dlg.ShowDialog() != DialogResult.OK) return;

            var v = dlg.Value;
            _pendingTargetPercent = v;
            Settings.TargetPercent = v;
            Settings.Save();

            _setTargetItem.Text = $"Set target volume… ({Settings.TargetPercent}%)";

            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        private void ApplyTargetPercent(int pct)
        {
            try
            {
                if (_watcher == null) return;
                if (!Settings.Enabled) return;

                _volumeCallbackArmed = false;
                _watcher.SetMicVolumePercent(pct);
                _volumeCallbackArmed = true;
            }
            catch { }
        }

        private void UpdateWatcherEnabledState()
        {
            try
            {
                if (!Settings.Enabled)
                {
                    _watcher?.Dispose();
                    _watcher = null;
                    return;
                }

                if (_watcher != null) return;

                _watcher = new CoreAudioWatcher();
                _watcher.OnMicVolumeChanged += OnMicVolumeChanged;

                _volumeCallbackArmed = false;
                _watcher.SetMicVolumePercent(Settings.TargetPercent);
                _volumeCallbackArmed = true;
            }
            catch
            {
                _watcher?.Dispose();
                _watcher = null;
            }
        }

        private void OnMicVolumeChanged(int newPercent)
        {
            if (!Settings.Enabled) return;
            if (!_volumeCallbackArmed) return;

            if (newPercent == Settings.TargetPercent) return;

            // Snap back to the target if something changed it.
            _pendingTargetPercent = Settings.TargetPercent;
            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        private static void TryTrimWorkingSet()
        {
            try
            {
                using var p = Process.GetCurrentProcess();
                // This is a gentle hint; Windows may ignore it. No harm if it does.
                p.MinWorkingSet = new IntPtr(-1);
                p.MaxWorkingSet = new IntPtr(-1);
            }
            catch { }
        }

        protected override void ExitThreadCore()
        {
            try { _watcher?.Dispose(); } catch { }
            try { _debounceTimer.Stop(); _debounceTimer.Dispose(); } catch { }
            try { _trimTimer.Stop(); _trimTimer.Dispose(); } catch { }
            try { _menu.Dispose(); } catch { }
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

    internal sealed class CoreAudioWatcher : IDisposable
    {
        public event Action<int>? OnMicVolumeChanged;

        private IMMDeviceEnumerator? _enumerator;
        private IMMDevice? _device;
        private IAudioEndpointVolume? _endpoint;
        private AudioEndpointVolumeCallback? _cb;

        public CoreAudioWatcher()
        {
            _enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
            BindToDefaultCapture();
        }

        public void SetMicVolumePercent(int percent)
        {
            var ep = _endpoint;
            if (ep == null) return;

            float scalar = percent / 100f;
            ep.SetMasterVolumeLevelScalar(scalar, Guid.Empty);
        }

        private void BindToDefaultCapture()
        {
            _device = null;
            _endpoint = null;
            _cb = null;

            var en = _enumerator;
            if (en == null) return;

            en.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eCommunications, out var dev);
            _device = dev;

            var iid = typeof(IAudioEndpointVolume).GUID;
            dev.Activate(ref iid, 0, IntPtr.Zero, out object obj);
            _endpoint = (IAudioEndpointVolume)obj;

            _cb = new AudioEndpointVolumeCallback(this);
            _endpoint.RegisterControlChangeNotify(_cb);
        }

        internal void RaiseVolumeChanged(float volumeScalar)
        {
            int pct = (int)Math.Round(volumeScalar * 100.0);
            if (pct < 0) pct = 0;
            if (pct > 100) pct = 100;

            try { OnMicVolumeChanged?.Invoke(pct); } catch { }
        }

        public void Dispose()
        {
            try
            {
                if (_endpoint != null && _cb != null)
                {
                    try { _endpoint.UnregisterControlChangeNotify(_cb); } catch { }
                }
            }
            catch { }

            try { if (_endpoint != null) Marshal.ReleaseComObject(_endpoint); } catch { }
            try { if (_device != null) Marshal.ReleaseComObject(_device); } catch { }
            try { if (_enumerator != null) Marshal.ReleaseComObject(_enumerator); } catch { }

            _endpoint = null;
            _device = null;
            _enumerator = null;
            _cb = null;
        }

        private sealed class AudioEndpointVolumeCallback : IAudioEndpointVolumeCallback
        {
            private readonly CoreAudioWatcher _parent;
            public AudioEndpointVolumeCallback(CoreAudioWatcher parent) => _parent = parent;

            public void OnNotify(IntPtr pNotify)
            {
                try
                {
                    var data = Marshal.PtrToStructure<AUDIO_VOLUME_NOTIFICATION_DATA>(pNotify);
                    _parent.RaiseVolumeChanged(data.fMasterVolume);
                }
                catch { }
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct AUDIO_VOLUME_NOTIFICATION_DATA
        {
            public Guid guidEventContext;
            [MarshalAs(UnmanagedType.Bool)]
            public bool bMuted;
            public float fMasterVolume;
            public uint nChannels;
        }

        private sealed class MMDeviceEnumerator
        {
        }
    }

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
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceEnumerator
    {
        int NotImpl1();
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);
    }

    [ComImport]
    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDevice
    {
        int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
    }

    [ComImport]
    [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioEndpointVolume
    {
        int RegisterControlChangeNotify(IAudioEndpointVolumeCallback pNotify);
        int UnregisterControlChangeNotify(IAudioEndpointVolumeCallback pNotify);
        int GetChannelCount(out uint pnChannelCount);
        int SetMasterVolumeLevel(float fLevelDB, Guid pguidEventContext);
        int SetMasterVolumeLevelScalar(float fLevel, Guid pguidEventContext);
        int GetMasterVolumeLevel(out float pfLevelDB);
        int GetMasterVolumeLevelScalar(out float pfLevel);
        int SetChannelVolumeLevel(uint nChannel, float fLevelDB, Guid pguidEventContext);
        int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, Guid pguidEventContext);
        int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);
        int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel);
        int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, Guid pguidEventContext);
        int GetMute(out bool pbMute);
        int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount);
        int VolumeStepUp(Guid pguidEventContext);
        int VolumeStepDown(Guid pguidEventContext);
        int QueryHardwareSupport(out uint pdwHardwareSupportMask);
        int GetVolumeRange(out float pflVolumeMindB, out float pflVolumeMaxdB, out float pflVolumeIncrementdB);
    }

    [ComImport]
    [Guid("657804FA-D6AD-4496-8A60-352752AF4F89")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioEndpointVolumeCallback
    {
        void OnNotify(IntPtr pNotify);
    }
}
