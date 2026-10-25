using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using ThreadingTimer = System.Threading.Timer;

namespace MicLockTray;

internal sealed class MicEnforcer : IDisposable
{
    private const int RefreshDelayMs = 100;
    private static readonly CoreAudio.ERole[] Roles =
    {
        CoreAudio.ERole.eConsole,
        CoreAudio.ERole.eMultimedia,
        CoreAudio.ERole.eCommunications
    };

    internal sealed record DeviceInfo(string Id, string Name, int CurrentPercent, bool IsDefault);

    private sealed class Binding
    {
        public string Id = "";
        public string Name = "";
        public CoreAudio.IAudioEndpointVolume Endpoint = default!;
        public VolumeCallback Callback = default!;
    }

    [ComVisible(true)]
    private sealed class VolumeCallback : IAudioEndpointVolumeCallback
    {
        private readonly WeakReference<MicEnforcer> _owner;
        private readonly string _id;
        private readonly Guid _eventContext;

        public VolumeCallback(MicEnforcer owner, string id, Guid eventContext)
        {
            _owner = new WeakReference<MicEnforcer>(owner);
            _id = id;
            _eventContext = eventContext;
        }

        public int OnNotify(IntPtr notificationData)
        {
            try
            {
                if (notificationData == IntPtr.Zero) return 0;
                var data = Marshal.PtrToStructure<AudioVolumeNotificationData>(notificationData);
                if (data.eventContext == _eventContext) return 0;
                if (_owner.TryGetTarget(out var owner))
                    owner.HandleVolumeChanged(_id, data.masterVolume);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"MicLockTray volume callback failed: {ex}");
            }
            return 0;
        }
    }

    [ComVisible(true)]
    private sealed class NotificationClient : CoreAudio.IMMNotificationClient
    {
        private readonly WeakReference<MicEnforcer> _owner;

        public NotificationClient(MicEnforcer owner)
        {
            _owner = new WeakReference<MicEnforcer>(owner);
        }

        public void OnDefaultDeviceChanged(CoreAudio.EDataFlow flow, CoreAudio.ERole role, string? id)
        {
            if (flow == CoreAudio.EDataFlow.eCapture && _owner.TryGetTarget(out var owner))
                owner.ScheduleRefresh();
        }

        public void OnDeviceStateChanged(string id, uint state)
        {
            if (_owner.TryGetTarget(out var owner)) owner.ScheduleRefresh();
        }

        public void OnDeviceAdded(string id)
        {
            if (_owner.TryGetTarget(out var owner)) owner.ScheduleRefresh();
        }

        public void OnDeviceRemoved(string id)
        {
            if (_owner.TryGetTarget(out var owner)) owner.ScheduleRefresh();
        }

        public void OnPropertyValueChanged(string id, CoreAudio.PropertyKey key)
        {
            if (_owner.TryGetTarget(out var owner)) owner.ScheduleRefresh();
        }
    }

    private readonly object _gate = new();
    private readonly object _refreshGate = new();
    private readonly Guid _eventContext = Guid.NewGuid();
    private readonly CoreAudio.IMMDeviceEnumerator _enumerator =
        (CoreAudio.IMMDeviceEnumerator)new CoreAudio.MMDeviceEnumerator();
    private readonly NotificationClient _notifications;
    private readonly Dictionary<string, Binding> _bindings = new(StringComparer.OrdinalIgnoreCase);
    private HashSet<string> _defaultIds = new(StringComparer.OrdinalIgnoreCase);
    private ThreadingTimer? _refreshTimer;
    private bool _registered;
    private bool _enabled;
    private bool _disposed;
    private long _generation;

    public MicEnforcer()
    {
        _notifications = new NotificationClient(this);
    }

    public bool IsEnabled { get { lock (_gate) return _enabled && !_disposed; } }

    public void Enable()
    {
        lock (_gate)
        {
            if (_disposed || _enabled) return;
            _enabled = true;
            _generation++;
        }

        bool registered = false;
        try { registered = _enumerator.RegisterEndpointNotificationCallback(_notifications) == 0; }
        catch (Exception ex) { Debug.WriteLine($"MicLockTray notification registration failed: {ex}"); }

        lock (_gate)
        {
            if (!_enabled || _disposed)
            {
                if (registered)
                {
                    try { _ = _enumerator.UnregisterEndpointNotificationCallback(_notifications); }
                    catch { }
                }
                return;
            }
            _registered = registered;
        }

        RefreshNow();
    }

    public void Disable()
    {
        Binding[] bindings;
        ThreadingTimer? timer;
        bool unregister;
        lock (_gate)
        {
            if (!_enabled && !_registered && _bindings.Count == 0 && _refreshTimer is null) return;
            _enabled = false;
            _generation++;
            bindings = _bindings.Values.ToArray();
            _bindings.Clear();
            _defaultIds.Clear();
            timer = _refreshTimer;
            _refreshTimer = null;
            unregister = _registered;
            _registered = false;
        }

        try { timer?.Dispose(); } catch { }
        foreach (var binding in bindings) UnregisterBinding(binding);

        if (unregister)
        {
            try { _ = _enumerator.UnregisterEndpointNotificationCallback(_notifications); }
            catch (Exception ex) { Debug.WriteLine($"MicLockTray notification unregister failed: {ex}"); }
        }
    }

    public DeviceInfo[] GetDevices()
    {
        (Binding binding, bool isDefault)[] bindings;
        lock (_gate)
        {
            bindings = _bindings.Values.Select(binding =>
                (binding, _defaultIds.Contains(binding.Id))).ToArray();
        }

        return bindings.Select(item =>
        {
            int volume = Settings.EffectiveTarget(item.binding.Id, item.isDefault) ?? Settings.TargetPercent;
            try
            {
                if (item.binding.Endpoint.GetMasterVolumeLevelScalar(out float scalar) == 0)
                    volume = Math.Clamp((int)Math.Round(scalar * 100), 0, 100);
            }
            catch (Exception ex) { Debug.WriteLine($"MicLockTray device volume read failed: {ex}"); }

            return new DeviceInfo(item.binding.Id, item.binding.Name, volume, item.isDefault);
        }).OrderByDescending(device => device.IsDefault)
          .ThenBy(device => device.Name, StringComparer.CurrentCultureIgnoreCase).ToArray();
    }

    public void ForceToTarget()
    {
        (CoreAudio.IAudioEndpointVolume endpoint, string id, bool isDefault)[] targets;
        lock (_gate)
        {
            if (!_enabled || _disposed) return;
            targets = _bindings.Values.Select(binding =>
                (binding.Endpoint, binding.Id, _defaultIds.Contains(binding.Id))).ToArray();
        }

        foreach (var target in targets)
        {
            int? percent = Settings.EffectiveTarget(target.id, target.isDefault);
            if (percent is not null)
                SetEndpointVolume(target.endpoint, percent.Value / 100f);
        }
    }

    private void HandleVolumeChanged(string id, float current)
    {
        CoreAudio.IAudioEndpointVolume? endpoint;
        bool isDefault;
        lock (_gate)
        {
            if (_disposed || !_enabled || !_bindings.TryGetValue(id, out var binding)) return;
            endpoint = binding.Endpoint;
            isDefault = _defaultIds.Contains(id);
        }

        int? percent = Settings.EffectiveTarget(id, isDefault);
        if (percent is null) return;
        float target = percent.Value / 100f;
        if (Math.Abs(current - target) > 0.005f)
            SetEndpointVolume(endpoint, target);
    }

    private void ScheduleRefresh()
    {
        lock (_gate)
        {
            if (_disposed || !_enabled) return;
            if (_refreshTimer is null)
                _refreshTimer = new ThreadingTimer(_ => RefreshNow(), null, RefreshDelayMs, Timeout.Infinite);
            else
                _refreshTimer.Change(RefreshDelayMs, Timeout.Infinite);
        }
    }

    private void RefreshNow()
    {
        // Device notifications may arrive concurrently or while the user pauses enforcement.
        lock (_refreshGate)
        {
            long generation;
            HashSet<string> knownIds;
            lock (_gate)
            {
                if (_disposed || !_enabled) return;
                generation = _generation;
                knownIds = _bindings.Keys.ToHashSet(StringComparer.OrdinalIgnoreCase);
            }

            var devices = new Dictionary<string, (CoreAudio.IMMDevice device, string name)>(
                StringComparer.OrdinalIgnoreCase);
            var defaults = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (_enumerator.EnumAudioEndpoints(CoreAudio.EDataFlow.eCapture,
                        CoreAudio.ActiveDeviceState, out var collection) != 0 || collection is null)
                    return;
                if (collection.GetCount(out uint count) != 0) return;
                for (uint i = 0; i < count; i++)
                {
                    if (collection.Item(i, out var device) != 0 || device is null) continue;
                    string? id = CoreAudio.GetDeviceId(device);
                    if (string.IsNullOrEmpty(id)) continue;
                    devices[id] = (device, CoreAudio.GetFriendlyName(device, id));
                }

                foreach (var role in Roles)
                {
                    if (_enumerator.GetDefaultAudioEndpoint(CoreAudio.EDataFlow.eCapture,
                            role, out var device) != 0 || device is null) continue;
                    string? id = CoreAudio.GetDeviceId(device);
                    if (id is not null) defaults.Add(id);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"MicLockTray capture enumeration failed: {ex}");
                return; // Never discard existing bindings on an unsuccessful scan.
            }

            var additions = new Dictionary<string, Binding>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in devices)
            {
                if (knownIds.Contains(pair.Key)) continue;
                var binding = CreateBinding(pair.Key, pair.Value.name, pair.Value.device);
                if (binding is not null) additions[pair.Key] = binding;
            }

            Binding[] removed = Array.Empty<Binding>();
            bool accepted;
            lock (_gate)
            {
                accepted = !_disposed && _enabled && generation == _generation;
                if (accepted)
                {
                    removed = _bindings.Where(pair => !devices.ContainsKey(pair.Key))
                        .Select(pair => pair.Value).ToArray();
                    foreach (var binding in removed) _bindings.Remove(binding.Id);
                    foreach (var pair in additions) _bindings[pair.Key] = pair.Value;
                    foreach (var pair in devices)
                    {
                        if (_bindings.TryGetValue(pair.Key, out var existing))
                            existing.Name = pair.Value.name;
                    }
                    _defaultIds = defaults;
                }
            }

            if (!accepted)
            {
                foreach (var binding in additions.Values) UnregisterBinding(binding);
                return;
            }
            foreach (var binding in removed) UnregisterBinding(binding);
            ForceToTarget();
        }
    }

    private Binding? CreateBinding(string id, string name, CoreAudio.IMMDevice device)
    {
        try
        {
            Guid guid = CoreAudio.EndpointVolumeGuid;
            if (device.Activate(ref guid, 0x1, IntPtr.Zero, out var instance) != 0 ||
                instance is not CoreAudio.IAudioEndpointVolume endpoint)
                return null;

            var callback = new VolumeCallback(this, id, _eventContext);
            if (endpoint.RegisterControlChangeNotify(callback) != 0) return null;
            return new Binding { Id = id, Name = name, Endpoint = endpoint, Callback = callback };
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"MicLockTray failed to bind capture endpoint: {ex}");
            return null;
        }
    }

    private void SetEndpointVolume(CoreAudio.IAudioEndpointVolume endpoint, float target)
    {
        try
        {
            if (endpoint.GetMasterVolumeLevelScalar(out float current) == 0 &&
                Math.Abs(current - target) <= 0.005f) return;
            int hr = endpoint.SetMasterVolumeLevelScalar(target, _eventContext);
            if (hr != 0) Debug.WriteLine($"MicLockTray endpoint volume write failed: 0x{hr:X8}");
        }
        catch (Exception ex) { Debug.WriteLine($"MicLockTray endpoint volume write failed: {ex}"); }
    }

    private static void UnregisterBinding(Binding binding)
    {
        try { _ = binding.Endpoint.UnregisterControlChangeNotify(binding.Callback); }
        catch (Exception ex) { Debug.WriteLine($"MicLockTray callback unregister failed: {ex}"); }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed) return;
            _disposed = true;
        }
        Disable();
    }
}
