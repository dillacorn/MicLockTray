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
    private const int RebindDelayMs = 100;

    private static readonly CoreAudio.ERole[] Roles =
    {
        CoreAudio.ERole.eConsole,
        CoreAudio.ERole.eMultimedia,
        CoreAudio.ERole.eCommunications
    };

    private sealed class Binding
    {
        public CoreAudio.IAudioEndpointVolume Endpoint = default!;
        public VolumeCallback Callback = default!;
    }

    [ComVisible(true)]
    private sealed class VolumeCallback : IAudioEndpointVolumeCallback
    {
        private readonly WeakReference<MicEnforcer> _owner;
        private readonly CoreAudio.ERole _role;
        private readonly Guid _eventContext;

        public VolumeCallback(MicEnforcer owner, CoreAudio.ERole role, Guid eventContext)
        {
            _owner = new WeakReference<MicEnforcer>(owner);
            _role = role;
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
                    owner.HandleVolumeChanged(_role, data.masterVolume);
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

        public void OnDefaultDeviceChanged(CoreAudio.EDataFlow flow, CoreAudio.ERole role, string? deviceId)
        {
            if (flow != CoreAudio.EDataFlow.eCapture) return;
            if (_owner.TryGetTarget(out var owner)) owner.ScheduleRebind(role);
        }

        public void OnDeviceStateChanged(string deviceId, uint newState) { }
        public void OnDeviceAdded(string deviceId) { }
        public void OnDeviceRemoved(string deviceId) { }
        public void OnPropertyValueChanged(string deviceId, CoreAudio.PropertyKey key) { }
    }

    private readonly object _gate = new();
    private readonly Func<float> _getTargetScalar;
    private readonly Guid _eventContext = Guid.NewGuid();
    private readonly CoreAudio.IMMDeviceEnumerator _enumerator =
        (CoreAudio.IMMDeviceEnumerator)new CoreAudio.MMDeviceEnumerator();
    private readonly NotificationClient _notificationClient;
    private readonly Dictionary<CoreAudio.ERole, Binding> _bindings = new();
    private readonly Dictionary<CoreAudio.ERole, ThreadingTimer> _rebindTimers = new();

    private bool _notificationRegistered;
    private bool _enabled;
    private bool _disposed;

    public MicEnforcer(Func<float> getTargetScalar)
    {
        _getTargetScalar = getTargetScalar;
        _notificationClient = new NotificationClient(this);
    }

    public bool IsEnabled
    {
        get
        {
            lock (_gate) return _enabled && !_disposed;
        }
    }

    public void Enable()
    {
        lock (_gate)
        {
            if (_disposed || _enabled) return;
            _enabled = true;
        }

        bool registered = false;
        try
        {
            registered = _enumerator.RegisterEndpointNotificationCallback(_notificationClient) == 0;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"MicLockTray endpoint notification registration failed: {ex}");
        }

        bool keepRegistration;
        lock (_gate)
        {
            keepRegistration = registered && !_disposed && _enabled;
            _notificationRegistered = keepRegistration;
        }

        if (registered && !keepRegistration)
        {
            try { _ = _enumerator.UnregisterEndpointNotificationCallback(_notificationClient); }
            catch { }
            return;
        }

        foreach (var role in Roles)
            RebindNow(role);
    }

    public void Disable()
    {
        Binding[] bindings;
        ThreadingTimer[] timers;
        bool unregisterNotifications;

        lock (_gate)
        {
            if (!_enabled && !_notificationRegistered && _bindings.Count == 0 && _rebindTimers.Count == 0)
                return;

            _enabled = false;
            bindings = _bindings.Values.ToArray();
            _bindings.Clear();
            timers = _rebindTimers.Values.ToArray();
            _rebindTimers.Clear();
            unregisterNotifications = _notificationRegistered;
            _notificationRegistered = false;
        }

        foreach (var timer in timers)
        {
            try { timer.Dispose(); } catch { }
        }

        foreach (var binding in bindings)
            UnregisterBinding(binding);

        if (!unregisterNotifications) return;

        try { _ = _enumerator.UnregisterEndpointNotificationCallback(_notificationClient); }
        catch (Exception ex)
        {
            Debug.WriteLine($"MicLockTray endpoint notification unregister failed: {ex}");
        }
    }

    public void ForceToTarget()
    {
        CoreAudio.IAudioEndpointVolume[] endpoints;

        lock (_gate)
        {
            if (_disposed || !_enabled) return;
            endpoints = _bindings.Values.Select(binding => binding.Endpoint).ToArray();
        }

        float target = Math.Clamp(_getTargetScalar(), 0f, 1f);
        foreach (var endpoint in endpoints)
            SetEndpointVolume(endpoint, target);
    }

    private void HandleVolumeChanged(CoreAudio.ERole role, float currentVolume)
    {
        float target = Math.Clamp(_getTargetScalar(), 0f, 1f);
        if (Math.Abs(currentVolume - target) <= 0.005f) return;

        CoreAudio.IAudioEndpointVolume? endpoint = null;
        lock (_gate)
        {
            if (!_disposed && _enabled && _bindings.TryGetValue(role, out var binding))
                endpoint = binding.Endpoint;
        }

        if (endpoint != null)
            SetEndpointVolume(endpoint, target);
    }

    private void ScheduleRebind(CoreAudio.ERole role)
    {
        lock (_gate)
        {
            if (_disposed || !_enabled) return;

            if (_rebindTimers.TryGetValue(role, out var timer))
            {
                timer.Change(RebindDelayMs, Timeout.Infinite);
                return;
            }

            _rebindTimers[role] = new ThreadingTimer(
                _ => RebindNow(role),
                null,
                RebindDelayMs,
                Timeout.Infinite
            );
        }
    }

    private void RebindNow(CoreAudio.ERole role)
    {
        lock (_gate)
        {
            if (_disposed || !_enabled) return;
        }

        var replacement = CreateBinding(role);
        if (replacement is null) return;

        Binding? previous = null;
        bool keepReplacement;

        lock (_gate)
        {
            keepReplacement = !_disposed && _enabled;
            if (keepReplacement)
            {
                _bindings.TryGetValue(role, out previous);
                _bindings[role] = replacement;
            }
        }

        if (!keepReplacement)
        {
            UnregisterBinding(replacement);
            return;
        }

        if (previous != null)
            UnregisterBinding(previous);

        SetEndpointVolume(replacement.Endpoint, Math.Clamp(_getTargetScalar(), 0f, 1f));
    }

    private Binding? CreateBinding(CoreAudio.ERole role)
    {
        try
        {
            int result = _enumerator.GetDefaultAudioEndpoint(
                CoreAudio.EDataFlow.eCapture,
                role,
                out var device
            );
            if (result != 0 || device is null) return null;

            Guid endpointVolumeInterface = new("5CDF2C82-841E-4546-9722-0CF74078229A");
            result = device.Activate(
                ref endpointVolumeInterface,
                0x1, // CLSCTX_INPROC_SERVER
                IntPtr.Zero,
                out var instance
            );
            if (result != 0 || instance is null) return null;

            var endpoint = (CoreAudio.IAudioEndpointVolume)instance;
            var callback = new VolumeCallback(this, role, _eventContext);
            result = endpoint.RegisterControlChangeNotify(callback);
            if (result != 0) return null;

            return new Binding { Endpoint = endpoint, Callback = callback };
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"MicLockTray failed to bind {role} capture endpoint: {ex}");
            return null;
        }
    }

    private void SetEndpointVolume(CoreAudio.IAudioEndpointVolume endpoint, float target)
    {
        try { _ = endpoint.SetMasterVolumeLevelScalar(target, _eventContext); }
        catch (Exception ex)
        {
            Debug.WriteLine($"MicLockTray failed to enforce microphone volume: {ex}");
        }
    }

    private static void UnregisterBinding(Binding binding)
    {
        try { _ = binding.Endpoint.UnregisterControlChangeNotify(binding.Callback); }
        catch (Exception ex)
        {
            Debug.WriteLine($"MicLockTray volume callback unregister failed: {ex}");
        }
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
