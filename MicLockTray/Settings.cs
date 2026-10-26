using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace MicLockTray;

internal static class Settings
{
    internal sealed record DeviceSetting(string Id, string Name, bool Locked, int TargetPercent);

    private sealed class DeviceDto
    {
        public string name { get; set; } = "";
        public bool locked { get; set; }
        public int target_percent { get; set; } = 100;
    }

    private sealed class ConfigDto
    {
        // Keep the original field for existing v1.5 installations.
        public int target_percent { get; set; } = 100;
        public bool default_locked { get; set; } = true;
        public bool notifications_enabled { get; set; } = true;
        public bool onboarding_shown { get; set; }
        public string menu_theme { get; set; } = "system";
        public Dictionary<string, DeviceDto> devices { get; set; } = new();
    }

    private static readonly object Gate = new();
    private static string _directoryPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MicLockTray");
    private static string ConfigPath => Path.Combine(_directoryPath, "config.json");
    private static string TempConfigPath => Path.Combine(_directoryPath, "config.json.tmp");

    private static int _targetPercent = 100;
    private static bool _defaultLocked = true;
    private static bool _notificationsEnabled = true;
    private static bool _onboardingShown;
    private static string _menuTheme = "system";
    private static Dictionary<string, DeviceSetting> _devices = new(StringComparer.OrdinalIgnoreCase);

    public static int TargetPercent { get { lock (Gate) return _targetPercent; } }
    public static bool DefaultLocked { get { lock (Gate) return _defaultLocked; } }
    public static bool NotificationsEnabled { get { lock (Gate) return _notificationsEnabled; } }
    public static bool OnboardingShown { get { lock (Gate) return _onboardingShown; } }
    public static string MenuTheme { get { lock (Gate) return _menuTheme; } }

    // Isolate the real per-user settings when running the standalone regression tests.
    internal static void UseTestDirectory(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A test directory is required.", nameof(path));
        lock (Gate)
        {
            _directoryPath = path;
            _targetPercent = 100;
            _defaultLocked = true;
            _notificationsEnabled = true;
            _onboardingShown = false;
            _menuTheme = "system";
            _devices = new Dictionary<string, DeviceSetting>(StringComparer.OrdinalIgnoreCase);
        }
    }

    public static void Load()
    {
        lock (Gate)
        {
            try
            {
                if (!File.Exists(ConfigPath)) return;
                var config = JsonSerializer.Deserialize<ConfigDto>(File.ReadAllText(ConfigPath));
                if (config is null) return;

                _targetPercent = Math.Clamp(config.target_percent, 1, 100);
                _defaultLocked = config.default_locked;
                _notificationsEnabled = config.notifications_enabled;
                _onboardingShown = config.onboarding_shown;
                _menuTheme = NormalizeTheme(config.menu_theme);
                _devices = new Dictionary<string, DeviceSetting>(StringComparer.OrdinalIgnoreCase);
                foreach (var entry in config.devices ?? new Dictionary<string, DeviceDto>())
                {
                    if (string.IsNullOrWhiteSpace(entry.Key) || entry.Value is null) continue;
                    _devices[entry.Key] = new DeviceSetting(
                        entry.Key, entry.Value.name ?? "", entry.Value.locked,
                        Math.Clamp(entry.Value.target_percent, 1, 100));
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"MicLockTray settings load failed: {ex}");
            }
        }
    }

    public static DeviceSetting? GetDevice(string id)
    {
        lock (Gate) return _devices.GetValueOrDefault(id);
    }

    public static DeviceSetting[] GetSavedDevices()
    {
        lock (Gate) return _devices.Values.OrderBy(device => device.Name, StringComparer.CurrentCultureIgnoreCase).ToArray();
    }

    // An explicit unlock overrides the legacy default-microphone fallback.
    public static int? EffectiveTarget(string id, bool isDefault)
    {
        lock (Gate)
        {
            if (_devices.TryGetValue(id, out var device))
                return device.Locked ? device.TargetPercent : null;
            return isDefault && _defaultLocked ? _targetPercent : null;
        }
    }

    public static void SetTarget(int percent)
    {
        lock (Gate)
        {
            _targetPercent = Math.Clamp(percent, 1, 100);
            Save();
        }
    }

    public static void SetDefaultLocked(bool value)
    {
        lock (Gate) { _defaultLocked = value; Save(); }
    }

    public static void SetNotificationsEnabled(bool value)
    {
        lock (Gate) { _notificationsEnabled = value; Save(); }
    }

    private static string NormalizeTheme(string? theme) =>
        theme is "light" or "dark" ? theme : "system";

    public static void SetMenuTheme(string theme)
    {
        lock (Gate) { _menuTheme = NormalizeTheme(theme); Save(); }
    }

    public static void SetOnboardingShown()
    {
        lock (Gate) { _onboardingShown = true; Save(); }
    }

    public static void SetDeviceLock(string id, string name, bool locked, int initialTarget)
    {
        lock (Gate)
        {
            var target = _devices.TryGetValue(id, out var previous)
                ? previous.TargetPercent : Math.Clamp(initialTarget, 1, 100);
            _devices[id] = new DeviceSetting(id, name, locked, target);
            Save();
        }
    }

    public static void SetDeviceTarget(string id, string name, int target)
    {
        lock (Gate)
        {
            _devices[id] = new DeviceSetting(id, name, true, Math.Clamp(target, 1, 100));
            Save();
        }
    }

    public static void ForgetDevice(string id)
    {
        lock (Gate)
        {
            if (_devices.Remove(id)) Save();
        }
    }

    private static void Save()
    {
        // Save always runs under Gate, so callbacks never observe partially updated configuration.
        try
        {
            Directory.CreateDirectory(_directoryPath);
            var config = new ConfigDto
            {
                target_percent = _targetPercent,
                default_locked = _defaultLocked,
                notifications_enabled = _notificationsEnabled,
                onboarding_shown = _onboardingShown,
                menu_theme = _menuTheme,
                devices = _devices.ToDictionary(
                    pair => pair.Key,
                    pair => new DeviceDto
                    {
                        name = pair.Value.Name,
                        locked = pair.Value.Locked,
                        target_percent = pair.Value.TargetPercent
                    },
                    StringComparer.OrdinalIgnoreCase)
            };

            File.WriteAllText(TempConfigPath,
                JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true }));
            File.Move(TempConfigPath, ConfigPath, true);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"MicLockTray settings save failed: {ex}");
            try { File.Delete(TempConfigPath); } catch { }
        }
    }
}
