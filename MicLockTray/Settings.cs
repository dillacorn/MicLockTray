using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.Json;

namespace MicLockTray;

internal static class Settings
{
    internal sealed class DevicePreference
    {
        public string name { get; set; } = "";
        public bool locked { get; set; }
        public int target_percent { get; set; } = 100;

        public DevicePreference Copy() => new()
        {
            name = name,
            locked = locked,
            target_percent = target_percent
        };
    }

    private sealed class ConfigDto
    {
        // Keep the legacy field and default behavior for existing installations.
        public int target_percent { get; set; } = 100;
        public bool follow_default { get; set; } = true;
        public bool notifications_enabled { get; set; } = true;
        public bool onboarding_shown { get; set; }
        public Dictionary<string, DevicePreference> microphones { get; set; } =
            new(StringComparer.OrdinalIgnoreCase);
    }

    private static readonly object Gate = new();
    private static readonly string DirectoryPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MicLockTray");
    private static readonly string ConfigPath = Path.Combine(DirectoryPath, "config.json");
    private static readonly string TempConfigPath = Path.Combine(DirectoryPath, "config.json.tmp");
    private static ConfigDto _config = new();

    public static int TargetPercent { get { lock (Gate) return _config.target_percent; } }
    public static bool FollowDefault { get { lock (Gate) return _config.follow_default; } }
    public static bool NotificationsEnabled { get { lock (Gate) return _config.notifications_enabled; } }
    public static bool ShowWelcome { get { lock (Gate) return !_config.onboarding_shown; } }

    public static void Load()
    {
        lock (Gate)
        {
            _config = new ConfigDto();
            try
            {
                if (!File.Exists(ConfigPath)) return;
                var saved = JsonSerializer.Deserialize<ConfigDto>(File.ReadAllText(ConfigPath));
                if (saved is null) return;
                saved.target_percent = Math.Clamp(saved.target_percent, 1, 100);
                var normalized = new Dictionary<string, DevicePreference>(StringComparer.OrdinalIgnoreCase);
                if (saved.microphones != null)
                {
                    foreach (var pair in saved.microphones)
                    {
                        if (string.IsNullOrWhiteSpace(pair.Key) || pair.Value == null) continue;
                        var preference = pair.Value.Copy();
                        preference.target_percent = Math.Clamp(preference.target_percent, 1, 100);
                        normalized[pair.Key] = preference;
                    }
                }
                saved.microphones = normalized;
                _config = saved;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"MicLockTray settings load failed: {ex}");
            }
        }
    }

    // An explicit per-device choice overrides the legacy default-device policy, including unlock.
    public static int? GetEffectiveTarget(string id, bool isDefault)
    {
        lock (Gate)
        {
            if (_config.microphones.TryGetValue(id, out var preference))
                return preference.locked ? preference.target_percent : null;
            return isDefault && _config.follow_default ? _config.target_percent : null;
        }
    }

    public static DevicePreference? GetPreference(string id)
    {
        lock (Gate)
            return _config.microphones.TryGetValue(id, out var preference) ? preference.Copy() : null;
    }

    public static Dictionary<string, DevicePreference> GetPreferences()
    {
        lock (Gate)
        {
            var result = new Dictionary<string, DevicePreference>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in _config.microphones) result[pair.Key] = pair.Value.Copy();
            return result;
        }
    }

    public static void SetTarget(int percent)
    {
        lock (Gate)
        {
            _config.target_percent = Math.Clamp(percent, 1, 100);
            SaveLocked();
        }
    }

    public static void SetFollowDefault(bool enabled)
    {
        lock (Gate) { _config.follow_default = enabled; SaveLocked(); }
    }

    public static void SetNotifications(bool enabled)
    {
        lock (Gate) { _config.notifications_enabled = enabled; SaveLocked(); }
    }

    public static void MarkWelcomeShown()
    {
        lock (Gate) { _config.onboarding_shown = true; SaveLocked(); }
    }

    public static void SetDeviceLock(string id, string name, bool locked, int initialPercent)
    {
        if (string.IsNullOrWhiteSpace(id)) return;
        lock (Gate)
        {
            if (!_config.microphones.TryGetValue(id, out var preference))
            {
                preference = new DevicePreference { target_percent = Math.Clamp(initialPercent, 1, 100) };
                _config.microphones[id] = preference;
            }
            preference.name = name;
            preference.locked = locked;
            SaveLocked();
        }
    }

    public static void SetDeviceTarget(string id, string name, int percent)
    {
        if (string.IsNullOrWhiteSpace(id)) return;
        lock (Gate)
        {
            if (!_config.microphones.TryGetValue(id, out var preference))
            {
                preference = new DevicePreference();
                _config.microphones[id] = preference;
            }
            preference.name = name;
            preference.target_percent = Math.Clamp(percent, 1, 100);
            preference.locked = true;
            SaveLocked();
        }
    }

    private static void SaveLocked()
    {
        try
        {
            Directory.CreateDirectory(DirectoryPath);
            var json = JsonSerializer.Serialize(_config, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(TempConfigPath, json);
            File.Move(TempConfigPath, ConfigPath, true);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"MicLockTray settings save failed: {ex}");
            try { File.Delete(TempConfigPath); } catch { }
        }
    }
}
