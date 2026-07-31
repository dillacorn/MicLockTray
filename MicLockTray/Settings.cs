using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Threading;

namespace MicLockTray;

internal static class Settings
{
    private static readonly string DirectoryPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MicLockTray"
    );

    private static readonly string ConfigPath = Path.Combine(DirectoryPath, "config.json");
    private static readonly string TempConfigPath = Path.Combine(DirectoryPath, "config.json.tmp");
    private static int _targetPercent = 100;

    public static int TargetPercent => Volatile.Read(ref _targetPercent);

    public static void Load()
    {
        try
        {
            if (!File.Exists(ConfigPath)) return;

            var json = File.ReadAllText(ConfigPath);
            var config = JsonSerializer.Deserialize<ConfigDto>(json);
            if (config != null)
                Volatile.Write(ref _targetPercent, Math.Clamp(config.target_percent, 1, 100));
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"MicLockTray settings load failed: {ex}");
        }
    }

    public static void SetTarget(int percent)
    {
        percent = Math.Clamp(percent, 1, 100);
        Volatile.Write(ref _targetPercent, percent);
        Save(percent);
    }

    private static void Save(int percent)
    {
        try
        {
            Directory.CreateDirectory(DirectoryPath);
            var config = new ConfigDto { target_percent = percent };
            var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });

            File.WriteAllText(TempConfigPath, json);
            File.Move(TempConfigPath, ConfigPath, true);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"MicLockTray settings save failed: {ex}");
            try { File.Delete(TempConfigPath); } catch { }
        }
    }

    private sealed class ConfigDto
    {
        public int target_percent { get; set; } = 100;
    }
}
