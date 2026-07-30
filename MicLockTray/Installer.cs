using System;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Win32;

namespace MicLockTray;

internal static class Installer
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string RunValueName = "MicLockTray";
    private static string ExePath => Application.ExecutablePath;

    public static bool IsInstalled()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, false);
            var value = key?.GetValue(RunValueName) as string;
            return !string.IsNullOrWhiteSpace(value);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"MicLockTray autorun check failed: {ex}");
            return false;
        }
    }

    public static void Install()
    {
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath, true);
            if (key is null)
                throw new InvalidOperationException("The current-user autorun registry key could not be opened.");

            key.SetValue(RunValueName, $"\"{ExePath}\" --hidden", RegistryValueKind.String);
            MessageBox.Show("Autorun installed.", "MicLockTray", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to install autorun:\n{ex.Message}", "MicLockTray", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        try
        {
            var startInfo = new ProcessStartInfo(ExePath)
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            startInfo.ArgumentList.Add("--hidden");
            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"MicLockTray failed to start after autorun installation: {ex}");
        }
    }

    public static void Uninstall()
    {
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath, true);
            if (key is null)
                throw new InvalidOperationException("The current-user autorun registry key could not be opened.");

            key.DeleteValue(RunValueName, false);
            MessageBox.Show("Autorun removed.", "MicLockTray", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to remove autorun:\n{ex.Message}", "MicLockTray", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
