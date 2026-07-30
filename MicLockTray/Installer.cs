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
        var text = UiText.Current;

        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath, true);
            if (key is null)
                throw new InvalidOperationException(text.AutorunRegistryUnavailable);

            key.SetValue(RunValueName, $"\"{ExePath}\" --hidden", RegistryValueKind.String);
            MessageBox.Show(text.AutorunInstalled, text.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                UiText.Format(text.AutorunInstallFailedFormat, ex.Message),
                text.AppName,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
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
        var text = UiText.Current;

        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath, true);
            if (key is null)
                throw new InvalidOperationException(text.AutorunRegistryUnavailable);

            key.DeleteValue(RunValueName, false);
            MessageBox.Show(text.AutorunRemoved, text.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                UiText.Format(text.AutorunRemoveFailedFormat, ex.Message),
                text.AppName,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}
