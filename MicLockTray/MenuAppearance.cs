using System;
using Microsoft.Win32;

namespace MicLockTray;

internal static class MenuAppearance
{
    internal static bool IsDark(string preference, bool windowsDark) =>
        preference switch
        {
            "dark" => true,
            "light" => false,
            _ => windowsDark
        };

    internal static bool WindowsDarkMode()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            // AppsUseLightTheme: 0 = dark apps, 1 = light apps. Missing => light.
            return key?.GetValue("AppsUseLightTheme") is int value && value == 0;
        }
        catch { return false; }
    }
}
