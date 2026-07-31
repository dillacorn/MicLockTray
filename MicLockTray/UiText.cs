using System;
using System.Globalization;

namespace MicLockTray;

internal static class UiText
{
    internal sealed class Strings
    {
        public string AppName { get; init; } = "";
        public string TrayTooltipFormat { get; init; } = "";
        public string PauseEnforcement { get; init; } = "";
        public string ResumeEnforcement { get; init; } = "";
        public string SetTargetVolume { get; init; } = "";
        public string InstallAutorun { get; init; } = "";
        public string RemoveAutorun { get; init; } = "";
        public string Exit { get; init; } = "";
        public string StartupBalloonFormat { get; init; } = "";
        public string PausedBalloon { get; init; } = "";
        public string ResumedBalloonFormat { get; init; } = "";
        public string TargetSetBalloonFormat { get; init; } = "";
        public string VolumeDialogTitle { get; init; } = "";
        public string VolumeLabel { get; init; } = "";
        public string Ok { get; init; } = "";
        public string Cancel { get; init; } = "";
        public string AutorunInstalled { get; init; } = "";
        public string AutorunRemoved { get; init; } = "";
        public string AutorunInstallFailedFormat { get; init; } = "";
        public string AutorunRemoveFailedFormat { get; init; } = "";
        public string AutorunRegistryUnavailable { get; init; } = "";
    }

    private static readonly Strings English = new()
    {
        AppName = "MicLockTray",
        TrayTooltipFormat = "MicLockTray: target {0}%",
        PauseEnforcement = "Pause enforcement",
        ResumeEnforcement = "Resume enforcement",
        SetTargetVolume = "Set target volume…",
        InstallAutorun = "Install autorun",
        RemoveAutorun = "Remove autorun",
        Exit = "Exit",
        StartupBalloonFormat = "Microphone volume locked to {0}%.",
        PausedBalloon = "Paused.",
        ResumedBalloonFormat = "Resumed. Locking at {0}% on change.",
        TargetSetBalloonFormat = "Target set to {0}%.",
        VolumeDialogTitle = "Set target volume (%)",
        VolumeLabel = "Volume (1–100):",
        Ok = "OK",
        Cancel = "Cancel",
        AutorunInstalled = "Autorun installed.",
        AutorunRemoved = "Autorun removed.",
        AutorunInstallFailedFormat = "Failed to install autorun:\n{0}",
        AutorunRemoveFailedFormat = "Failed to remove autorun:\n{0}",
        AutorunRegistryUnavailable = "The current-user autorun registry key could not be opened."
    };

    // Simplified Chinese translation adapted from BaoZiFly-233/MicLockTray.
    private static readonly Strings SimplifiedChinese = new()
    {
        AppName = "MicLockTray",
        TrayTooltipFormat = "MicLockTray：目标 {0}%",
        PauseEnforcement = "暂停锁定",
        ResumeEnforcement = "恢复锁定",
        SetTargetVolume = "设置目标音量…",
        InstallAutorun = "安装开机自启",
        RemoveAutorun = "移除开机自启",
        Exit = "退出",
        StartupBalloonFormat = "麦克风音量已锁定到 {0}%。",
        PausedBalloon = "已暂停。",
        ResumedBalloonFormat = "已恢复。检测到变化时将锁定到 {0}%。",
        TargetSetBalloonFormat = "目标音量已设置为 {0}%。",
        VolumeDialogTitle = "设置目标音量 (%)",
        VolumeLabel = "音量 (1–100)：",
        Ok = "确定",
        Cancel = "取消",
        AutorunInstalled = "开机自启已安装。",
        AutorunRemoved = "开机自启已移除。",
        AutorunInstallFailedFormat = "安装开机自启失败：\n{0}",
        AutorunRemoveFailedFormat = "移除开机自启失败：\n{0}",
        AutorunRegistryUnavailable = "无法打开当前用户的开机自启注册表项。"
    };

    public static Strings Current { get; } = Resolve(CultureInfo.CurrentUICulture);

    public static string Format(string format, params object[] args) =>
        string.Format(CultureInfo.CurrentCulture, format, args);

    internal static Strings Resolve(CultureInfo uiCulture) =>
        IsSimplifiedChinese(uiCulture) ? SimplifiedChinese : English;

    private static bool IsSimplifiedChinese(CultureInfo culture)
    {
        for (CultureInfo current = culture;
             !string.IsNullOrEmpty(current.Name);
             current = current.Parent)
        {
            string name = current.Name;
            if (name.StartsWith("zh-Hans", StringComparison.OrdinalIgnoreCase) ||
                name.Equals("zh-CN", StringComparison.OrdinalIgnoreCase) ||
                name.Equals("zh-SG", StringComparison.OrdinalIgnoreCase) ||
                name.Equals("zh-MY", StringComparison.OrdinalIgnoreCase) ||
                name.Equals("zh-CHS", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }
}
