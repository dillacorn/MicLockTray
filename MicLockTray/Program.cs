using System;
using System.Threading;
using System.Windows.Forms;

namespace MicLockTray;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        var arg = args.Length > 0 ? args[0].Trim().ToLowerInvariant() : string.Empty;
        if (arg == "--install") { Installer.Install(); return; }
        if (arg == "--uninstall") { Installer.Uninstall(); return; }

        using var mutex = new Mutex(true, $@"Local\MicLockTray-{Environment.UserName}", out bool createdNew);
        if (!createdNew) return;

        Settings.Load();

        ApplicationConfiguration.Initialize();
        Application.Run(new TrayApp());
    }
}
