using System;
using System.Drawing;

namespace MicLockTray;

// Pure screen-space placement used by the tray menu and regression tests.
internal static class MenuPlacement
{
    internal static Point Calculate(Point cursor, Rectangle workArea, Size menuSize,
        int bottomMargin = 24)
    {
        if (workArea.Width <= 0 || workArea.Height <= 0)
            return cursor;

        const int edge = 8;
        int left = workArea.Left + edge;
        int right = Math.Max(left, workArea.Right - edge - Math.Max(0, menuSize.Width));
        int top = workArea.Top + edge;
        int bottom = Math.Max(top, workArea.Bottom - Math.Max(0, bottomMargin)
            - Math.Max(0, menuSize.Height));

        // Prefer the menu entirely above a bottom-docked tray icon. Clamping
        // keeps it on-screen for top/side taskbars and secondary monitors too.
        int x = Math.Clamp(cursor.X - menuSize.Width + 16, left, right);
        int y = Math.Clamp(cursor.Y - menuSize.Height - 12, top, bottom);
        return new Point(x, y);
    }
}
