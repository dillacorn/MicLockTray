using System.Windows.Forms;

namespace MicLockTray;

// Lock toggles update in place without dismissing the microphone submenu.
// Other commands retain the normal menu-closing behavior.
internal sealed class PersistentLockMenuItem : ToolStripMenuItem
{
    protected override bool DismissWhenClicked => false;

    internal bool KeepsMenuOpen => !DismissWhenClicked;
}
