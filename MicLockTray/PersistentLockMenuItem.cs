using System;
using System.Drawing;
using System.Windows.Forms;

namespace MicLockTray;

// A hosted button avoids ToolStripDropDown.OnItemClicked's automatic close.
// Only microphone lock toggles use this; normal menu commands remain native.
internal sealed class PersistentLockMenuItem : ToolStripControlHost
{
    private readonly Button _button;

    internal PersistentLockMenuItem(string caption, bool dark)
        : base(CreateButton(caption, dark))
    {
        _button = (Button)Control;
        AutoSize = false;
        Size = _button.Size;
        Margin = Padding.Empty;
        Padding = Padding.Empty;
    }

    private static Button CreateButton(string caption, bool dark)
    {
        var background = dark ? Color.FromArgb(27, 27, 27) : Color.FromArgb(211, 216, 222);
        var foreground = dark ? Color.FromArgb(233, 237, 241) : Color.FromArgb(32, 42, 54);
        var hover = dark ? Color.FromArgb(53, 55, 58) : Color.FromArgb(170, 184, 198);

        var button = new Button
        {
            Text = caption,
            Font = new Font(SystemFonts.MenuFont, FontStyle.Bold),
            Size = new Size(236, 26),
            BackColor = background,
            ForeColor = foreground,
            TextAlign = ContentAlignment.MiddleLeft,
            FlatStyle = FlatStyle.Flat,
            Margin = Padding.Empty,
            Padding = new Padding(5, 0, 0, 0),
            UseVisualStyleBackColor = false,
            AccessibleName = caption
        };
        button.FlatAppearance.BorderSize = 0;
        button.FlatAppearance.MouseOverBackColor = hover;
        button.FlatAppearance.MouseDownBackColor = hover;
        return button;
    }

    public override string Text
    {
        get => _button.Text;
        set
        {
            _button.Text = value;
            _button.AccessibleName = value;
        }
    }

    // Direct hosted-button event. No ToolStripDropDownItem ItemClicked dismissal.
    internal void TriggerClick() => _button.PerformClick();

    internal bool UsesHostedButton => Control is Button;
}
