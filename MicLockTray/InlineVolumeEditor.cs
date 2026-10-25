using System;
using System.Drawing;
using System.Windows.Forms;

namespace MicLockTray;

// Embedded in the microphone's ToolStrip submenu; no separate volume dialog.
internal sealed class InlineVolumeEditor : UserControl
{
    private readonly NumericUpDown _input;

    public int Value => (int)_input.Value;

    public InlineVolumeEditor(int value, bool dark, string caption, Action<int> onChange)
    {
        AutoScaleMode = AutoScaleMode.Dpi;
        Size = new Size(236, 34);
        MinimumSize = new Size(236, 34);
        BackColor = dark ? Color.FromArgb(27, 27, 27) : Color.FromArgb(211, 216, 222);
        ForeColor = dark ? Color.FromArgb(233, 237, 241) : Color.FromArgb(32, 42, 54);
        Margin = Padding.Empty;

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            BackColor = BackColor,
            Padding = new Padding(7, 2, 7, 2)
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 64));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 36));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var label = new Label
        {
            AutoSize = true,
            Text = caption,
            ForeColor = ForeColor,
            Anchor = AnchorStyles.Left
        };

        _input = new NumericUpDown
        {
            Minimum = 1,
            Maximum = 100,
            Value = Math.Clamp(value, 1, 100),
            DecimalPlaces = 0,
            TextAlign = HorizontalAlignment.Right,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = dark ? Color.FromArgb(44, 44, 44) : Color.White,
            ForeColor = dark ? Color.White : Color.FromArgb(32, 42, 54)
        };

        layout.Controls.Add(label, 0, 0);
        layout.Controls.Add(_input, 1, 0);
        Controls.Add(layout);

        // Wire after initialization so constructing the menu never changes settings.
        _input.ValueChanged += (_, _) => onChange((int)_input.Value);
    }

    // Allows WinForms UI smoke tests to simulate a change without a visible tray.
    internal void SetValue(int value) => _input.Value = Math.Clamp(value, 1, 100);
}
