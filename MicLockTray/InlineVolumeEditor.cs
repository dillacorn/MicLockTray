using System;
using System.Drawing;
using System.Windows.Forms;

namespace MicLockTray;

// A single compact row matching the surrounding ToolStrip menu items.
internal sealed class InlineVolumeEditor : UserControl
{
    private readonly NumericUpDown _input;

    public int Value => (int)_input.Value;

    internal NumericUpDown Input => _input;

    public InlineVolumeEditor(int value, bool dark, string caption, Action<int> onChange)
    {
        AutoScaleMode = AutoScaleMode.Dpi;
        Size = new Size(236, 27);
        MinimumSize = Size;
        Margin = Padding.Empty;
        Padding = Padding.Empty;
        Font = SystemFonts.MenuFont;
        BackColor = dark ? Color.FromArgb(27, 27, 27) : Color.FromArgb(211, 216, 222);
        ForeColor = dark ? Color.FromArgb(233, 237, 241) : Color.FromArgb(32, 42, 54);

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = new Padding(12, 1, 10, 1),
            BackColor = BackColor
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 66));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var label = new Label
        {
            Text = caption,
            AutoSize = true,
            Margin = Padding.Empty,
            Anchor = AnchorStyles.Left,
            ForeColor = ForeColor,
            BackColor = BackColor
        };

        _input = new NumericUpDown
        {
            Minimum = 1,
            Maximum = 100,
            Value = Math.Clamp(value, 1, 100),
            DecimalPlaces = 0,
            TextAlign = HorizontalAlignment.Right,
            Margin = Padding.Empty,
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = dark ? Color.FromArgb(36, 36, 36) : Color.White,
            ForeColor = dark ? Color.FromArgb(233, 237, 241) : Color.FromArgb(32, 42, 54)
        };

        layout.Controls.Add(label, 0, 0);
        layout.Controls.Add(_input, 1, 0);
        Controls.Add(layout);

        // Initial menu construction must never write the microphone target.
        _input.ValueChanged += (_, _) => onChange((int)_input.Value);
    }

    internal void SetValue(int value) => _input.Value = Math.Clamp(value, 1, 100);
}
