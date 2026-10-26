using System;
using System.Drawing;
using System.Windows.Forms;

namespace MicLockTray;

// Compact inline target editor, with a plain numeric input and unobtrusive +/-
// buttons instead of the default WinForms spinner. No separate dialog.
internal sealed class InlineVolumeEditor : UserControl
{
    private readonly TextBox _input;
    private readonly Action<int> _onChange;
    private int _value;

    public int Value => _value;
    internal TextBox Input => _input;

    public InlineVolumeEditor(int value, bool dark, string caption, Action<int> onChange)
    {
        _value = Math.Clamp(value, 1, 100);
        _onChange = onChange ?? throw new ArgumentNullException(nameof(onChange));

        AutoScaleMode = AutoScaleMode.Dpi;
        Size = new Size(236, 27);
        MinimumSize = Size;
        Margin = Padding.Empty;
        Padding = Padding.Empty;
        Font = SystemFonts.MenuFont;
        BackColor = dark ? Color.FromArgb(27, 27, 27) : Color.FromArgb(211, 216, 222);
        ForeColor = dark ? Color.FromArgb(233, 237, 241) : Color.FromArgb(32, 42, 54);

        var editorBackground = dark ? Color.FromArgb(38, 38, 38) : Color.FromArgb(244, 246, 248);
        var editorForeground = dark ? Color.FromArgb(235, 235, 235) : Color.FromArgb(32, 42, 54);

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = new Padding(2, 2, 7, 2),
            BackColor = BackColor
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 88));
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

        var controlBorder = new Panel
        {
            Dock = DockStyle.Fill,
            Margin = Padding.Empty,
            Padding = new Padding(1),
            BackColor = dark ? Color.FromArgb(70, 70, 70) : Color.FromArgb(147, 155, 164)
        };
        var controls = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
            BackColor = editorBackground
        };
        controls.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 21));
        controls.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        controls.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 21));
        controls.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        _input = new TextBox
        {
            BorderStyle = BorderStyle.None,
            TextAlign = HorizontalAlignment.Center,
            MaxLength = 3,
            Text = _value.ToString(),
            ForeColor = editorForeground,
            BackColor = editorBackground,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            // The native borderless EDIT text sits slightly above centered
            // WinForms button glyphs; offset its row placement by ~1–2 pixels.
            Margin = new Padding(0, 3, 0, 0)
        };
        _input.KeyPress += (_, e) =>
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
                e.Handled = true;
        };
        _input.KeyDown += (_, e) =>
        {
            switch (e.KeyCode)
            {
                case Keys.Enter:
                    CommitText();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
                case Keys.Up:
                    Step(1);
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
                case Keys.Down:
                    Step(-1);
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
                case Keys.Escape:
                    _input.Text = _value.ToString();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
            }
        };
        _input.Leave += (_, _) => CommitText();

        var minus = NewStepButton("−", editorBackground, editorForeground);
        var plus = NewStepButton("+", editorBackground, editorForeground);
        minus.Click += (_, _) => { CommitText(); Step(-1); };
        plus.Click += (_, _) => { CommitText(); Step(1); };

        controls.Controls.Add(minus, 0, 0);
        controls.Controls.Add(_input, 1, 0);
        controls.Controls.Add(plus, 2, 0);
        controlBorder.Controls.Add(controls);
        layout.Controls.Add(label, 0, 0);
        layout.Controls.Add(controlBorder, 1, 0);
        Controls.Add(layout);
    }

    private static Button NewStepButton(string text, Color background, Color foreground)
    {
        var button = new Button
        {
            Text = text,
            Font = SystemFonts.MenuFont,
            Dock = DockStyle.Fill,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
            FlatStyle = FlatStyle.Flat,
            BackColor = background,
            ForeColor = foreground,
            TabStop = false,
            UseVisualStyleBackColor = false
        };
        button.FlatAppearance.BorderSize = 0;
        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(
            background.R < 100 ? 63 : 221,
            background.G < 100 ? 63 : 226,
            background.B < 100 ? 63 : 230);
        return button;
    }

    internal void SetValue(int value)
    {
        int next = Math.Clamp(value, 1, 100);
        _input.Text = next.ToString();
        if (next == _value) return;
        _value = next;
        _onChange(next);
    }

    internal void Step(int change) => SetValue(_value + change);

    internal void CommitText()
    {
        if (int.TryParse(_input.Text, out int typed))
            SetValue(typed);
        else
            _input.Text = _value.ToString();
    }
}
