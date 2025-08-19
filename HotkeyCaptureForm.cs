public class HotkeyCaptureForm : Form
{
    public string HotkeyString { get; private set; } = "";

    private Keys currentModifiers = Keys.None;
    private Keys currentKey = Keys.None;

    public HotkeyCaptureForm()
    {
        this.KeyPreview = true;
        this.Text = "Press Hotkey";
        this.Width = 320;   // Set your desired width
        this.Height = 140;  // Set your desired height
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.StartPosition = FormStartPosition.CenterParent;

        var label = new Label { Text = "Press your hotkey...", Dock = DockStyle.Top, Height = 30, TextAlign = ContentAlignment.MiddleCenter };
        var okButton = new Button { Text = "OK", Dock = DockStyle.Bottom, DialogResult = DialogResult.OK, Enabled = false };
        this.Controls.Add(label);
        this.Controls.Add(okButton);

        this.KeyDown += (s, e) =>
        {
            currentModifiers = e.Modifiers;
            currentKey = e.KeyCode;
            HotkeyString = GetHotkeyString(currentModifiers, currentKey);
            label.Text = HotkeyString;
            okButton.Enabled = !string.IsNullOrWhiteSpace(HotkeyString);
            e.Handled = true;
        };
        okButton.Click += (s, e) => { this.DialogResult = DialogResult.OK; this.Close(); };
    }

    private string GetHotkeyString(Keys modifiers, Keys key)
    {
        if (key == Keys.ControlKey || key == Keys.ShiftKey || key == Keys.Menu) return "";
        var parts = new List<string>();
        if (modifiers.HasFlag(Keys.Control)) parts.Add("Ctrl");
        if (modifiers.HasFlag(Keys.Alt)) parts.Add("Alt");
        if (modifiers.HasFlag(Keys.Shift)) parts.Add("Shift");
        if (key != Keys.None && key != Keys.ControlKey && key != Keys.ShiftKey && key != Keys.Menu)
            parts.Add(key.ToString().ToUpper());
        return string.Join("+", parts);
    }
}