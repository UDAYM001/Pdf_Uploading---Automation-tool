using System;
using System.Windows.Forms;
using System.Drawing;

namespace PdfAutomationApp
{
    public partial class DelayPopup : Form
    {
        public int DelayValue { get; private set; } = 1000;

        public DelayPopup()
        {
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "⏱ Set Timer";
            this.Size = new Size(130, 170); // Approx 3x4 cm
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.FromArgb(40, 40, 40);  // Light black background

            Label titleLabel = new Label
            {
                Text = "Set Delay (ms)",
                Size = new Size(100, 20),
                Location = new Point(10, 10),
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.White,
                BackColor = Color.Transparent
            };

            NumericUpDown delayInput = new NumericUpDown
            {
                Minimum = 100,
                Maximum = 10000,
                Increment = 100,
                Value = 1000,
                Size = new Size(110, 25),
                Location = new Point(10, 40),
                BackColor = Color.White,
                ForeColor = Color.Black
            };

            Button okButton = new Button
            {
                Text = "OK",
                Size = new Size(50, 30),
                Location = new Point(10, 80),
                BackColor = Color.DimGray,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            Button cancelButton = new Button
            {
                Text = "Cancel",
                Size = new Size(65, 30),
                Location = new Point(75, 80),
                BackColor = Color.DimGray,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            okButton.Click += (s, e) =>
            {
                DelayValue = (int)delayInput.Value;
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            cancelButton.Click += (s, e) =>
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            };

            this.Controls.Add(titleLabel);
            this.Controls.Add(delayInput);
            this.Controls.Add(okButton);
            this.Controls.Add(cancelButton);
        }
    }
}
