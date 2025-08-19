using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using System.Linq;
using System.Text;


namespace PdfAutomationApp
{
    public partial class Form1 : Form
    {
        private int resumeIndex = 0;
        private bool isResuming = false;

        public class ActionPoint
        {
            public string Type { get; set; } = "";
            public Point Location { get; set; }
            public int? DelayBefore { get; set; }
            public int? DelayAfter { get; set; }
            public string? RequiredHexColor { get; set; }
            public string Label { get; set; } = "";
            public bool IsFileDialog { get; set; } = false;
            public string? TextToType { get; set; }
            public string? Hotkey { get; set; }

            public override string ToString()
            {
                string beforeNote = DelayBefore.HasValue ? $" ⏱Before: {DelayBefore.Value}ms" : "";
                string afterNote = DelayAfter.HasValue ? $" ⏱After: {DelayAfter.Value}ms" : "";
                string colorNote = !string.IsNullOrEmpty(RequiredHexColor) ? $" | Color: #{RequiredHexColor}" : "";
                return $"{Label}{beforeNote}{afterNote}{colorNote}";
            }
        }

        public static class Keyboard
        {
            [DllImport("user32.dll")]
            private static extern short GetAsyncKeyState(Keys vKey);

            public static bool IsKeyDown(Keys key)
            {
                return (GetAsyncKeyState(key) & 0x8000) == 0x8000;
            }
        }

        private void RefreshListDisplay()
        {
            pointListBox.Items.Clear();
            foreach (var point in orderedPoints)
            {
                pointListBox.Items.Add(point.ToString());
            }
        }

        private List<ActionPoint> orderedPoints = new();
        private List<string> uploadedPdfPaths = new();
        private string uploadedPdfPath = "";

        private string firstPart = "";
        private string secondPart = "";
        private ComboBox startOptionsDropdown;
        private Keys assignedKey = Keys.None;
        private Button startButton = new();
        private Button continueButton = new();

        private Button uploadButton = new();
        private Button tapperButton = new();
        private Button scrollButton = new();
        private Button textInputButton = new();
        private int dragIndex = -1;
        private bool isDragging = false;
        private RichTextBox uploadedFileTextBox = new();
        private ListBox pointListBox = new();
        private ListBox logListBox = new();
        private ComboBox inputOptionsDropdown = new();
        private Button saveProfileButton = new();
        private ComboBox savedProfilesDropdown = new();
        private readonly string profilesFolder = Path.Combine(
            Directory.GetParent(Application.StartupPath)?.Parent?.Parent?.FullName
            ?? Application.StartupPath,
            "Profiles"
        );

        private CancellationTokenSource? cts;

        private ContextMenuStrip pointListContextMenu = new();
        private ContextMenuStrip uploadedPdfContextMenu = new();

        private int selectedPointIndex = -1;

        [DllImport("user32.dll")] static extern void SetCursorPos(int X, int Y);
        [DllImport("user32.dll")] static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);

        const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        const uint MOUSEEVENTF_LEFTUP = 0x0004;
        const uint MOUSEEVENTF_WHEEL = 0x0800;
        const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        const uint MOUSEEVENTF_RIGHTUP = 0x0010;

        public Form1()
        {
            InitializeComponent();
            startOptionsDropdown = new ComboBox();
            InitUI();
        }

        #region UI Initialization and Setup

        private void InitUI()
        {
            this.Text = "AutoFlow PDF";
            this.Size = new Size(700, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            try { this.Icon = new Icon("favi.ico"); } catch { }
            this.BackColor = ColorTranslator.FromHtml("#fafafa");

            ToolTip tooltip = new()
            {
                OwnerDraw = true,
                BackColor = Color.LightYellow,
                ForeColor = Color.Black,
                InitialDelay = 500,
                ReshowDelay = 100,
                AutoPopDelay = 10000,
                ShowAlways = true
            };

            tooltip.Draw += (s, e) =>
            {
                e.Graphics.FillRectangle(new SolidBrush(Color.LightYellow), e.Bounds);
                e.Graphics.DrawRectangle(Pens.Gray, e.Bounds);
                TextRenderer.DrawText(e.Graphics, e.ToolTipText, SystemFonts.DefaultFont,
                                    e.Bounds, Color.Black, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
            };

            int btnSize = 40;
            int y = this.ClientSize.Height - btnSize - 10;

            startButton = new Button
            {
                Text = "Start",
                Size = new Size(100, 40),
                Location = new Point(170, 400),
                Enabled = false
            };
            tooltip.SetToolTip(startButton, "Start");
            startButton.Click += StartButton_Click;
            this.Controls.Add(startButton);

            startOptionsDropdown = new ComboBox
            {
                Location = new Point(startButton.Right + 15, startButton.Top + 5),
                Size = new Size(170, 50),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            startOptionsDropdown.Items.AddRange(new string[]
            {
                "Chart Documents",
                "MedPA,Benefits,Denial,Approval",
                "Authorization Forms",
                "Documents",
                "ELIG",
                "2021documents",
                "cha",
                "Insurance Verification"
            });
            startOptionsDropdown.SelectedIndex = 0;
            this.Controls.Add(startOptionsDropdown);

            // Context menu for editing items
            ContextMenuStrip dropdownContextMenu = new();
            dropdownContextMenu.Items.Add("Edit", null, (s, e) =>
            {
                if (startOptionsDropdown.SelectedIndex >= 0)
                {
                    var selectedItem = startOptionsDropdown.Items[startOptionsDropdown.SelectedIndex];
                    string currentValue = selectedItem?.ToString() ?? string.Empty;

                    // Pass a non-null value to InputBox
                    string input = Microsoft.VisualBasic.Interaction.InputBox("Edit Item:", "Edit Option", currentValue ?? "");

                    if (!string.IsNullOrWhiteSpace(input))
                    {
                        startOptionsDropdown.Items[startOptionsDropdown.SelectedIndex] = input;
                    }
                }
            });

            // Show menu on right-click
            startOptionsDropdown.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Right && startOptionsDropdown.SelectedIndex >= 0)
                {
                    dropdownContextMenu.Show(startOptionsDropdown, e.Location);
                }
            };


            uploadButton = new Button
            {
                Text = "Upload PDF",
                Size = new Size(160, 40),
                Location = new Point(230, 50)
            };
            uploadButton.Click += UploadButton_Click;
            this.Controls.Add(uploadButton);

            Button refreshButton = new Button
            {
                Text = "🔄",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                Size = new Size(40, 40),
                Location = new Point(uploadButton.Right + 10, uploadButton.Top),
                FlatStyle = FlatStyle.Flat
            };
            refreshButton.FlatAppearance.BorderSize = 0;

            refreshButton.Click += (s, e) =>
            {
                orderedPoints.Clear();
                uploadedPdfPaths.Clear();
                pointListBox.Items.Clear();
                uploadedFileTextBox.Clear();
                logListBox.Items.Clear();
                inputOptionsDropdown.Text = "";
                pointListBox.SelectedIndex = -1;
                startButton.Enabled = false;
                continueButton.Enabled = false;
                resumeIndex = 0;
                isResuming = false;
                this.Text = "PDF Automation Tool";
            };

            this.Controls.Add(refreshButton);

            uploadedFileTextBox = new RichTextBox
            {
                ReadOnly = true,
                Multiline = true,
                Location = new Point(165, 100),
                Size = new Size(300, 280)
            };

            this.Controls.Add(uploadedFileTextBox);

            uploadedPdfContextMenu.Items.Add("🗑 Delete PDF", null, (s, e) =>
            {
                if (uploadedPdfContextMenu.Tag is not string clickedFileName)
                    return;

                string? fullPath = uploadedPdfPaths.FirstOrDefault(p => Path.GetFileName(p).Equals(clickedFileName, StringComparison.OrdinalIgnoreCase));

                if (!string.IsNullOrEmpty(fullPath))
                {
                    var confirm = MessageBox.Show($"Delete '{clickedFileName}'?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (confirm == DialogResult.Yes)
                    {
                        uploadedPdfPaths.Remove(fullPath);
                        UpdateUploadedFileListDisplay();
                    }
                }
            });

            uploadedFileTextBox.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    uploadedFileTextBox.Focus();

                    int caretIndex = uploadedFileTextBox.GetCharIndexFromPosition(e.Location);
                    int lineIndex = uploadedFileTextBox.GetLineFromCharIndex(caretIndex);

                    if (lineIndex >= 0 && lineIndex < uploadedFileTextBox.Lines.Length)
                    {
                        string clickedLine = uploadedFileTextBox.Lines[lineIndex];
                        int start = uploadedFileTextBox.GetFirstCharIndexFromLine(lineIndex);
                        uploadedFileTextBox.Select(start, clickedLine.Length);

                        uploadedPdfContextMenu.Tag = clickedLine.Trim();
                        uploadedPdfContextMenu.Show(uploadedFileTextBox, e.Location);
                    }
                }
            };

            int leftX = 10, topY = 40, buttonWidth = 130, buttonHeight = 40, gapY = 5;

            tapperButton = new Button { Text = "Add Tap Point", Size = new Size(buttonWidth, buttonHeight), Location = new Point(leftX, topY) }; tooltip.SetToolTip(tapperButton, "Add a click/tap location to the automation.");
            scrollButton = new Button { Text = "Add Scroll Point", Size = new Size(buttonWidth, buttonHeight), Location = new Point(leftX, topY + buttonHeight + gapY) }; tooltip.SetToolTip(scrollButton, "Add a scroll action (like mouse wheel down).");
            textInputButton = new Button { Text = "Add Text Input", Size = new Size(buttonWidth, buttonHeight), Location = new Point(leftX, topY + 2 * (buttonHeight + gapY)) }; tooltip.SetToolTip(textInputButton, "Add this button where you want to type text.");
            Button fileDialogTapButton = new Button
            {
                Text = "FileDialog",
                Size = new Size(buttonWidth, buttonHeight),
                Location = new Point(leftX, topY + 3 * (buttonHeight + gapY))
            };
            tooltip.SetToolTip(fileDialogTapButton, "Tap a button that opens a FileDialog. Auto-fill path.");

            tapperButton.Click += (s, e) => CapturePoint("Tap");
            scrollButton.Click += (s, e) => CapturePoint("Scroll");
            textInputButton.Click += (s, e) => CapturePoint("TextInput");
            fileDialogTapButton.Click += (s, e) => CaptureFileDialogPoint();

            this.Controls.AddRange(new Control[] { tapperButton, scrollButton, textInputButton, fileDialogTapButton });

            Button decideButton = new Button
            {
                Text = "Ctrl Button",
                Size = new Size(buttonWidth, buttonHeight),
                Location = new Point(leftX, topY + 4 * (buttonHeight + gapY))
            };
            decideButton.Click += (s, e) => CaptureHotkeyPoint();
            this.Controls.Add(decideButton);
            tooltip.SetToolTip(decideButton, "Tap only if the color at point matches. Waits max 10s.");

            Button rightClickButton = new Button
            {
                Text = "R. Click",
                Size = new Size(buttonWidth, buttonHeight),
                Location = new Point(leftX, topY + 5 * (buttonHeight + gapY))
            };

            rightClickButton.Click += (s, e) => CaptureRightClickPoint();
            this.Controls.Add(rightClickButton);
            tooltip.SetToolTip(rightClickButton, "Capture point to simulate right-click.");

            inputOptionsDropdown = new ComboBox
            {
                Location = new Point(leftX, topY + 6 * (buttonHeight + gapY)),
                Size = new Size(buttonWidth, buttonHeight),
                DropDownStyle = ComboBoxStyle.DropDown,
                Font = new Font("Segoe UI", 10, FontStyle.Regular)
            };

            inputOptionsDropdown.Items.AddRange(new string[] { "Valdez,Norma", "Lomas,Sandra" });
            this.Controls.Add(inputOptionsDropdown);

            logListBox = new ListBox
            {
                Location = new Point(490, 40),
                Size = new Size(180, 150),
                HorizontalScrollbar = true,
                Font = new Font("Segoe UI", 9, FontStyle.Italic)
            };
            this.Controls.Add(logListBox);

            pointListBox = new ListBox
            {
                Location = new Point(490, logListBox.Bottom + 10),
                Size = new Size(180, 180),
                HorizontalScrollbar = true
            };
            this.Controls.Add(pointListBox);

            continueButton = new Button
            {
                Text = "Continue",
                Size = new Size(100, 40),
                Location = new Point(leftX, topY + 7 * (buttonHeight + gapY)),
                Enabled = false
            };
            tooltip.SetToolTip(continueButton, "Continue automation from selected point.");
            continueButton.Click += ContinueButton_Click;
            this.Controls.Add(continueButton);

            Button browserButton = new Button
            {
                Text = "Open Browser",
                Size = new Size(120, 40),
                Location = new Point(continueButton.Left, continueButton.Bottom + 10)
            };
            browserButton.Click += (s, e) =>
            {
                if (Application.OpenForms.OfType<BrowserForm>().Any())
                {
                    var existing = Application.OpenForms.OfType<BrowserForm>().First();
                    existing.Focus();
                }
                else
                {
                    BrowserForm browserForm = new BrowserForm();
                    browserForm.Show();
                }
            };
            this.Controls.Add(browserButton);

            pointListContextMenu.Items.Add("🧹 Delete", null, DeletePoint_Click);
            pointListContextMenu.Items.Add("📝 Update", null, UpdatePoint_Click);
            pointListContextMenu.Items.Add("✏ Rename", null, RenamePoint_Click);
            pointListContextMenu.Items.Add("⏱ Set Timer", null, SetTimerForPoint_Click);
            pointListContextMenu.Items.Add("🎯 Set Required Color", null, SetColorRequirement_Click);

            pointListBox.AllowDrop = true;
            pointListBox.MouseDown += PointListBox_MouseDown;
            pointListBox.MouseMove += PointListBox_MouseMove;
            pointListBox.DragOver += PointListBox_DragOver;
            pointListBox.DragDrop += PointListBox_DragDrop;

            pointListBox.MouseUp += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    int index = pointListBox.IndexFromPoint(e.Location);
                    if (index != ListBox.NoMatches)
                    {
                        selectedPointIndex = index;
                        pointListBox.SelectedIndex = index;
                        pointListContextMenu.Show(pointListBox, e.Location);
                    }
                }
            };

            saveProfileButton = new Button
            {
                Text = "Save",
                Size = new Size(80, 30),
                Location = new Point(
                    pointListBox.Left + (pointListBox.Width - 170) / 2,
                    pointListBox.Bottom + -1
                )
            };
            saveProfileButton.Click += SaveProfileButton_Click;
            this.Controls.Add(saveProfileButton);

            savedProfilesDropdown = new ComboBox
            {
                Location = new Point(pointListBox.Left, saveProfileButton.Bottom + 5),
                Size = new Size(pointListBox.Width, 30),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            savedProfilesDropdown.SelectedIndexChanged += SavedProfilesDropdown_SelectedIndexChanged;
            this.Controls.Add(savedProfilesDropdown);

            Directory.CreateDirectory(profilesFolder);
            LoadSavedProfilesToDropdown();

            ContextMenuStrip profileContextMenu = new();
            profileContextMenu.Items.Add("🗑 Delete (From Disk too)", null, (s, e) =>
            {
                string selectedProfile = savedProfilesDropdown.SelectedItem?.ToString() ?? "";
                if (string.IsNullOrWhiteSpace(selectedProfile)) return;

                string filePath = Path.Combine(profilesFolder, selectedProfile + ".json");

                if (File.Exists(filePath))
                {
                    var confirm = MessageBox.Show($"Are you sure to delete '{selectedProfile}' permanently?", "Confirm Delete", MessageBoxButtons.YesNo);
                    if (confirm == DialogResult.Yes)
                    {
                        try
                        {
                            File.Delete(filePath);
                            LoadSavedProfilesToDropdown();
                            MessageBox.Show($"Profile '{selectedProfile}' deleted.");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Could not delete file: " + ex.Message);
                        }
                    }
                }
            });

            savedProfilesDropdown.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Right && savedProfilesDropdown.SelectedIndex != -1)
                {
                    profileContextMenu.Show(savedProfilesDropdown, e.Location);
                }
            };

            Button clearButton = new Button
            {
                Text = "Clear",
                Size = new Size(80, 30),
                Location = new Point(saveProfileButton.Right + 16, saveProfileButton.Top)
            };
            clearButton.Click += (s, e) =>
            {
                orderedPoints.Clear();
                pointListBox.Items.Clear();
                startButton.Enabled = false;
            };
            this.Controls.Add(clearButton);
        }

        #endregion

        #region Point Capture and Editing

        private void CaptureRightClickPoint()
        {
            this.Hide();
            MessageBox.Show("Click where you want to simulate a right-click...");

            var timer = new System.Windows.Forms.Timer { Interval = 100 };
            timer.Tick += (s, e) =>
            {
                if (Control.MouseButtons == MouseButtons.Left)
                {
                    var point = Cursor.Position;
                    int count = orderedPoints.Count(p => p.Type == "RightClick") + 1;
                    string label = $"{count} RightClick";

                    var ap = new ActionPoint
                    {
                        Type = "RightClick",
                        Location = point,
                        Label = label
                    };

                    orderedPoints.Add(ap);
                    RefreshListDisplay();
                    timer.Stop();
                    timer.Dispose();
                    startButton.Enabled = true;
                    this.Show();
                    this.Activate();
                }
            };

            timer.Start();
        }

        private void CaptureHotkeyPoint()
        {
            this.Hide();
            using (var hotkeyForm = new HotkeyCaptureForm())
            {
                if (hotkeyForm.ShowDialog() == DialogResult.OK)
                {
                    string hotkey = hotkeyForm.HotkeyString;
                    if (!string.IsNullOrWhiteSpace(hotkey))
                    {
                        var ap = new ActionPoint
                        {
                            Type = "Hotkey",
                            Label = $"Hotkey: {hotkey}",
                            Hotkey = hotkey
                        };
                        orderedPoints.Add(ap);
                        RefreshListDisplay();
                        startButton.Enabled = true;
                    }
                }
            }
            this.Show();
            this.Activate();
        }

        private void CapturePoint(string type)
        {
            this.Hide();
            MessageBox.Show("Click on the screen to capture point.");

            var timer = new System.Windows.Forms.Timer { Interval = 100 };
            timer.Tick += (s, e) =>
            {
                if (Control.MouseButtons == MouseButtons.Left)
                {
                    var point = Cursor.Position;
                    int count = orderedPoints.FindAll(p => p.Type == type).Count + 1;
                    string label = $"{count} {type}";

                    var ap = new ActionPoint
                    {
                        Type = type,
                        Location = point,
                        Label = label
                    };

                    orderedPoints.Add(ap);
                    RefreshListDisplay();
                    timer.Stop();
                    timer.Dispose();
                    startButton.Enabled = true;
                    this.Show();
                    this.Activate();
                }
            };

            timer.Start();
        }

        private void CaptureFileDialogPoint()
        {
            this.Hide();
            MessageBox.Show("Click to mark the File Dialog open position.");
            var timer = new System.Windows.Forms.Timer { Interval = 100 };
            timer.Tick += (s, e) =>
            {
                if (Control.MouseButtons == MouseButtons.Left)
                {
                    var point = Cursor.Position;
                    int count = orderedPoints.Count(p => p.IsFileDialog) + 1;

                    var ap = new ActionPoint
                    {
                        Type = "Tap",
                        Location = point,
                        Label = $"{count} Tap FileDialog",
                        IsFileDialog = true
                    };

                    orderedPoints.Add(ap);
                    RefreshListDisplay();
                    timer.Stop();
                    timer.Dispose();
                    startButton.Enabled = true;
                    this.Show();
                    this.Activate();
                }
            };
            timer.Start();
        }

        #endregion

        #region ListBox Drag and Drop

        private void PointListBox_MouseDown(object? sender, MouseEventArgs e)
        {
            dragIndex = pointListBox.IndexFromPoint(e.Location);
            isDragging = dragIndex != ListBox.NoMatches;
        }

        private void PointListBox_MouseMove(object? sender, MouseEventArgs e)
        {
            if (isDragging && dragIndex != -1 && e.Button == MouseButtons.Left)
            {
                pointListBox.DoDragDrop(pointListBox.Items[dragIndex], DragDropEffects.Move);
                isDragging = false;
            }
        }

        private void PointListBox_DragOver(object? sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void PointListBox_DragDrop(object? sender, DragEventArgs e)
        {
            Point point = pointListBox.PointToClient(new Point(e.X, e.Y));
            int targetIndex = pointListBox.IndexFromPoint(point);

            if (targetIndex != ListBox.NoMatches && dragIndex != targetIndex)
            {
                var item = orderedPoints[dragIndex];
                orderedPoints.RemoveAt(dragIndex);
                orderedPoints.Insert(targetIndex, item);

                RefreshListDisplay();
                pointListBox.SelectedIndex = targetIndex;
                selectedPointIndex = targetIndex;
            }

            dragIndex = -1;
        }

        #endregion

        #region Profile Save/Load

        private void SaveProfileButton_Click(object? sender, EventArgs e)
        {
            string profileName = Interaction.InputBox("Enter a name to save this profile:", "Save Profile", "MyProfile");
            if (string.IsNullOrWhiteSpace(profileName)) return;

            string filePath = Path.Combine(profilesFolder, profileName + ".json");

            try
            {
                string json = System.Text.Json.JsonSerializer.Serialize(orderedPoints, new System.Text.Json.JsonSerializerOptions
                {
                    WriteIndented = true
                });

                File.WriteAllText(filePath, json);
                MessageBox.Show($"Profile '{profileName}' saved successfully.");

                LoadSavedProfilesToDropdown();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving profile: " + ex.Message);
            }
        }

        private void LoadSavedProfilesToDropdown()
        {
            savedProfilesDropdown.Items.Clear();
            foreach (var file in Directory.GetFiles(profilesFolder, "*.json"))
            {
                savedProfilesDropdown.Items.Add(Path.GetFileNameWithoutExtension(file));
            }
        }

        private void SavedProfilesDropdown_SelectedIndexChanged(object? sender, EventArgs e)
        {
            string selectedProfile = savedProfilesDropdown.SelectedItem?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(selectedProfile)) return;

            string filePath = Path.Combine(profilesFolder, selectedProfile + ".json");

            try
            {
                string json = File.ReadAllText(filePath);
                var loadedPoints = System.Text.Json.JsonSerializer.Deserialize<List<ActionPoint>>(json);

                if (loadedPoints != null)
                {
                    orderedPoints.Clear();
                    orderedPoints.AddRange(loadedPoints);
                    RefreshListDisplay();
                    this.Text = $"PDF Automation Tool – [{selectedProfile}]";
                    startButton.Enabled = orderedPoints.Count > 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading profile: " + ex.Message);
            }
        }

        #endregion

        #region Point Context Menu Actions

        private void DeletePoint_Click(object? sender, EventArgs e)
        {
            if (selectedPointIndex >= 0 && selectedPointIndex < orderedPoints.Count)
            {
                orderedPoints.RemoveAt(selectedPointIndex);
                RefreshListDisplay();
                selectedPointIndex = -1;
                startButton.Enabled = orderedPoints.Count > 0;
            }
        }

        private void UpdatePoint_Click(object? sender, EventArgs e)
        {
            if (selectedPointIndex < 0 || selectedPointIndex >= orderedPoints.Count) return;

            this.Hide(); MessageBox.Show("Click to choose new location.");
            var timer = new System.Windows.Forms.Timer { Interval = 100 };
            timer.Tick += (s, e) =>
            {
                if (Control.MouseButtons == MouseButtons.Left)
                {
                    var pt = Cursor.Position;
                    orderedPoints[selectedPointIndex].Location = pt;
                    RefreshListDisplay();
                    timer.Stop(); timer.Dispose(); this.Show(); this.Activate();
                }
            };
            timer.Start();
        }

        private void RenamePoint_Click(object? sender, EventArgs e)
        {
            if (selectedPointIndex < 0 || selectedPointIndex >= pointListBox.Items.Count)
                return;

            var ap = orderedPoints[selectedPointIndex];
            string currentLabel = ap.Label ?? "Label";
            string input = Microsoft.VisualBasic.Interaction.InputBox("Enter new label for this point:", "Rename Point", currentLabel);

            if (!string.IsNullOrWhiteSpace(input))
            {
                ap.Label = input;
                RefreshListDisplay();
            }
        }

        private void SetTimerForPoint_Click(object? sender, EventArgs e)
        {
            if (selectedPointIndex < 0 || selectedPointIndex >= orderedPoints.Count)
                return;

            var ap = orderedPoints[selectedPointIndex];

            var choice = MessageBox.Show("Do you want to set a delay **before** the action?", "Set Timer",
                                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            if (choice == DialogResult.Cancel)
                return;

            bool isBefore = choice == DialogResult.Yes;
            string input = Interaction.InputBox(
                $"Enter delay in milliseconds to apply {(isBefore ? "before" : "after")} the action:",
                "Set Delay", "1000");

            if (string.IsNullOrWhiteSpace(input))
                return;

            if (int.TryParse(input, out int delay) && delay >= 0)
            {
                if (isBefore)
                    ap.DelayBefore = delay;
                else
                    ap.DelayAfter = delay;

                RefreshListDisplay();
            }
            else
            {
                MessageBox.Show("Invalid number entered.");
            }
        }

        private void SetColorRequirement_Click(object? sender, EventArgs e)
        {
            if (selectedPointIndex < 0 || selectedPointIndex >= orderedPoints.Count)
                return;

            var ap = orderedPoints[selectedPointIndex];

            if (ap.Type != "Tap")
            {
                MessageBox.Show("Color check is only supported for Tap points.");
                return;
            }

            this.Hide();
            MessageBox.Show("Click on the pixel to pick required color...");

            var timer = new System.Windows.Forms.Timer { Interval = 100 };
            timer.Tick += (s, e) =>
            {
                if (Control.MouseButtons == MouseButtons.Left)
                {
                    var pt = Cursor.Position;
                    Color color;
                    try
                    {
                        color = GetColorAt(pt);
                    }
                    catch
                    {
                        color = Color.Black;
                    }

                    ap.RequiredHexColor = $"{color.R:X2}{color.G:X2}{color.B:X2}";
                    RefreshListDisplay();

                    timer.Stop();
                    timer.Dispose();
                    this.Show();
                    this.Activate();
                }
            };
            timer.Start();
        }

        #endregion

        #region PDF Upload and Parsing

        private void UploadButton_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                Multiselect = true
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                uploadedPdfPaths.Clear();
                uploadedPdfPaths.AddRange(ofd.FileNames);
                uploadedFileTextBox.Text = string.Join(Environment.NewLine, uploadedPdfPaths.Select(Path.GetFileName));

                if (uploadedPdfPaths.Count > 0)
                {
                    uploadedPdfPath = uploadedPdfPaths[0];
                    ParsePdfFilename(uploadedPdfPath);

                    MessageBox.Show($"First file parsed:\nAccount Number: {firstPart}\nFolder Name: {secondPart}\nFull Name: {this.Tag}");
                }
            }
        }

        private void ParsePdfFilename(string pdfPath)
        {
            var name = Path.GetFileNameWithoutExtension(pdfPath)?.Trim() ?? "";

            name = Regex.Replace(name, @"\s*,\s*", ",");

            int lastComma = name.LastIndexOf(',');
            if (lastComma >= 0 && lastComma < name.Length - 1)
            {
                firstPart = name.Substring(lastComma + 1).Replace(" ", "");
            }
            else
            {
                firstPart = name; // fallback to full name if no comma
            }

            secondPart = ""; // You can update this as needed
            this.Tag = name;
        }

        #endregion

        #region Automation Logic

        private async void StartButton_Click(object? sender, EventArgs e)
        {
            await RunAutomation(isResuming ? resumeIndex : 0);
        }

        private async void ContinueButton_Click(object? sender, EventArgs e)
        {
            if (uploadedPdfPaths.Count == 0)
            {
                MessageBox.Show("No uploaded PDF files remaining to process.", "Missing Files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pointListBox.SelectedIndex < 0 || pointListBox.SelectedIndex >= orderedPoints.Count)
            {
                MessageBox.Show("Please select a point to resume from.", "Resume Required", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            resumeIndex = pointListBox.SelectedIndex;
            isResuming = true;
            continueButton.Enabled = false;

            await RunAutomation(resumeIndex);
        }

        private async Task RunAutomation(int startIndex)
        {
            if (uploadedPdfPaths.Count == 0)
            {
                MessageBox.Show("No PDF files to process.");
                return;
            }

            using var popup = new DelayPopup();
            if (popup.ShowDialog(this) != DialogResult.OK)
                return;

            int defaultDelay = popup.DelayValue;
            cts = new CancellationTokenSource();
            CancellationToken token = cts.Token;

            _ = Task.Run(() =>
            {
                while (!token.IsCancellationRequested)
                {
                    if (Keyboard.IsKeyDown(Keys.ControlKey) &&
                        Keyboard.IsKeyDown(Keys.ShiftKey) &&
                        Keyboard.IsKeyDown(Keys.Z))
                    {
                        cts.Cancel();
                        break;
                    }
                    Thread.Sleep(100);
                }
            });

            try
            {
                while (uploadedPdfPaths.Count > 0)
                {
                    var pdfPath = uploadedPdfPaths[0];
                    uploadedPdfPath = pdfPath;
                    ParsePdfFilename(uploadedPdfPath);
                    Invoke(() => HighlightCurrentFileInTextBox(pdfPath));

                    for (int i = startIndex; i < orderedPoints.Count; i++)
                    {
                        if (token.IsCancellationRequested)
                        {
                            resumeIndex = i;
                            Invoke(() => continueButton.Enabled = true);
                            isResuming = true;
                            return;
                        }

                        var ap = orderedPoints[i];
                        Invoke(() => pointListBox.SelectedIndex = i);
                        SetCursorPos(ap.Location.X, ap.Location.Y);

                        int delayBefore = ap.DelayBefore ?? 0;
                        if (delayBefore > 0)
                            await Task.Delay(delayBefore, token);

                        if (ap.Type == "Tap" && !string.IsNullOrEmpty(ap.RequiredHexColor))
                        {
                            bool matched = false;
                            for (int r = 0; r < 20 && !token.IsCancellationRequested; r++)
                            {
                                var color = GetColorAt(ap.Location);
                                string currentHex = $"{color.R:X2}{color.G:X2}{color.B:X2}".ToUpperInvariant();
                                if (string.Equals(currentHex, ap.RequiredHexColor?.ToUpperInvariant(), StringComparison.Ordinal))
                                {
                                    matched = true;
                                    break;
                                }
                                await Task.Delay(500, token);
                            }
                            if (!matched)
                            {
                                Invoke(() => logListBox.Items.Add($"[SKIPPED] Tap '{ap.Label}' – color never matched."));
                                continue;
                            }
                        }

                        try
                        {
                            switch (ap.Type)
                            {
                               case "Tap":
                                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

                                    if (ap.IsFileDialog)
                                    {
                                        await Task.Delay(3500, token);

                                        if (!IsFileDialogStillActive())
                                        {
                                            Invoke(() =>
                                                MessageBox.Show("File Dialog did not open. Automation paused. " +
                                                                "Please open the dialog manually and click Continue.",
                                                "File Dialog Missing", MessageBoxButtons.OK, MessageBoxIcon.Warning));

                                            resumeIndex = i;       
                                            isResuming = true;      
                                            Invoke(() => continueButton.Enabled = true);

                                            return; 
                                        }
                                        string fileName = Path.GetFileName(uploadedPdfPath);
                                        string directory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                                        string fullPath = Path.Combine(directory, fileName);

                                        try
                                        {
                                            SendKeys.SendWait(fullPath);
                                            await Task.Delay(300, token);
                                            SendKeys.SendWait("{ENTER}");
                                            await Task.Delay(3000, token);

                                            if (IsFileDialogStillActive())
                                            {
                                                Invoke(() =>
                                                    MessageBox.Show("File Dialog did not close. Automation paused. " +
                                                                    "Close the dialog and click Continue to resume.",
                                                    "File Dialog Issue", MessageBoxButtons.OK, MessageBoxIcon.Warning));

                                                resumeIndex = i;
                                                isResuming = true;
                                                Invoke(() => continueButton.Enabled = true);
                                                return;
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Invoke(() => logListBox.Items.Add("SendKeys error: " + ex.Message));
                                        }

                                        Invoke(() =>
                                        {
                                            this.Show(); this.Activate(); this.BringToFront(); this.Focus();
                                        });
                                        SetForegroundWindow(this.Handle);
                                        await Task.Delay(500, token);
                                    }

                                    break;
                                case "Scroll":
                                    mouse_event(MOUSEEVENTF_WHEEL, 0, 0, unchecked((uint)-120), 0);
                                    break;

                                case "TextInput":
                                    int inputIndex = orderedPoints.Take(i + 1).Count(p => p.Type == "TextInput");
                                    if (inputIndex == 1)
                                    {
                                        SendKeys.SendWait(firstPart);
                                        await Task.Delay(2500, token);
                                        SendKeys.SendWait("{ENTER}");
                                    }
                                    else if (inputIndex == 2)
                                    {
                                        SendKeys.SendWait(string.IsNullOrWhiteSpace(startOptionsDropdown.Text) ? "[Empty Input]" : startOptionsDropdown.Text);
                                    }
                                    else
                                    {
                                        SendKeys.SendWait(string.IsNullOrWhiteSpace(inputOptionsDropdown.Text) ? "[Empty Input]" : inputOptionsDropdown.Text);
                                    }
                                    break;

                                case "RightClick":
                                    mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                                    break;
                                case "Hotkey":
                                    if (!string.IsNullOrWhiteSpace(ap.Hotkey))
                                    {
                                        SendHotkey(ap.Hotkey);
                                        await Task.Delay(300, token);
                                    }
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Invoke(() => logListBox.Items.Add("Action error: " + ex.Message));
                        }

                        string logMsg = $"{ap.Label} {(ap.Type == "Tap" ? "completed" : ap.Type == "Scroll" ? "scrolled" : "typed")}";
                        Invoke(() => logListBox.Items.Add(logMsg));

                        int delayAfter = ap.DelayAfter ?? defaultDelay;
                        await Task.Delay(delayAfter, token);
                    }

                    isResuming = false;
                    resumeIndex = 0;

                    Invoke(() =>
                    {
                        if (uploadedPdfPaths.Count > 0)
                        {
                            uploadedPdfPaths.RemoveAt(0);
                            UpdateUploadedFileListDisplay();
                            logListBox.Items.Add($"✅ Finished: {Path.GetFileName(pdfPath)} removed from queue.");
                        }
                    });

                    startIndex = 0;
                }
            }
            catch (OperationCanceledException)
            {
                resumeIndex = pointListBox.SelectedIndex;
                Invoke(() => continueButton.Enabled = true);
                isResuming = true;
            }
            catch (Exception ex)
            {
                Invoke(() => MessageBox.Show("Error: " + ex.Message));
            }
        }

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);

        private void SendHotkey(string hotkey)
        {
            var keys = hotkey.Split('+');
            var keyList = new List<Keys>();
            foreach (var k in keys)
            {
                switch (k.Trim().ToUpper())
                {
                    case "CTRL": keyList.Add(Keys.ControlKey); break;
                    case "ALT": keyList.Add(Keys.Menu); break;
                    case "SHIFT": keyList.Add(Keys.ShiftKey); break;
                    default:
                        if (Enum.TryParse(typeof(Keys), k, true, out var result))
                            keyList.Add((Keys)result);
                        break;
                }
            }

            // Press modifiers
            foreach (var k in keyList.Where(x => x == Keys.ControlKey || x == Keys.Menu || x == Keys.ShiftKey))
                keybd_event((byte)k, 0, 0, 0);

            // Press main key
            var mainKey = keyList.FirstOrDefault(x => x != Keys.ControlKey && x != Keys.Menu && x != Keys.ShiftKey);
            if (mainKey != Keys.None)
                keybd_event((byte)mainKey, 0, 0, 0);

            // Release main key
            if (mainKey != Keys.None)
                keybd_event((byte)mainKey, 0, 2, 0);

            // Release modifiers
            foreach (var k in keyList.Where(x => x == Keys.ControlKey || x == Keys.Menu || x == Keys.ShiftKey).Reverse())
                keybd_event((byte)k, 0, 2, 0);
        }

        #endregion

        #region Helper Methods

        private void HighlightCurrentFileInTextBox(string filePath)
        {
            string fileName = Path.GetFileName(filePath);
            int start = uploadedFileTextBox.Text.IndexOf(fileName, StringComparison.OrdinalIgnoreCase);
            if (start >= 0)
            {
                uploadedFileTextBox.SelectAll();
                uploadedFileTextBox.SelectionBackColor = uploadedFileTextBox.BackColor;
                uploadedFileTextBox.Select(start, fileName.Length);
                uploadedFileTextBox.SelectionBackColor = Color.LightBlue;
            }
            else
            {
                uploadedFileTextBox.SelectAll();
                uploadedFileTextBox.SelectionBackColor = uploadedFileTextBox.BackColor;
            }
        }

        private void UpdateUploadedFileListDisplay()
        {
            uploadedFileTextBox.Text = string.Join(Environment.NewLine, uploadedPdfPaths.Select(Path.GetFileName));
            if (uploadedPdfPaths.Count > 0)
                HighlightCurrentFileInTextBox(uploadedPdfPaths[0]);
        }

        private Color GetColorAt(Point location)
        {
            using var bmp = new Bitmap(1, 1);
            using var g = Graphics.FromImage(bmp);
            g.CopyFromScreen(location, Point.Empty, new Size(1, 1));
            return bmp.GetPixel(0, 0);
        }
        
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        private bool IsFileDialogStillActive()
        {
            IntPtr hWnd = GetForegroundWindow();

            if (hWnd == IntPtr.Zero)
                return false;

            StringBuilder sb = new StringBuilder(256);
            GetWindowText(hWnd, sb, sb.Capacity);
            string title = sb.ToString();

            // Typical file dialog titles, you can add more if needed
            string[] fileDialogTitles = { "Open", "Choose File", "Select File", "File Upload" };

            return fileDialogTitles.Any(t => title.Contains(t, StringComparison.OrdinalIgnoreCase));
        }
        #endregion
    }
}