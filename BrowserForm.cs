using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;

namespace PdfAutomationApp
{
    public partial class BrowserForm : Form
    {
        private WebView2 browser = null!;
        private TextBox urlBox = null!;
        private Button goButton = null!;
        private Panel topPanel = null!;

        private const string SettingsFile = "browser_settings.txt";
        private const string UserDataFolderName = "WebView2UserData";
        private double _zoomFactor = 1.0;

        public BrowserForm()
        {
            this.Text = "Isolated Browser";
            this.Size = new Size(1920, 1200);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = false;

            // Restore window size, position, and zoom if settings file exists
            RestoreWindowSettings();

            InitializeUI();

            this.Load += BrowserForm_Load;
        }

        private void InitializeUI()
        {
            // Top panel for input
            topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 45,
                Padding = new Padding(10),
                BackColor = Color.LightGray
            };

            urlBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10),
                PlaceholderText = "Enter URL here..."
            };

            goButton = new Button
            {
                Text = "Go",
                Width = 60,
                Dock = DockStyle.Right
            };

            goButton.Click += (s, e) => NavigateToUrl();
            urlBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    NavigateToUrl();
                    e.SuppressKeyPress = true;
                }
            };

            topPanel.Controls.Add(urlBox);
            topPanel.Controls.Add(goButton);
            this.Controls.Add(topPanel);

            // WebView2 Browser
            browser = new WebView2
            {
                Dock = DockStyle.Fill
            };
            this.Controls.Add(browser);

            // Optional: Add zoom in/out buttons
            var zoomInButton = new Button
            {
                Text = "+",
                Width = 30,
                Dock = DockStyle.Right
            };
            var zoomOutButton = new Button
            {
                Text = "-",
                Width = 30,
                Dock = DockStyle.Right
            };
            zoomInButton.Click += (s, e) => ChangeZoom(0.1);
            zoomOutButton.Click += (s, e) => ChangeZoom(-0.1);
            topPanel.Controls.Add(zoomInButton);
            topPanel.Controls.Add(zoomOutButton);
        }

        private async void BrowserForm_Load(object? sender, EventArgs e)
        {
            // Set a persistent user data folder for WebView2
            string userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PdfAutomationApp", UserDataFolderName
            );
            Directory.CreateDirectory(userDataFolder); // Ensure the folder exists

            var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

            await browser.EnsureCoreWebView2Async(env);
            browser.CoreWebView2.NavigationCompleted += Browser_NavigationCompleted;

            // Restore zoom factor using JavaScript
            SetBrowserZoom(_zoomFactor);

            this.WindowState = FormWindowState.Maximized;
        }

        private void Browser_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            // Hide the top panel after the first page load
            topPanel.Visible = false;

            // Re-apply zoom after navigation
            SetBrowserZoom(_zoomFactor);
        }

        private void NavigateToUrl()
        {
            string url = urlBox.Text.Trim();

            if (string.IsNullOrWhiteSpace(url)) return;

            if (!url.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            {
                url = "https://" + url;
            }

            browser.CoreWebView2.Navigate(url);
        }

        private void ChangeZoom(double delta)
        {
            _zoomFactor = Math.Max(0.25, Math.Min(5.0, _zoomFactor + delta));
            SetBrowserZoom(_zoomFactor);
        }

        private async void SetBrowserZoom(double zoom)
        {
            // Use JavaScript to set zoom
            if (browser.CoreWebView2 != null)
            {
                string script = $"document.body.style.zoom = '{zoom * 100}%';";
                try
                {
                    await browser.ExecuteScriptAsync(script);
                }
                catch { /* ignore errors if navigation not complete */ }
            }
        }

        // Save window size, position, and zoom on close
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            try
            {
                var settings = $"{this.Width},{this.Height},{this.Left},{this.Top},{_zoomFactor}";
                File.WriteAllText(SettingsFile, settings);
            }
            catch { /* ignore errors */ }
        }

        // Restore window size, position, and zoom on open
        private void RestoreWindowSettings()
        {
            if (File.Exists(SettingsFile))
            {
                var parts = File.ReadAllText(SettingsFile).Split(',');
                if (parts.Length == 5 &&
                    int.TryParse(parts[0], out int w) &&
                    int.TryParse(parts[1], out int h) &&
                    int.TryParse(parts[2], out int x) &&
                    int.TryParse(parts[3], out int y) &&
                    double.TryParse(parts[4], out double zoom))
                {
                    this.Size = new Size(w, h);
                    this.StartPosition = FormStartPosition.Manual;
                    this.Location = new Point(x, y);
                    _zoomFactor = zoom;
                }
            }
        }

        // Disable minimize button
        protected override CreateParams CreateParams
        {
            get
            {
                const int WS_MINIMIZEBOX = 0x00020000;
                CreateParams cp = base.CreateParams;
                cp.Style &= ~WS_MINIMIZEBOX;
                return cp;
            }
        }
    }
}