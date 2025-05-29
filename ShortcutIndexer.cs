using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Security.Principal;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace ShortcutIndexer
{
    class ColoredCombo : ComboBox
    {
        private static bool IsWindowsDarkTheme()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    var value = key.GetValue("AppsUseLightTheme");
                    if (value is int)
                    {
                        return (int)value == 0; // 0 = dark theme, 1 = light theme
                    }
                }
            }
            catch
            {
                // Fall back to light theme if unable to detect
            }
            return false;
        }

        private Color borderColor = IsWindowsDarkTheme() ? Color.Gray : Color.LightGray;
        public Color BorderColor
        {
            get { return borderColor; }
            set
            {
                if (borderColor != value)
                {
                    borderColor = value;
                    Invalidate();
                }
            }
        }

        private Color buttonColor = IsWindowsDarkTheme() ? Color.FromArgb(45, 45, 45) : Color.Transparent;
        public Color ButtonColor
        {
            get { return buttonColor; }
            set
            {
                if (buttonColor != value)
                {
                    buttonColor = value;
                    Invalidate();
                }
            }
        }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_PAINT && DropDownStyle != ComboBoxStyle.Simple)
            {
                var clientRect = ClientRectangle;
                var dropDownButtonWidth = SystemInformation.HorizontalScrollBarArrowWidth;
                var outerBorder = new Rectangle(clientRect.Location,
                    new Size(clientRect.Width - 1, clientRect.Height - 1));
                var innerBorder = new Rectangle(outerBorder.X + 1, outerBorder.Y + 1,
                    outerBorder.Width - dropDownButtonWidth - 2, outerBorder.Height - 2);
                var innerInnerBorder = new Rectangle(innerBorder.X + 1, innerBorder.Y + 1,
                    innerBorder.Width - 2, innerBorder.Height - 2);
                var dropDownRect = new Rectangle(innerBorder.Right + 1, innerBorder.Y,
                    dropDownButtonWidth, innerBorder.Height + 1);
                if (RightToLeft == RightToLeft.Yes)
                {
                    innerBorder.X = clientRect.Width - innerBorder.Right;
                    innerInnerBorder.X = clientRect.Width - innerInnerBorder.Right;
                    dropDownRect.X = clientRect.Width - dropDownRect.Right;
                    dropDownRect.Width += 1;
                }
                var innerBorderColor = Enabled ? BackColor : SystemColors.Control;
                var outerBorderColor = Enabled ? BorderColor : SystemColors.ControlDark;
                var buttonColor = Enabled ? ButtonColor : SystemColors.Control;
                var middle = new Point(dropDownRect.Left + dropDownRect.Width / 2,
                    dropDownRect.Top + dropDownRect.Height / 2);
                var arrow = new Point[]
                {
                    new Point(middle.X - 3, middle.Y - 2),
                    new Point(middle.X + 4, middle.Y - 2),
                    new Point(middle.X, middle.Y + 2)
                };
                var ps = new PAINTSTRUCT();
                bool shoulEndPaint = false;
                IntPtr dc;
                if (m.WParam == IntPtr.Zero)
                {
                    dc = BeginPaint(Handle, ref ps);
                    m.WParam = dc;
                    shoulEndPaint = true;
                }
                else
                {
                    dc = m.WParam;
                }
                var rgn = CreateRectRgn(innerInnerBorder.Left, innerInnerBorder.Top,
                    innerInnerBorder.Right, innerInnerBorder.Bottom);
                SelectClipRgn(dc, rgn);
                DefWndProc(ref m);
                DeleteObject(rgn);
                rgn = CreateRectRgn(clientRect.Left, clientRect.Top,
                    clientRect.Right, clientRect.Bottom);
                SelectClipRgn(dc, rgn);
                using (var g = Graphics.FromHdc(dc))
                {
                    using (var b = new SolidBrush(buttonColor))
                    {
                        g.FillRectangle(b, dropDownRect);
                    }
                    using (var b = new SolidBrush(outerBorderColor))
                    {
                        g.FillPolygon(b, arrow);
                    }
                    using (var p = new Pen(innerBorderColor))
                    {
                        g.DrawRectangle(p, innerBorder);
                        g.DrawRectangle(p, innerInnerBorder);
                    }
                    using (var p = new Pen(outerBorderColor))
                    {
                        g.DrawRectangle(p, outerBorder);
                    }
                }
                if (shoulEndPaint)
                    EndPaint(Handle, ref ps);
                DeleteObject(rgn);
            }
            else
                base.WndProc(ref m);
        }

        private const int WM_PAINT = 0xF;
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int L, T, R, B;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct PAINTSTRUCT
        {
            public IntPtr hdc;
            public bool fErase;
            public int rcPaint_left;
            public int rcPaint_top;
            public int rcPaint_right;
            public int rcPaint_bottom;
            public bool fRestore;
            public bool fIncUpdate;
            public int reserved1;
            public int reserved2;
            public int reserved3;
            public int reserved4;
            public int reserved5;
            public int reserved6;
            public int reserved7;
            public int reserved8;
        }
        [DllImport("user32.dll")]
        private static extern IntPtr BeginPaint(IntPtr hWnd,
            [In, Out] ref PAINTSTRUCT lpPaint);

        [DllImport("user32.dll")]
        private static extern bool EndPaint(IntPtr hWnd, ref PAINTSTRUCT lpPaint);

        [DllImport("gdi32.dll")]
        public static extern int SelectClipRgn(IntPtr hDC, IntPtr hRgn);

        [DllImport("user32.dll")]
        public static extern int GetUpdateRgn(IntPtr hwnd, IntPtr hrgn, bool fErase);
        public enum RegionFlags
        {
            ERROR = 0,
            NULLREGION = 1,
            SIMPLEREGION = 2,
            COMPLEXREGION = 3,
        }
        [DllImport("gdi32.dll")]
        internal static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateRectRgn(int x1, int y1, int x2, int y2);
    }

    public partial class MainForm : Form
    {
        private TextBox txtShortcutName;
        private ColoredCombo cmbLocation;
        private TextBox txtTargetPath;
        private TextBox txtArguments;
        private ColoredCombo cmbStartIn;
        private ColoredCombo cmbRunAs;
        private Button btnCreate;
        private Button btnCancel;
        private Button btnBrowseTarget;
        private CheckBox chkRunAsAdmin;
        private Label lblIcon;

        // Theme colors
        private Color _backgroundColor;
        private Color _foregroundColor;
        private Color _controlBackgroundColor;
        private Color _buttonBackgroundColor;
        private Color _accentColor;
        private bool _isDarkTheme;

        // Store custom location path
        private string _customLocationPath = null;

        public string TargetFile { get; set; }

        // Windows API for dark title bar
        [DllImport("dwmapi.dll", PreserveSig = true)]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, uint attr, ref int attrValue, int attrSize);

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("User32.dll")]
        private static extern IntPtr GetWindowDC(IntPtr hWnd);

        private const uint DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;
        private const uint DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
        private const int WM_NCPAINT = 0x85;

        public MainForm(string targetFile = null)
        {
            TargetFile = targetFile;
            DetectAndApplyTheme();
            InitializeComponent();
            LoadDefaultValues();
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == WM_NCPAINT && _isDarkTheme)
            {
                IntPtr hdc = GetWindowDC(m.HWnd);
                if (hdc != IntPtr.Zero)
                {
                    Graphics g = Graphics.FromHdc(hdc);

                    // Paint the title bar with dark color
                    using (var brush = new SolidBrush(Color.FromArgb(45, 45, 45)))
                    {
                        // Title bar height is typically around 30-32 pixels
                        g.FillRectangle(brush, new Rectangle(0, 0, this.Width, 32));
                    }

                    // Paint the title text
                    using (var textBrush = new SolidBrush(Color.White))
                    using (var font = new Font("Segoe UI", 9F))
                    {
                        var titleRect = new Rectangle(8, 6, this.Width - 150, 20);
                        g.DrawString(this.Text, font, textBrush, titleRect);
                    }

                    g.Flush();
                    ReleaseDC(m.HWnd, hdc);
                    g.Dispose();
                }
            }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            if (_isDarkTheme)
            {
                ApplyDarkTitleBar();
            }
        }

        private void ApplyDarkTitleBar()
        {
            if (Environment.OSVersion.Version.Major >= 10 && this.Handle != IntPtr.Zero)
            {
                var attribute = DWMWA_USE_IMMERSIVE_DARK_MODE;

                // Use older attribute for Windows 10 versions before 20H1
                if (Environment.OSVersion.Version.Build < 18985)
                {
                    attribute = DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1;
                }

                int useImmersiveDarkMode = 1;
                DwmSetWindowAttribute(this.Handle, attribute, ref useImmersiveDarkMode, sizeof(int));
            }
        }

        private bool IsRunningAsAdministrator()
        {
            try
            {
                using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
                {
                    WindowsPrincipal principal = new WindowsPrincipal(identity);
                    return principal.IsInRole(WindowsBuiltInRole.Administrator);
                }
            }
            catch
            {
                return false;
            }
        }

        private bool RequiresAdminPrivileges(int locationIndex)
        {
            // All Users locations require admin privileges
            if (locationIndex == 1 || locationIndex == 3 || locationIndex == 5)
                return true;

            // Custom location - check if the selected path requires admin privileges
            if (locationIndex == 6 && !string.IsNullOrEmpty(_customLocationPath))
            {
                return RequiresAdminForPath(_customLocationPath);
            }

            return false;
        }

        private bool RequiresAdminForPath(string path)
        {
            try
            {
                // Try to create a temporary file in the directory to test write permissions
                string testFile = Path.Combine(path, Path.GetRandomFileName());
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);
                return false; // No admin required
            }
            catch (UnauthorizedAccessException)
            {
                return true; // Admin required
            }
            catch (DirectoryNotFoundException)
            {
                // Directory doesn't exist, try to create it
                try
                {
                    Directory.CreateDirectory(path);
                    return false;
                }
                catch (UnauthorizedAccessException)
                {
                    return true;
                }
            }
            catch
            {
                return false; // Other errors, assume no admin required
            }
        }

        private void DetectAndApplyTheme()
        {
            _isDarkTheme = IsWindowsDarkTheme();

            if (_isDarkTheme)
            {
                // Dark theme colors
                _backgroundColor = Color.FromArgb(32, 32, 32);
                _foregroundColor = Color.White;
                _controlBackgroundColor = Color.FromArgb(45, 45, 45);
                _buttonBackgroundColor = Color.FromArgb(55, 55, 55);
                _accentColor = Color.FromArgb(0, 120, 215);
            }
            else
            {
                // Light theme colors
                _backgroundColor = Color.White;
                _foregroundColor = Color.Black;
                _controlBackgroundColor = Color.White;
                _buttonBackgroundColor = Color.FromArgb(240, 240, 240);
                _accentColor = Color.FromArgb(0, 120, 215);
            }
        }

        private bool IsWindowsDarkTheme()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    var value = key.GetValue("AppsUseLightTheme");
                    if (value is int)
                    {
                        return (int)value == 0; // 0 = dark theme, 1 = light theme
                    }
                }
            }
            catch
            {
                // Fall back to light theme if unable to detect
            }
            return false;
        }

        private void InitializeComponent()
        {
            this.Text = "Shortcut Indexer - Create Shortcut";
            this.Size = new Size(500, 360);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = _backgroundColor;
            this.ForeColor = _foregroundColor;

            // Modern Windows 11 styling
            this.Font = new Font("Segoe UI", 9F);

            // Shortcut Name
            var lblName = new Label();
            lblName.Text = "Shortcut Name:";
            lblName.Location = new Point(20, 20);
            lblName.Size = new Size(100, 23);
            lblName.ForeColor = _foregroundColor;
            this.Controls.Add(lblName);

            txtShortcutName = new TextBox();
            txtShortcutName.Location = new Point(130, 20);
            txtShortcutName.Size = new Size(320, 23);
            ApplyTextBoxTheme(txtShortcutName);
            this.Controls.Add(txtShortcutName);

            // Start In
            var lblStartIn = new Label();
            lblStartIn.Text = "Start In:";
            lblStartIn.Location = new Point(20, 60);
            lblStartIn.Size = new Size(100, 23);
            lblStartIn.ForeColor = _foregroundColor;
            this.Controls.Add(lblStartIn);

            cmbStartIn = new ColoredCombo();
            cmbStartIn.Location = new Point(130, 60);
            cmbStartIn.Size = new Size(320, 23);
            cmbStartIn.DropDownStyle = ComboBoxStyle.DropDown;
            ApplyComboBoxTheme(cmbStartIn);
            this.Controls.Add(cmbStartIn);

            // Target Path
            var lblTarget = new Label();
            lblTarget.Text = "Target:";
            lblTarget.Location = new Point(20, 100);
            lblTarget.Size = new Size(100, 23);
            lblTarget.ForeColor = _foregroundColor;
            this.Controls.Add(lblTarget);

            txtTargetPath = new TextBox();
            txtTargetPath.Location = new Point(130, 100);
            txtTargetPath.Size = new Size(270, 23);
            txtTargetPath.ReadOnly = true;
            ApplyTextBoxTheme(txtTargetPath);
            this.Controls.Add(txtTargetPath);

            btnBrowseTarget = new Button();
            btnBrowseTarget.Text = "...";
            btnBrowseTarget.Location = new Point(410, 100);
            btnBrowseTarget.Size = new Size(40, 23);
            btnBrowseTarget.Click += BtnBrowseTarget_Click;
            ApplyButtonTheme(btnBrowseTarget, false);
            this.Controls.Add(btnBrowseTarget);

            // Arguments
            var lblArgs = new Label();
            lblArgs.Text = "Arguments:";
            lblArgs.Location = new Point(20, 140);
            lblArgs.Size = new Size(100, 23);
            lblArgs.ForeColor = _foregroundColor;
            this.Controls.Add(lblArgs);

            txtArguments = new TextBox();
            txtArguments.Location = new Point(130, 140);
            txtArguments.Size = new Size(320, 23);
            ApplyTextBoxTheme(txtArguments);
            this.Controls.Add(txtArguments);

            // Location
            var lblLocation = new Label();
            lblLocation.Text = "Location:";
            lblLocation.Location = new Point(20, 180);
            lblLocation.Size = new Size(100, 23);
            lblLocation.ForeColor = _foregroundColor;
            this.Controls.Add(lblLocation);

            cmbLocation = new ColoredCombo();
            cmbLocation.Location = new Point(130, 180);
            cmbLocation.Size = new Size(320, 23);
            cmbLocation.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbLocation.Items.Add("Start Menu (Current User)");
            cmbLocation.Items.Add("Start Menu (All Users)");
            cmbLocation.Items.Add("Startup (Current User)");
            cmbLocation.Items.Add("Startup (All Users)");
            cmbLocation.Items.Add("Desktop (Current User)");
            cmbLocation.Items.Add("Desktop (All Users)");
            cmbLocation.Items.Add("Custom Location...");
            cmbLocation.SelectedIndex = 0;
            ApplyComboBoxTheme(cmbLocation);
            this.Controls.Add(cmbLocation);

            // Run As
            var lblRunAs = new Label();
            lblRunAs.Text = "Run As:";
            lblRunAs.Location = new Point(20, 220);
            lblRunAs.Size = new Size(100, 23);
            lblRunAs.ForeColor = _foregroundColor;
            this.Controls.Add(lblRunAs);

            cmbRunAs = new ColoredCombo();
            cmbRunAs.Location = new Point(130, 220);
            cmbRunAs.Size = new Size(320, 23);
            cmbRunAs.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbRunAs.Items.Add("Normal window");
            cmbRunAs.Items.Add("Minimized");
            cmbRunAs.Items.Add("Maximized");
            cmbRunAs.SelectedIndex = 0;
            ApplyComboBoxTheme(cmbRunAs);
            this.Controls.Add(cmbRunAs);

            // Run as Administrator checkbox
            chkRunAsAdmin = new CheckBox();
            chkRunAsAdmin.Text = "Run as administrator";
            chkRunAsAdmin.Location = new Point(20, 285);
            chkRunAsAdmin.Size = new Size(150, 20);
            chkRunAsAdmin.Font = new Font("Segoe UI", 9F);
            chkRunAsAdmin.ForeColor = _foregroundColor;
            chkRunAsAdmin.BackColor = _backgroundColor;
            this.Controls.Add(chkRunAsAdmin);

            // Buttons
            btnCreate = new Button();
            btnCreate.Text = "Create Shortcut";
            btnCreate.Location = new Point(280, 280);
            btnCreate.Size = new Size(100, 30);
            btnCreate.Click += BtnCreate_Click;
            ApplyButtonTheme(btnCreate, true);
            this.Controls.Add(btnCreate);

            btnCancel = new Button();
            btnCancel.Text = "Cancel";
            btnCancel.Location = new Point(390, 280);
            btnCancel.Size = new Size(80, 30);
            btnCancel.Click += BtnCancel_Click;
            ApplyButtonTheme(btnCancel, false);
            this.Controls.Add(btnCancel);
        }

        private void ApplyTextBoxTheme(TextBox textBox)
        {
            textBox.BackColor = _controlBackgroundColor;
            textBox.ForeColor = _foregroundColor;
            if (_isDarkTheme)
            {
                textBox.BorderStyle = BorderStyle.FixedSingle;
            }
        }

        private void ApplyComboBoxTheme(ComboBox comboBox)
        {
            comboBox.BackColor = _controlBackgroundColor;
            comboBox.ForeColor = _foregroundColor;
            comboBox.FlatStyle = FlatStyle.Flat;
        }

        private void ApplyButtonTheme(Button button, bool isPrimary)
        {
            if (isPrimary)
            {
                button.BackColor = _accentColor;
                button.ForeColor = Color.White;
            }
            else
            {
                button.BackColor = _buttonBackgroundColor;
                button.ForeColor = _foregroundColor;
            }
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = _isDarkTheme ? 1 : 0;
            if (_isDarkTheme && !isPrimary)
            {
                button.FlatAppearance.BorderColor = Color.FromArgb(70, 70, 70);
            }
        }

        private void LoadDefaultValues()
        {
            if (!string.IsNullOrEmpty(TargetFile))
            {
                txtTargetPath.Text = TargetFile;
                txtShortcutName.Text = Path.GetFileNameWithoutExtension(TargetFile);
                cmbStartIn.Text = Path.GetDirectoryName(TargetFile);
            }
        }

        private void BtnBrowseTarget_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "All Files (*.*)|*.*";
                openFileDialog.Title = "Select Target File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtTargetPath.Text = openFileDialog.FileName;
                    if (string.IsNullOrEmpty(txtShortcutName.Text))
                    {
                        txtShortcutName.Text = Path.GetFileNameWithoutExtension(openFileDialog.FileName);
                    }
                    if (string.IsNullOrEmpty(cmbStartIn.Text))
                    {
                        cmbStartIn.Text = Path.GetDirectoryName(openFileDialog.FileName);
                    }
                }
            }
        }

        private void BtnCreate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtTargetPath.Text))
            {
                MessageBox.Show("Please select a target file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrEmpty(txtShortcutName.Text))
            {
                MessageBox.Show("Please enter a shortcut name.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Handle custom location selection
            if (cmbLocation.SelectedIndex == 6 && string.IsNullOrEmpty(_customLocationPath))
            {
                using (var folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "Select destination folder";
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        _customLocationPath = folderDialog.SelectedPath;
                    }
                    else
                    {
                        return; // User cancelled
                    }
                }
            }

            // Check if admin privileges are required
            if (RequiresAdminPrivileges(cmbLocation.SelectedIndex) && !IsRunningAsAdministrator())
            {
                var result = MessageBox.Show(
                    "Creating shortcuts in this location requires administrator privileges.\n\n" +
                    "Would you like to create the shortcut with elevated privileges?",
                    "Administrator Privileges Required",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    CreateShortcutWithElevation();
                }
                return;
            }

            try
            {
                CreateShortcut();
                MessageBox.Show("Shortcut created successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show(
                    "Access denied. You may need administrator privileges to create shortcuts in this location.\n\n" +
                    "Try running the application as administrator or choose a different location.",
                    "Access Denied",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating shortcut: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CreateShortcutWithElevation()
        {
            try
            {
                // Create a temporary script file to execute with elevation
                string tempScriptPath = Path.GetTempFileName() + ".ps1";
                string shortcutPath = GetShortcutPathForElevation();

                // Build PowerShell script to create the shortcut
                string script = @"
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut('" + shortcutPath + @"')
$Shortcut.TargetPath = '" + txtTargetPath.Text.Replace("'", "''") + @"'
$Shortcut.Arguments = '" + txtArguments.Text.Replace("'", "''") + @"'
$Shortcut.WorkingDirectory = '" + cmbStartIn.Text.Replace("'", "''") + @"'
$Shortcut.WindowStyle = " + GetWindowStyleValue() + @"
$Shortcut.Save()

" + (chkRunAsAdmin.Checked ? GetRunAsAdminScript(shortcutPath) : "") + @"

Write-Host 'Shortcut created successfully!'
";

                File.WriteAllText(tempScriptPath, script);

                // Execute PowerShell with elevation
                var startInfo = new ProcessStartInfo()
                {
                    FileName = "powershell.exe",
                    Arguments = "-ExecutionPolicy Bypass -File \"" + tempScriptPath + "\"",
                    UseShellExecute = true,
                    Verb = "runas",
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (var process = Process.Start(startInfo))
                {
                    process.WaitForExit();

                    if (process.ExitCode == 0)
                    {
                        MessageBox.Show("Shortcut created successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Failed to create shortcut with elevated privileges.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                // Clean up temporary script
                try { File.Delete(tempScriptPath); } catch { }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to create shortcut with elevation: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetShortcutPathForElevation()
        {
            string shortcutName = txtShortcutName.Text;
            if (!shortcutName.EndsWith(".lnk"))
                shortcutName += ".lnk";

            string basePath;
            switch (cmbLocation.SelectedIndex)
            {
                case 0: // Start Menu (Current User)
                    basePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "ShortcutIndexer");
                    break;
                case 1: // Start Menu (All Users)
                    basePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms), "ShortcutIndexer");
                    break;
                case 2: // Startup (Current User)
                    basePath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
                    break;
                case 3: // Startup (All Users)
                    basePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonStartup);
                    break;
                case 4: // Desktop (Current User)
                    basePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    break;
                case 5: // Desktop (All Users)
                    basePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory);
                    break;
                case 6: // Custom Location
                    basePath = _customLocationPath ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    break;
                default:
                    basePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    break;
            }

            return Path.Combine(basePath, shortcutName).Replace("'", "''");
        }

        private int GetWindowStyleValue()
        {
            switch (cmbRunAs.SelectedIndex)
            {
                case 0: return 1; // Normal
                case 1: return 7; // Minimized
                case 2: return 3; // Maximized
                default: return 1;
            }
        }

        private string GetRunAsAdminScript(string shortcutPath)
        {
            return @"
# Set Run as Administrator flag
$bytes = [System.IO.File]::ReadAllBytes('" + shortcutPath + @"')
if ($bytes.Length -gt 21) {
    $linkFlags = [System.BitConverter]::ToInt32($bytes, 20)
    $linkFlags = $linkFlags -bor 0x00002000
    $flagBytes = [System.BitConverter]::GetBytes($linkFlags)
    [System.Array]::Copy($flagBytes, 0, $bytes, 20, 4)
    [System.IO.File]::WriteAllBytes('" + shortcutPath + @"', $bytes)
}
";
        }

        private void CreateShortcut()
        {
            string shortcutPath = GetShortcutPath();

            // Create the shortcut using IWshRuntimeLibrary
            Type shellType = Type.GetTypeFromProgID("WScript.Shell");
            dynamic shell = Activator.CreateInstance(shellType);
            var shortcut = shell.CreateShortcut(shortcutPath);

            shortcut.TargetPath = txtTargetPath.Text;
            shortcut.Arguments = txtArguments.Text;
            shortcut.WorkingDirectory = cmbStartIn.Text;

            // Set window style based on selection
            switch (cmbRunAs.SelectedIndex)
            {
                case 0: shortcut.WindowStyle = 1; break; // Normal
                case 1: shortcut.WindowStyle = 7; break; // Minimized
                case 2: shortcut.WindowStyle = 3; break; // Maximized
            }

            shortcut.Save();

            // Set "Run as administrator" property if checkbox is checked
            if (chkRunAsAdmin.Checked)
            {
                SetShortcutRunAsAdmin(shortcutPath);
            }

            // Update Windows Search Index
            UpdateSearchIndex(shortcutPath);
        }

        private string GetShortcutPath()
        {
            string shortcutName = txtShortcutName.Text;
            if (!shortcutName.EndsWith(".lnk"))
                shortcutName += ".lnk";

            string basePath;
            switch (cmbLocation.SelectedIndex)
            {
                case 0: // Start Menu (Current User)
                    basePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "ShortcutIndexer");
                    break;
                case 1: // Start Menu (All Users)
                    basePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms), "ShortcutIndexer");
                    break;
                case 2: // Startup (Current User)
                    basePath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
                    break;
                case 3: // Startup (All Users)
                    basePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonStartup);
                    break;
                case 4: // Desktop (Current User)
                    basePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    break;
                case 5: // Desktop (All Users)
                    basePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory);
                    break;
                case 6: // Custom Location
                    if (string.IsNullOrEmpty(_customLocationPath))
                    {
                        using (var folderDialog = new FolderBrowserDialog())
                        {
                            folderDialog.Description = "Select destination folder";
                            if (folderDialog.ShowDialog() == DialogResult.OK)
                            {
                                basePath = folderDialog.SelectedPath;
                                _customLocationPath = basePath;
                            }
                            else
                            {
                                throw new OperationCanceledException("No destination folder selected.");
                            }
                        }
                    }
                    else
                    {
                        basePath = _customLocationPath;
                    }
                    break;
                default:
                    basePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    break;
            }

            // Create the directory if it doesn't exist (especially for ShortcutIndexer subfolders)
            if (!Directory.Exists(basePath))
            {
                Directory.CreateDirectory(basePath);
            }

            return Path.Combine(basePath, shortcutName);
        }

        [DllImport("shell32.dll")]
        private static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        }

        private void UpdateSearchIndex(string shortcutPath)
        {
            // Notify Windows that a new file has been created
            SHChangeNotify(0x00002000, 0x0000, IntPtr.Zero, IntPtr.Zero); // SHCNE_ASSOCCHANGED

            // Force a refresh of the search index
            try
            {
                Process.Start("searchindexer.exe", "/reindex");
            }
            catch
            {
                // Ignore if search indexer is not available
            }
        }

        private void SetShortcutRunAsAdmin(string shortcutPath)
        {
            try
            {
                // Read the shortcut file and modify its properties
                byte[] shortcutBytes = File.ReadAllBytes(shortcutPath);

                // Set the "Run as Administrator" flag by modifying the shortcut file
                // This sets the SLDF_RUNAS_USER flag in the LinkFlags field
                if (shortcutBytes.Length > 21)
                {
                    // The LinkFlags field is at offset 20-23 (4 bytes)
                    // SLDF_RUNAS_USER flag is 0x00002000
                    int linkFlags = BitConverter.ToInt32(shortcutBytes, 20);
                    linkFlags |= 0x00002000; // Set the SLDF_RUNAS_USER flag
                    byte[] flagBytes = BitConverter.GetBytes(linkFlags);
                    Array.Copy(flagBytes, 0, shortcutBytes, 20, 4);

                    // Write the modified bytes back to the file
                    File.WriteAllBytes(shortcutPath, shortcutBytes);
                }
            }
            catch (Exception ex)
            {
                // If direct file modification fails, show a warning but don't fail completely
                MessageBox.Show(
                    "Shortcut created successfully, but couldn't set 'Run as administrator' property: " + ex.Message +
                    "\n\nYou can manually set this by right-clicking the shortcut, selecting Properties, clicking Advanced, and checking 'Run as administrator'.",
                    "Warning",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }
    }

    public class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string targetFile = args.Length > 0 ? args[0] : null;
            Application.Run(new MainForm(targetFile));
        }
    }
}
