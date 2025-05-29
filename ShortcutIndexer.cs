using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Security.Principal;
using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace ShortcutIndexer
{
    public partial class MainForm : Form
    {
        private TextBox txtShortcutName;
        private ComboBox cmbLocation;
        private TextBox txtTargetPath;
        private TextBox txtArguments;
        private ComboBox cmbStartIn;
        private ComboBox cmbRunAs;
        private Button btnCreate;
        private Button btnCancel;
        private Button btnBrowseTarget;
        private CheckBox chkRunAsAdmin;
        private Label lblIcon;

        public string TargetFile { get; set; }

        public MainForm(string targetFile = null)
        {
            TargetFile = targetFile;
            InitializeComponent();
            LoadDefaultValues();
        }

        private void InitializeComponent()
        {
            this.Text = "Shortcut Indexer - Create Shortcut";
            this.Size = new Size(500, 360);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;

            // Modern Windows 11 styling
            this.Font = new Font("Segoe UI", 9F);

            // Shortcut Name
            var lblName = new Label();
            lblName.Text = "Shortcut Name:";
            lblName.Location = new Point(20, 20);
            lblName.Size = new Size(100, 23);
            this.Controls.Add(lblName);

            txtShortcutName = new TextBox();
            txtShortcutName.Location = new Point(130, 20);
            txtShortcutName.Size = new Size(320, 23);
            this.Controls.Add(txtShortcutName);

            // Start In
            var lblStartIn = new Label();
            lblStartIn.Text = "Start In:";
            lblStartIn.Location = new Point(20, 60);
            lblStartIn.Size = new Size(100, 23);
            this.Controls.Add(lblStartIn);

            cmbStartIn = new ComboBox();
            cmbStartIn.Location = new Point(130, 60);
            cmbStartIn.Size = new Size(320, 23);
            cmbStartIn.DropDownStyle = ComboBoxStyle.DropDown;
            this.Controls.Add(cmbStartIn);

            // Target Path
            var lblTarget = new Label();
            lblTarget.Text = "Target:";
            lblTarget.Location = new Point(20, 100);
            lblTarget.Size = new Size(100, 23);
            this.Controls.Add(lblTarget);

            txtTargetPath = new TextBox();
            txtTargetPath.Location = new Point(130, 100);
            txtTargetPath.Size = new Size(270, 23);
            txtTargetPath.ReadOnly = true;
            this.Controls.Add(txtTargetPath);

            btnBrowseTarget = new Button();
            btnBrowseTarget.Text = "...";
            btnBrowseTarget.Location = new Point(410, 100);
            btnBrowseTarget.Size = new Size(40, 23);
            btnBrowseTarget.Click += BtnBrowseTarget_Click;
            this.Controls.Add(btnBrowseTarget);

            // Arguments
            var lblArgs = new Label();
            lblArgs.Text = "Arguments:";
            lblArgs.Location = new Point(20, 140);
            lblArgs.Size = new Size(100, 23);
            this.Controls.Add(lblArgs);

            txtArguments = new TextBox();
            txtArguments.Location = new Point(130, 140);
            txtArguments.Size = new Size(320, 23);
            this.Controls.Add(txtArguments);

            // Location
            var lblLocation = new Label();
            lblLocation.Text = "Location:";
            lblLocation.Location = new Point(20, 180);
            lblLocation.Size = new Size(100, 23);
            this.Controls.Add(lblLocation);

            cmbLocation = new ComboBox();
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
            this.Controls.Add(cmbLocation);

            // Run As
            var lblRunAs = new Label();
            lblRunAs.Text = "Run As:";
            lblRunAs.Location = new Point(20, 220);
            lblRunAs.Size = new Size(100, 23);
            this.Controls.Add(lblRunAs);

            cmbRunAs = new ComboBox();
            cmbRunAs.Location = new Point(130, 220);
            cmbRunAs.Size = new Size(320, 23);
            cmbRunAs.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbRunAs.Items.Add("Normal window");
            cmbRunAs.Items.Add("Minimized");
            cmbRunAs.Items.Add("Maximized");
            cmbRunAs.SelectedIndex = 0;
            this.Controls.Add(cmbRunAs);

            // Run as Administrator checkbox
            chkRunAsAdmin = new CheckBox();
            chkRunAsAdmin.Text = "Run as administrator";
            chkRunAsAdmin.Location = new Point(20, 285);
            chkRunAsAdmin.Size = new Size(150, 20);
            chkRunAsAdmin.Font = new Font("Segoe UI", 9F);
            this.Controls.Add(chkRunAsAdmin);

            // Buttons
            btnCreate = new Button();
            btnCreate.Text = "Create Shortcut";
            btnCreate.Location = new Point(280, 280);
            btnCreate.Size = new Size(100, 30);
            btnCreate.BackColor = Color.FromArgb(0, 120, 215);
            btnCreate.ForeColor = Color.White;
            btnCreate.FlatStyle = FlatStyle.Flat;
            btnCreate.FlatAppearance.BorderSize = 0;
            btnCreate.Click += BtnCreate_Click;
            this.Controls.Add(btnCreate);

            btnCancel = new Button();
            btnCancel.Text = "Cancel";
            btnCancel.Location = new Point(390, 280);
            btnCancel.Size = new Size(80, 30);
            btnCancel.FlatStyle = FlatStyle.Flat;
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);
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

            try
            {
                CreateShortcut();
                MessageBox.Show("Shortcut created successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }            catch (Exception ex)
            {
                MessageBox.Show("Error creating shortcut: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
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
                    using (var folderDialog = new FolderBrowserDialog())
                    {
                        folderDialog.Description = "Select destination folder";
                        if (folderDialog.ShowDialog() == DialogResult.OK)
                        {
                            basePath = folderDialog.SelectedPath;
                        }
                        else
                        {
                            throw new OperationCanceledException("No destination folder selected.");
                        }
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
