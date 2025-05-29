using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using Microsoft.Win32;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.Reflection;

// Assembly attributes for COM registration
[assembly: ComVisible(true)]
[assembly: AssemblyTitle("ShortcutIndexer Shell Extension")]
[assembly: AssemblyDescription("Context menu shell extension for creating shortcuts")]
[assembly: AssemblyVersion("1.0.0.0")]

namespace ShortcutIndexerHandler
{
    [ComVisible(true)]
    [Guid("A7B15F12-9876-4321-ABCD-123456789ABC")]
    [ClassInterface(ClassInterfaceType.None)]
    public class ShortcutContextMenuHandler : IShellExtInit, IContextMenu
    {
        private string selectedFile;
        private const uint CMD_JUSTME = 0;
        private const uint CMD_ALLUSERS = 1;
        private const uint CMD_ADVANCED = 2;

        #region IShellExtInit Implementation
        public int Initialize(IntPtr pidlFolder, IntPtr pDataObj, IntPtr hKeyProgID)
        {
            if (pDataObj == IntPtr.Zero)
                return -1; // E_FAIL

            try
            {
                // Get the file path from the data object using CF_HDROP format
                var dataObject = (IDataObject)Marshal.GetObjectForIUnknown(pDataObj);
                var formatEtc = new FORMATETC();
                formatEtc.cfFormat = 15; // CF_HDROP
                formatEtc.ptd = IntPtr.Zero;
                formatEtc.dwAspect = DVASPECT.DVASPECT_CONTENT;
                formatEtc.lindex = -1;
                formatEtc.tymed = TYMED.TYMED_HGLOBAL; var medium = new STGMEDIUM();
                int result = 0;
                try
                {
                    dataObject.GetData(ref formatEtc, out medium);
                    result = 0; // Assume success if no exception
                }
                catch
                {
                    result = -1; // Error
                }

                if (result == 0 && medium.unionmember != IntPtr.Zero)
                {
                    // Get the file path using DragQueryFile
                    var sb = new StringBuilder(260);
                    uint fileCount = DragQueryFile(medium.unionmember, 0xFFFFFFFF, null, 0);
                    if (fileCount > 0 && DragQueryFile(medium.unionmember, 0, sb, sb.Capacity) > 0)
                    {
                        selectedFile = sb.ToString();
                        ReleaseStgMedium(ref medium);
                        return 0; // S_OK
                    }
                    ReleaseStgMedium(ref medium);
                }
            }
            catch (Exception ex)
            {
                // Log error for debugging
                System.Diagnostics.Debug.WriteLine("ShortcutIndexer Initialize error: " + ex.Message);
            }

            return -1; // E_FAIL
        }
        #endregion

        #region IContextMenu Implementation
        public int QueryContextMenu(IntPtr hMenu, uint iMenu, uint idCmdFirst, uint idCmdLast, uint uFlags)
        {
            // Skip if this is the default verb only
            if ((uFlags & 0x000F) == 0x000F) // CMF_DEFAULTONLY
                return 0;

            // Skip if no file is selected
            if (string.IsNullOrEmpty(selectedFile))
                return 0;

            try
            {
                // Create submenu
                IntPtr hSubmenu = CreatePopupMenu();
                if (hSubmenu == IntPtr.Zero)
                    return 0;

                // Add menu items to submenu
                AppendMenu(hSubmenu, 0x00000000, idCmdFirst + CMD_JUSTME, "&Just for me");
                AppendMenu(hSubmenu, 0x00000000, idCmdFirst + CMD_ALLUSERS, "&For all users");
                AppendMenu(hSubmenu, 0x00000800, 0, null); // MF_SEPARATOR
                AppendMenu(hSubmenu, 0x00000000, idCmdFirst + CMD_ADVANCED, "&Additional options...");

                // Add the main menu item with submenu
                InsertMenu(hMenu, iMenu, 0x00000400 | 0x00000010, (uint)hSubmenu, "Shortcut Indexer");

                return 4; // Return number of commands added
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("QueryContextMenu error: " + ex.Message);
                return 0;
            }
        }

        public int InvokeCommand(IntPtr pici)
        {
            try
            {
                var ici = (CMINVOKECOMMANDINFO)Marshal.PtrToStructure(pici, typeof(CMINVOKECOMMANDINFO));
                uint commandId = (uint)ici.lpVerb.ToInt32();

                switch (commandId)
                {
                    case CMD_JUSTME:
                        CreateQuickShortcut(false);
                        break;
                    case CMD_ALLUSERS:
                        CreateQuickShortcut(true);
                        break;
                    case CMD_ADVANCED:
                        ShowAdvancedDialog();
                        break;
                }

                return 0; // S_OK
            }
            catch
            {
                return -1; // E_FAIL
            }
        }

        public int GetCommandString(uint idCmd, uint uFlags, IntPtr pwReserved, StringBuilder commandString, int cchMax)
        {
            switch (idCmd)
            {
                case CMD_JUSTME:
                    commandString.Clear();
                    commandString.Append("Create shortcut for current user only");
                    break;
                case CMD_ALLUSERS:
                    commandString.Clear();
                    commandString.Append("Create shortcut for all users");
                    break;
                case CMD_ADVANCED:
                    commandString.Clear();
                    commandString.Append("Open advanced shortcut options");
                    break;
                default:
                    return -1; // E_FAIL
            }
            return 0; // S_OK
        }
        #endregion

        #region Helper Methods
        private void CreateQuickShortcut(bool allUsers)
        {
            if (string.IsNullOrEmpty(selectedFile) || !File.Exists(selectedFile))
                return;

            try
            {
                string startMenuPath;

                if (allUsers)
                {
                    // Check if we have admin privileges
                    if (!IsRunAsAdministrator())
                    {
                        // Prompt user to run with elevation
                        var result = System.Windows.Forms.MessageBox.Show(
                            "Creating shortcuts for all users requires administrator privileges.\n\n" +
                            "Would you like to restart with administrator privileges?",
                            "Administrator Required",
                            System.Windows.Forms.MessageBoxButtons.YesNo,
                            System.Windows.Forms.MessageBoxIcon.Question);

                        if (result == System.Windows.Forms.DialogResult.Yes)
                        {
                            // Launch elevated process to create the shortcut
                            LaunchElevatedShortcutCreation();
                            return;
                        }
                        else
                        {
                            return; // User declined elevation
                        }
                    }
                    startMenuPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms);
                }
                else
                {
                    startMenuPath = Environment.GetFolderPath(Environment.SpecialFolder.Programs);
                }

                // Create ShortcutIndexer subfolder
                string shortcutIndexerFolder = Path.Combine(startMenuPath, "ShortcutIndexer");
                if (!Directory.Exists(shortcutIndexerFolder))
                {
                    Directory.CreateDirectory(shortcutIndexerFolder);
                }

                string shortcutName = Path.GetFileNameWithoutExtension(selectedFile) + ".lnk";
                string shortcutPath = Path.Combine(shortcutIndexerFolder, shortcutName);

                // Create shortcut
                Type shellType = Type.GetTypeFromProgID("WScript.Shell");
                dynamic shell = Activator.CreateInstance(shellType);
                var shortcut = shell.CreateShortcut(shortcutPath);

                shortcut.TargetPath = selectedFile;
                shortcut.WorkingDirectory = Path.GetDirectoryName(selectedFile);
                shortcut.Save();

                // Update search index
                UpdateSearchIndex(shortcutPath);                // Show success message
                string location = allUsers ? "All Users Start Menu\\ShortcutIndexer" : "Your Start Menu\\ShortcutIndexer";
                System.Windows.Forms.MessageBox.Show(
                    "Shortcut created successfully in " + location + "!",
                    "Success",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error creating shortcut: " + ex.Message,
                    "ShortcutIndexer Error", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private bool IsRunAsAdministrator()
        {
            try
            {
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }

        private void LaunchElevatedShortcutCreation()
        {
            try
            {
                // Create a simple batch script that will create the shortcut with elevation
                string tempScript = Path.GetTempFileName() + ".bat";
                string shortcutName = Path.GetFileNameWithoutExtension(selectedFile);
                string allUsersStartMenu = Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms);
                string shortcutIndexerFolder = Path.Combine(allUsersStartMenu, "ShortcutIndexer");
                string shortcutPath = Path.Combine(shortcutIndexerFolder, shortcutName + ".lnk");

                // Create PowerShell script content that creates the shortcut
                string scriptContent = "@echo off\n" +
                    "powershell -Command \"" +
                    "if (!(Test-Path '" + shortcutIndexerFolder.Replace("'", "''") + "')) { New-Item -ItemType Directory -Path '" + shortcutIndexerFolder.Replace("'", "''") + "' -Force }; " +
                    "$WshShell = New-Object -comObject WScript.Shell; " +
                    "$Shortcut = $WshShell.CreateShortcut('" + shortcutPath.Replace("'", "''") + "'); " +
                    "$Shortcut.TargetPath = '" + selectedFile.Replace("'", "''") + "'; " +
                    "$Shortcut.WorkingDirectory = '" + Path.GetDirectoryName(selectedFile).Replace("'", "''") + "'; " +
                    "$Shortcut.Save(); " +
                    "Write-Host 'Shortcut created successfully in All Users Start Menu ShortcutIndexer folder!'\"\n" +
                    "pause\n" +
                    "del \"" + tempScript + "\"";

                File.WriteAllText(tempScript, scriptContent);

                // Launch the script with elevation
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = tempScript;
                startInfo.Verb = "runas"; // This triggers UAC elevation
                startInfo.UseShellExecute = true;

                Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Failed to launch elevated process: " + ex.Message + "\n\n" +
                    "Please try using 'Additional options' to create the shortcut manually.",
                    "Elevation Failed",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        private void ShowAdvancedDialog()
        {
            try
            {
                string exePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    "ShortcutIndexer", "ShortcutIndexer.exe");
                if (File.Exists(exePath))
                {
                    Process.Start(exePath, "\"" + selectedFile + "\"");
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error launching ShortcutIndexer: " + ex.Message,
                    "ShortcutIndexer Error", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void UpdateSearchIndex(string shortcutPath)
        {
            try
            {
                // Notify shell of change
                SHChangeNotify(0x00002000, 0x0000, IntPtr.Zero, IntPtr.Zero);

                // Try to update search index
                Process.Start("searchindexer.exe", "/reindex");
            }
            catch
            {
                // Ignore errors
            }
        }
        #endregion

        #region Win32 API
        [DllImport("user32.dll")]
        private static extern IntPtr CreatePopupMenu();

        [DllImport("user32.dll")]
        private static extern bool AppendMenu(IntPtr hMenu, uint uFlags, uint uIDNewItem, string lpNewItem);

        [DllImport("user32.dll")]
        private static extern bool InsertMenu(IntPtr hMenu, uint uPosition, uint uFlags, uint uIDNewItem, string lpNewItem);

        [DllImport("shell32.dll")]
        private static extern uint DragQueryFile(IntPtr hDrop, uint iFile, StringBuilder lpszFile, int cch);

        [DllImport("ole32.dll")]
        private static extern void ReleaseStgMedium(ref STGMEDIUM pmedium);

        [DllImport("kernel32.dll")]
        private static extern short RegisterClipboardFormat(string lpszFormat);

        [DllImport("shell32.dll")]
        private static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);

        [StructLayout(LayoutKind.Sequential)]
        private struct CMINVOKECOMMANDINFO
        {
            public uint cbSize;
            public uint fMask;
            public IntPtr hwnd;
            public IntPtr lpVerb;
            public string lpParameters;
            public string lpDirectory;
            public int nShow;
            public uint dwHotKey;
            public IntPtr hIcon;
        }
        #endregion

        #region COM Registration
        [ComRegisterFunction]
        public static void RegisterServer(Type t)
        {
            try
            {
                // Register the shell extension
                var guid = t.GUID.ToString("B");
                // HKCR\*\shellex\ContextMenuHandlers\ShortcutIndexer
                using (var key = Registry.ClassesRoot.CreateSubKey(@"*\shellex\ContextMenuHandlers\ShortcutIndexer"))
                {
                    key.SetValue("", guid);
                }

                // Notify shell of changes
                SHChangeNotify(0x08000000, 0x1000, IntPtr.Zero, IntPtr.Zero); // SHCNE_ASSOCCHANGED
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to register shell extension: " + ex.Message);
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterServer(Type t)
        {
            try
            {
                // Remove registry entries
                Registry.ClassesRoot.DeleteSubKeyTree(@"*\shellex\ContextMenuHandlers\ShortcutIndexer", false);

                // Notify shell of changes
                SHChangeNotify(0x08000000, 0x1000, IntPtr.Zero, IntPtr.Zero); // SHCNE_ASSOCCHANGED
            }
            catch
            {
                // Ignore errors during unregistration
            }
        }
        #endregion
    }

    #region Interfaces
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("000214E8-0000-0000-C000-000000000046")]
    public interface IShellExtInit
    {
        [PreserveSig]
        int Initialize(IntPtr pidlFolder, IntPtr pDataObj, IntPtr hKeyProgID);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("000214E4-0000-0000-C000-000000000046")]
    public interface IContextMenu
    {
        [PreserveSig]
        int QueryContextMenu(IntPtr hMenu, uint iMenu, uint idCmdFirst, uint idCmdLast, uint uFlags);

        [PreserveSig]
        int InvokeCommand(IntPtr pici);

        [PreserveSig]
        int GetCommandString(uint idCmd, uint uFlags, IntPtr pwReserved, StringBuilder commandString, int cchMax);
    }
    #endregion
}
