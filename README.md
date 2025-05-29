# ShortcutIndexer

A Windows context menu extension that allows users to easily create shortcuts with advanced options and Windows Search Index integration.

## Features

- **Right-click Context Menu**: Right-click any file to see "Shortcut Indexer" submenu
- **Quick Options**:
  - "Just for me" - Creates shortcut in current user's Start Menu (ShortcutIndexer folder)
  - "For all users" - Creates shortcut in all users' Start Menu (ShortcutIndexer folder, requires admin)
  - "Additional options" - Opens advanced GUI for custom configuration
- **Advanced GUI**: Modern Windows 11-style interface with options for:
  - Custom shortcut name
  - Location selection (Start Menu, Startup, Desktop, Custom)
  - Command line arguments
  - Working directory
  - Window state (Normal, Minimized, Maximized)
  - **Run as administrator** option
- **Organized Structure**: All shortcuts are automatically organized in dedicated "ShortcutIndexer" subfolders
- **Windows Search Integration**: Automatically updates Windows Search Index for new shortcuts
- **Easy Installation/Uninstallation**: Single-file installer with complete uninstaller

## Installation

### Method 1: Installer Script (Recommended)

1. Download all source files to a folder
2. Right-click `Installer.bat` and select "Run as administrator"
3. Follow the installation prompts

### Method 2: Manual Installation

1. Compile the source files (see "Building from Source" section)
2. Copy compiled files to `%ProgramFiles%\ShortcutIndexer`
3. Register the shell extension using RegAsm
4. Add registry entries for context menu integration

## Requirements

- Windows 10 or Windows 11
- .NET Framework 4.0 or later (usually pre-installed)
- Administrator privileges for installation and shell extension registration
- RegAsm.exe (part of .NET Framework SDK - typically available on most Windows systems)

## File Structure

```
ShortcutIndexer/
├── ShortcutIndexer.cs         # Main GUI application source
├── ShortcutIndexerHandler.cs  # COM shell extension source
├── Installer.bat             # Installation script
├── Uninstaller.bat           # Standalone uninstaller
├── LICENSE                   # MIT License file
├── .gitignore                # Git ignore rules
└── README.md                 # This file
```

## How It Works

1. **Shell Extension**: `ShortcutIndexerHandler.dll` is a COM shell extension that adds the context menu
2. **GUI Application**: `ShortcutIndexer.exe` provides the advanced options interface with "Run as administrator" support
3. **Registry Integration**: Registers with Windows shell for context menu functionality
4. **Organized Structure**: Automatically creates and manages "ShortcutIndexer" subfolders for organization
5. **Elevation Handling**: Seamlessly handles UAC prompts for "For all users" shortcuts
6. **Search Index Update**: Notifies Windows Search to index new shortcuts immediately

## Usage

After installation:

1. **Quick Shortcuts**: Right-click any file → "Shortcut Indexer" → Choose option
   - Shortcuts are automatically organized in "ShortcutIndexer" subfolders within the Start Menu
2. **Advanced Options**: Right-click any file → "Shortcut Indexer" → "Additional options"
   - Choose from multiple location options (Start Menu, Startup, Desktop, Custom)
   - Enable "Run as administrator" for elevated shortcuts
   - Configure custom arguments and working directories
3. **Search**: New shortcuts are immediately searchable in Windows Search

## Uninstallation

### Method 1: Windows Settings

1. Open Windows Settings
2. Go to Apps
3. Find "ShortcutIndexer" and click Uninstall

### Method 2: Uninstaller Script

1. Run `%ProgramFiles%\ShortcutIndexer\Uninstall.bat` as administrator

### Method 3: Standalone Uninstaller

1. Run `Uninstaller.bat` as administrator

## Technical Details

- **Language**: C# with Windows API calls (.NET Framework 4.0)
- **COM Interface**: Implements `IShellExtInit` and `IContextMenu`
- **Registration**: Uses RegAsm.exe for .NET COM component registration
- **Registry Keys**:
  - `HKCR\*\shellex\ContextMenuHandlers\ShortcutIndexer`
  - `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved`
  - `HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\ShortcutIndexer`
- **Installation Path**: `%ProgramFiles%\ShortcutIndexer`
- **COM GUID**: `{A7B15F12-9876-4321-ABCD-123456789ABC}`

## Compatibility

- ✅ Windows 10
- ✅ Windows 11
- ✅ 32-bit and 64-bit systems
- ✅ Multiple user environments
- ✅ Domain and workgroup computers

## Shortcut Organization

ShortcutIndexer automatically organizes all shortcuts in dedicated subfolders:

### Quick Options ("Just for me" / "For all users"):

- **Current User**: `%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\ShortcutIndexer\`
- **All Users**: `%ProgramData%\Microsoft\Windows\Start Menu\Programs\ShortcutIndexer\`

### Advanced GUI Options:

- **Start Menu (Current User)**: `Programs\ShortcutIndexer\`
- **Start Menu (All Users)**: `All Users Programs\ShortcutIndexer\`
- **Startup (Current User)**: User's Startup folder (no subfolder)
- **Startup (All Users)**: All Users Startup folder (no subfolder)
- **Desktop**: User's or All Users Desktop (no subfolder)
- **Custom Location**: User-specified folder

This organization keeps shortcuts tidy and makes them easy to find in the Start Menu under "ShortcutIndexer".

## Security Features

- Requires administrator privileges for installation only
- No network access required
- No data collection or telemetry
- Open source code for transparency

## Troubleshooting

### Context menu not appearing

1. Restart Windows Explorer: `taskkill /f /im explorer.exe; Start-Sleep 2; Start-Process explorer.exe`
2. Verify shell extension is registered: Check registry key `HKCR\*\shellex\ContextMenuHandlers\ShortcutIndexer`
3. Verify shell extension is approved: Check registry key `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved`
4. Reboot the computer
5. Verify installation: Check if files exist in `%ProgramFiles%\ShortcutIndexer`

### Registration errors

- Ensure you have .NET Framework 4.0 or later installed
- Use RegAsm.exe instead of regsvr32.exe for .NET assemblies
- Run registration commands as administrator

### Permission errors

- Ensure you're running as administrator during installation
- Check Windows User Account Control (UAC) settings

### .NET Framework errors

- Install .NET Framework 4.0 or later from Microsoft's website
- Ensure Windows is up to date
- Verify RegAsm.exe is available at `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe`

## Building from Source

If you want to modify and rebuild:

```batch
# Set up paths (adjust for your system if needed)
set CSC_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
set REGASM_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe

# Compile main application
"%CSC_PATH%" /target:winexe /out:ShortcutIndexer.exe /reference:System.dll /reference:System.Windows.Forms.dll /reference:System.Drawing.dll ShortcutIndexer.cs

# Compile shell extension (COM library)
"%CSC_PATH%" /target:library /out:ShortcutIndexerHandler.dll /reference:System.dll /reference:System.Windows.Forms.dll ShortcutIndexerHandler.cs

# Register shell extension (.NET COM component)
"%REGASM_PATH%" ShortcutIndexerHandler.dll /codebase /tlb

# Add to approved shell extensions (requires admin)
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved" /v "{A7B15F12-9876-4321-ABCD-123456789ABC}" /t REG_SZ /d "ShortcutIndexer Context Menu" /f
```

### Build Requirements

- .NET Framework 4.0 SDK or later
- Visual Studio (optional, but helpful for debugging)
- Windows SDK (for advanced shell programming)
- Administrator privileges for registration

## License

This project is licensed under the MIT License - see below for details.

### MIT License

Copyright (c) 2025 Harel Zadok

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Known Issues

- Shell extension may not appear immediately after installation - restart Windows Explorer
- On some systems, the approved shell extensions registry entry must be added manually
- Debug output is written to Visual Studio Output window or DebugView utility
- "Run as administrator" flag setting uses direct file manipulation - may fail on some systems with strict file permissions (fallback instructions provided)

## Contributing

1. Fork the repository
2. Make your changes
3. Test thoroughly on different Windows versions
4. Submit a pull request with detailed description of changes

## Version History

- **v1.10**: Enhanced functionality and organization (Current)

  - **NEW**: "Run as administrator" checkbox in advanced GUI
  - **NEW**: Organized folder structure - all shortcuts created in dedicated "ShortcutIndexer" subfolders
  - **NEW**: Startup folder options for auto-run shortcuts
  - **IMPROVED**: Better location options in advanced GUI (Start Menu, Startup, Desktop, Custom)
  - **IMPROVED**: Automatic directory creation for organized structure
  - **IMPROVED**: Enhanced elevation handling for "For all users" option
  - **TECHNICAL**: Direct shortcut file manipulation for administrator privileges
  - **TECHNICAL**: Uses Programs folders instead of StartMenu root for better organization

- **v1.0**: Initial release with full functionality
  - Context menu integration with Windows Explorer
  - Three quick options: "Just for me", "For all users", "Additional options"
  - Advanced GUI for custom shortcut configuration
  - Windows Search Index integration
  - Complete installer and uninstaller
  - .NET Framework COM shell extension implementation
