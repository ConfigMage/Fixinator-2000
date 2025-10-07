# The Fixinator 2000

The Fixinator 2000 is a comprehensive remote support toolkit designed to empower IT professionals with streamlined automation and deployment capabilities.

## About

Built during the 2025 DOR Hackathon by team "The Last Picks":
- Andrew Fredrickson
- Edgar Pozos
- William Gorham
- Chris Solario

**Fixing problems since Y2K!**

## Features

### 🖥️ Software Center
Deploy and install software remotely with ease. Connect to remote machines and manage software installations from a centralized interface.

### 🔧 PowerShell Toolbox
Access essential local utilities for system management:
- **Defender Toolkit** - Security and threat management tools
- **Generate a GPO Report** - Group Policy documentation
- **Generate an AD Lockout Report** - Active Directory troubleshooting
- **Printer Tool 3000** - Print management utilities
- **Lifecycle Migration Tool** - User data migration assistance

### 📜 Script Library
Create, store, and execute custom automation scripts. Build your own library of PowerShell solutions with syntax highlighting and easy execution.

## Requirements

- Windows PowerShell 5.1 or later
- Windows operating system
- Administrator privileges (for certain features)
- .NET Framework 4.5 or later

## Installation

1. Clone this repository
2. Ensure PowerShell execution policy allows script execution:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
3. Run `RemoteHelpDeskTool.ps1`

## Authentication Methods

- **Smart Card** - Use CAC/PIV card authentication
- **Username/Password** - Standard credential authentication
- **Current User** - Use current Windows credentials

## Directory Structure

```
Remote_tool final/
├── RemoteHelpDeskTool.ps1    # Main application
├── Modules/                   # Feature modules
│   ├── Auth.psm1
│   ├── ConfigManager.psm1
│   ├── ScriptLibrary.psm1
│   ├── SoftwareCenter.psm1
│   └── Toolbox.psm1
├── ToolBox/                   # Toolbox PowerShell scripts
├── Config/                    # Configuration files
└── Logs/                      # Application logs
```

## License

© 2025 DOR Hackathon - Team "The Last Picks"

## Version

Version 2000.0
