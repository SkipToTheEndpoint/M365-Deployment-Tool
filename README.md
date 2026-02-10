# Microsoft 365 Apps Deployment Helper

A simple PowerShell-based tool with optional GUI to build Microsoft 365 deployment packages.

---

## 🧰 Requirements

- Windows 10 or 11
- PowerShell 5.1 or later

To use the optional GUI:
- Node.js LTS from https://nodejs.org

---

## ▶️ Option 1: Use the PowerShell Script

1. Open a PowerShell window.
2. Navigate to the folder.
3. Run:

```powershell
.\Invoke-M365AppsHelper.ps1 -ConfigXML .\samples\sample1.xml -OnlineMode
```

This will create a lightweight Microsoft 365 Apps deployment package.

### Online Mode vs Full Download

- `-OnlineMode` creates a lightweight stub package (just `setup.exe` and metadata). The Office binaries are downloaded during install.
- If `-OnlineMode` is not passed, the script downloads **all Office data files** into the package. This can be preferred for fully offline installs, but results in **very large downloads** (multi-GB).
- For PMPC custom apps without `-OnlineMode`, you can use `-Compress` to zip Office data files and the script will also generate a **PreScript** to extract them during install.

> ⚠️ **Note:** The sample XML is for example purposes only.  
> You should generate your own configuration at [https://config.office.com](https://config.office.com) and export the XML.

### What the sample installs:

- **Microsoft 365 Apps for Enterprise (EEA No Teams)**
- **Visio Professional**
- Architecture: **64-bit**
- Update Channel: **Monthly Enterprise**
- Languages: `en-gb`, `fi-fi`, and `MatchPreviousMSI`
- Excludes: **OneDrive for Business (Groove)** and **Skype for Business (Lync)**
- MSI Removal: Enabled (removes previous MSI-based Office)
- Updates: Enabled
- Display: **Silent install**, no UI (`Display Level="None"`)
- Accept EULA: **Not accepted automatically**

> `-OnlineMode` creates a stub-based package (just `setup.exe` and metadata).  
> If not passed, the script downloads **all Office binaries** and packages them.

---

## 🖥️ Option 2: Use the GUI (Electron)

1. Install [Node.js LTS](https://nodejs.org)
2. Open a terminal in the folder.
3. Run:

```bash
npm install
npm start
```

The GUI will launch. This may take a minute the first time.

---

## 💼 Intune + Patch My PC Support

This tool is ideal for:

### ✅ **Microsoft Intune**

You can use your exported XML file from [config.office.com](https://config.office.com) and:

- Automatically generate a ready-to-upload `.intunewin` package
- Includes:
  - `setup.exe`
  - install.xml and uninstall.xml generated from your provided xml
  - Optional full Office content (if not using `-OnlineMode`)
- A detection script
- Microsoft Logo
- Details txt file containing all the information required to populate the Win32 app steps

The package is structured for seamless deployment via Microsoft Intune.

When generating Win32 packages (non-PMPC custom apps), the tool also produces a **Win32 app details** `.txt` file that tells you exactly what to fill in during Intune app creation (install/uninstall commands, detection, and metadata). The Microsoft logo file is downloaded automatically for use as the app icon.

### ✅ **Patch My PC Custom App**

If you're a PMPC Cloud customer, add the `-PMPCCustomApp` parameter:

```powershell
.\Invoke-M365AppsHelper.ps1 -ConfigXML .\your.xml -PMPCCustomApp
```

This will generate all files and metadata required to upload a **custom Office installer** to the Patch My PC Cloud portal, including:

- Office source content
- Optional compressed Office content and **PreScript** when `-Compress` is used without `-OnlineMode`

---

## ⚙️ Parameters (Detailed)

- **ConfigXML**: Optional. Path to the Office configuration XML. If omitted, the script auto-detects a single XML in the script folder.
- **BasePath**: Optional. Default: `$env:APPDATA\M365AppsHelper`. Root for output folders (`Packages`, `Downloads`, `Logs`).
- **SetupUrl**: Optional. Default: `https://officecdn.microsoft.com/pr/wsus/setup.exe`.
- **OfficeVersionUrl**: Optional. Default: `https://clients.config.office.net/releases/v1.0/OfficeReleases`.
- **OfficeIconUrl**: Optional. Default: `https://www.svgrepo.com/show/452062/microsoft.svg` (downloaded as `Microsoft.png`).
- **LogName**: Optional. Default: `Invoke-M365AppsHelper.log`. Can be a file name or a full path.
- **Win32ContentPrepToolUrl**: Optional. Default: `https://raw.githubusercontent.com/microsoft/Microsoft-Win32-Content-Prep-Tool/master/IntuneWinAppUtil.exe`.
- **CreateIntuneWin**: Switch. Generate a `.intunewin` package (non-PMPC custom app).
- **NoZip**: Switch. Skip creating `Office.zip` when Office content is downloaded.
- **OnlineMode**: Switch. Build a stub package only (no Office binaries). Without this switch, full Office data files are downloaded (large).
- **SkipAPICheck**: Switch. Skip Office API validation (only works if a version is set in the XML).
- **PMPCCustomApp**: Switch. Generate PMPC custom app output instead of standard Win32 output.
- **ApiRetryDelaySeconds**: Optional. Default: `3`. Delay in seconds between API retries. Range: `1-30`.
- **ApiMaxExtendedAttempts**: Optional. Default: `10`. Maximum API retry attempts. Range: `1-20`.

---

## ℹ️ About This Script

- **Created by**: Ben Whitmore @ Patch My PC  
- **Filename**: `Invoke-M365AppsHelper.ps1`  
- **Created on**: 07/09/2025  
- **Updated on**: 25/01/2025

This script automates the process of creating Microsoft 365 Office deployment packages by:

- Parsing Office configuration XML files (no hardcoded values)
- Downloading the required setup files
- Creating organized deployment packages
- Supporting zip packaging with optional PreScript
- Validating Office versions using Microsoft’s public Office REST API

### 🧪 Version Validation

If your XML defines a version, the script will:

- Query Microsoft’s API to confirm the version is valid for the specified channel
- Prevent deployment failures from invalid/nonexistent Office builds

If no version is specified, the latest version for the channel will be used automatically.

The script also:

- Retries API queries when needed (e.g., rate limiting, incomplete responses)
- Helps ensure builds are accurate and up-to-date

---

## Notes

- No admin rights needed
- All processing is local
- No data is sent externally

---

## License

GPL-3.0 — see LICENSE
