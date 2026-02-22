# Microsoft 365 Apps Deployment Helper

A simple PowerShell-based tool with optional GUI to build Microsoft 365 deployment packages.

---

## Requirements

- PowerShell 5.1 or later

To use the optional GUI:

- Node.js from https://nodejs.org

---

## Option 1: Use the PowerShell Script

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
- For PMPC custom apps without `-OnlineMode`, Office data files will be compressed automatically and the script will also generate a **PreScript** to extract them during install. Use the `-NoZip` parameter if you do not want to compress the office data files.

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

## Option 2: Use the GUI (Electron)

1. Install [Node.js LTS](https://nodejs.org)
2. Download the latest release zip and unpack to a folder.
3. Open PowerShell, change directory to the same folder.
4. Run:

```bash
npm install
npm start
```

The GUI will launch. This may take a minute the first time as node modules are downloaded.

---

## 🤷 Who is this tool for?

This tool is ideal for:

### ✅ **Microsoft Intune**

You can use your exported XML file from [config.office.com](https://config.office.com) and:

- Automatically generate a ready-to-upload `.intunewin` package
- Includes:
  - `setup.exe`
  - install.xml and uninstall.xml generated from your provided xml
  - Optional full Office content (if not using `-OnlineMode`)
- Detection Script (Leveraging DisplayName and DisplayVersion)
- Microsoft Logo
- Win32 app details.txt file that tells you exactly what to fill in during Intune app creation (install/uninstall commands, detection, and metadata)

### ✅ **Patch My PC Custom App**

If you're a Patch My PC customer, add the `-PMPCCustomApp` parameter:

```powershell
.\Invoke-M365AppsHelper.ps1 -ConfigXML .\your.xml -PMPCCustomApp
```

This will generate all files and metadata required to create a **Custom App** in the Patch My PC Cloud portal, including:

- Includes:
  - `setup.exe`
  - install.xml and uninstall.xml generated from your provided xml
  - Optional full Office content (if not using `-OnlineMode`)
  - Compressed Office content and **PreScript** (if not using `-OnlineMode`). Use the `-NoZip` parameter if you do not want to compress the office data files.
---

## ⚙️ Parameters

**ConfigXML**
Path to the Office configuration XML. If omitted, the script auto-detects a single XML in the script folder.
```
Type: String
Required: False
Default: Auto-detect single XML
Incompatible with: None
```

**BasePath**
Root path for output folders (`Packages`, `Downloads`, `Logs`).
```
Type: String
Required: False
Default: $env:APPDATA\M365AppsHelper
Incompatible with: None
```

**SetupUrl**
Office setup executable download URL.
```
Type: String
Required: False
Default: https://officecdn.microsoft.com/pr/wsus/setup.exe
Incompatible with: None
```

**OfficeVersionUrl**
Office version API endpoint used for validation.
```
Type: String
Required: False
Default: https://clients.config.office.net/releases/v1.0/OfficeReleases
Incompatible with: None
```

**OfficeIconUrl**
Icon download URL (saved as `Microsoft.png`).
```
Type: String
Required: False
Default: https://www.svgrepo.com/show/452062/microsoft.svg
Incompatible with: None
```

**LogName**
Log file name or full path.
```
Type: String
Required: False
Default: Invoke-M365AppsHelper.log
Incompatible with: None
```

**Win32ContentPrepToolUrl**
Win32 Content Prep Tool URL (`IntuneWinAppUtil.exe`).
```
Type: String
Required: False
Default: https://raw.githubusercontent.com/microsoft/Microsoft-Win32-Content-Prep-Tool/master/IntuneWinAppUtil.exe
Incompatible with: None
```

**CreateIntuneWin**
Generate a `.intunewin` package (non-PMPC custom app).
```
Type: Switch
Required: False
Default: False
Incompatible with: NoZip, PMPCCustomApp
```

**NoZip**
Skip creating `Office.zip` when Office content is downloaded.
```
Type: Switch
Required: False
Default: False
Incompatible with: CreateIntuneWin
```

**OnlineMode**
Build a stub package only (no Office binaries). Without this switch, full Office data files are downloaded (large).
```
Type: Switch
Required: False
Default: False
Incompatible with: None
```

**SkipAPICheck**
Skip Office API validation (only works if a version is set in the XML).
```
Type: Switch
Required: False
Default: False
Incompatible with: None
```

**PMPCCustomApp**
Generate PMPC custom app output instead of standard Win32 output.
```
Type: Switch
Required: False
Default: False
Incompatible with: CreateIntuneWin
```

**ApiRetryDelaySeconds**
Delay in seconds between API retries.
```
Type: Integer
Required: False
Default: 3
Range: 1-30
Incompatible with: None
```

**ApiMaxExtendedAttempts**
Maximum API retry attempts.
```
Type: Integer
Required: False
Default: 10
Range: 1-20
Incompatible with: None
```

---

## ℹ️ About This Script

- **Created by**: Ben Whitmore @ Patch My PC  
- **Filename**: `Invoke-M365AppsHelper.ps1`  
- **Created on**: 07/09/2025  
- **Updated on**: 22/02/2025

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
