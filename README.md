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
  - Your XML
  - A generated detection script
  - Optional full Office content (if not using `-OnlineMode`)

The package is structured for seamless deployment via Microsoft Intune.

### ✅ **Patch My PC Custom App**

If you're a PMPC Cloud customer, add the `-PMPCCustomApp` parameter:

```powershell
.\Invoke-M365AppsHelper.ps1 -ConfigXML .\your.xml -PMPCCustomApp
```

This will generate all files and metadata required to upload a **custom Office installer** to the Patch My PC Cloud portal, including:

- Custom app manifest
- Office source content
- Detection script

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
