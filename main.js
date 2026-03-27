const { app, BrowserWindow, Menu, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const fs = require('fs');

let mainWindow;
let currentProcess = null;

// Send messages to renderer safely
const sendToRenderer = (channel, payload) => {
  if (mainWindow && mainWindow.webContents) {
    mainWindow.webContents.send(channel, payload);
  }
};

// Get default paths for the app
const getDefaultPaths = () => {
  const basePath = path.join(app.getPath('userData'), '..', 'M365AppsHelper');
  return { outputPath: basePath };
};

const SCRIPT_PATH = path.join(__dirname, 'Invoke-M365AppsHelper.ps1');

const createWindow = async () => {
  mainWindow = new BrowserWindow({
    width: 1600,
    height: 900,
    minWidth: 1000,
    minHeight: 700,
    maxWidth: 1700,
    maxHeight: 1080,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(__dirname, 'preload.js')
    },
    icon: path.join(__dirname, 'src', 'assets', 'icon.png')
  });

  mainWindow.loadFile(path.join(__dirname, 'src', 'index.html'));

  mainWindow.on('closed', () => {
    mainWindow = null;
    if (currentProcess && !currentProcess.killed) {
      currentProcess.kill();
    }
  });
};

app.on('ready', () => {
  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (mainWindow === null) {
    createWindow();
  }
});

// Remove default menu
Menu.setApplicationMenu(null);

// IPC Handlers - moved after app ready to ensure mainWindow exists
function setupIPCHandlers() {
  ipcMain.handle('dialog:openFile', async () => {
    if (!mainWindow) return null;
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openFile'],
      filters: [
        { name: 'XML Files', extensions: ['xml'] },
        { name: 'All Files', extensions: ['*'] }
      ]
    });
    return result.filePaths[0] || null;
  });

  ipcMain.handle('dialog:openFolder', async () => {
    if (!mainWindow) return null;
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openDirectory']
    });
    return result.filePaths[0] || null;
  });

  ipcMain.handle('get:defaultPaths', () => {
    return getDefaultPaths();
  });

  ipcMain.handle('open:folder', async (event, folderPath) => {
    try {
      await shell.openPath(folderPath);
      return { success: true };
    } catch (error) {
      return { success: false, error: error.message };
    }
  });

  ipcMain.handle('list:xmlFiles', () => {
    try {
      const currentDir = process.cwd();
      const files = fs.readdirSync(currentDir)
        .filter((f) => f.endsWith('.xml'))
        .map((f) => ({ name: f, path: path.join(currentDir, f) }));
      return files;
    } catch (error) {
      return [];
    }
  });

  ipcMain.handle('execute:deployment', async (event, payload) => {
    const {
      configXml,
      outputPath,
      logName,
      noZip,
      onlineMode,
      skipAPICheck,
      pmpcCustomApp,
      createIntuneWin,
      apiRetryDelaySeconds,
      apiMaxExtendedAttempts,
      setupUrl,
      officeVersionUrl,
      officeIconUrl
    } = payload || {};

    // Enforce flags: NoZip only valid if PMPCCustomApp checked and OnlineMode NOT active
    const effectiveNoZip = Boolean(pmpcCustomApp && noZip && !onlineMode);
    const effectiveCreateIntuneWin = Boolean(createIntuneWin);

    sendToRenderer('log', { message: 'Starting Office deployment package creation...', severity: 'info' });

    const workingDir = path.dirname(SCRIPT_PATH);
    const resolvePathIfNeeded = (p) => {
      if (!p || p === 'undefined' || !p.trim()) return null;
      return path.isAbsolute(p) ? p : path.resolve(workingDir, p);
    };

    let args = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', SCRIPT_PATH];

    const cfgPath = resolvePathIfNeeded(configXml);
    const logPath = resolvePathIfNeeded(logName);
    const basePath = resolvePathIfNeeded(outputPath);

    if (basePath) {
      args.push('-BasePath', basePath);
    }
    if (cfgPath) {
      args.push('-ConfigXML', cfgPath);
    }
    if (logPath) {
      args.push('-LogName', logPath);
    }
    if (setupUrl) args.push('-SetupUrl', setupUrl);
    if (officeVersionUrl) args.push('-OfficeVersionUrl', officeVersionUrl);
    if (officeIconUrl) args.push('-OfficeIconUrl', officeIconUrl);
    if (apiRetryDelaySeconds) args.push('-ApiRetryDelaySeconds', apiRetryDelaySeconds.toString());
    if (apiMaxExtendedAttempts) args.push('-ApiMaxExtendedAttempts', apiMaxExtendedAttempts.toString());
    if (effectiveNoZip) args.push('-NoZip');
    if (onlineMode) args.push('-OnlineMode');
    if (skipAPICheck) args.push('-SkipAPICheck');
    if (pmpcCustomApp) args.push('-PMPCCustomApp');
    if (effectiveCreateIntuneWin) args.push('-CreateIntuneWin');

    const psExecutable = 'pwsh.exe';
    const ps = spawn(psExecutable, args, {
      stdio: ['pipe', 'pipe', 'pipe'],
      shell: false,
      cwd: path.dirname(SCRIPT_PATH)
    });

    currentProcess = ps;
    let progressCounter = 0;
    let actualPackagePath = null;

    ps.stdout.on('data', (data) => {
      const message = data.toString().trim();
      if (!message) return;

      // Filter out Win32 Content Prep Tool INFO lines
      if (message.startsWith('INFO:') || message.includes('[INFO]')) {
        return;
      }

      const packagePathMatch = message.match(/M365 Office deployment package created (?:with zip compression |without zip compression )?at (.+?)$/i);
      if (packagePathMatch && packagePathMatch[1]) {
        actualPackagePath = packagePathMatch[1].trim();
      }

      const detailsKeywords = ['PatchMyPC', 'Custom App', 'Configuration', 'Channel:', 'Version:', 'contains:', 'Package Contents', 'custom app information'];
      const isDetailsMessage = detailsKeywords.some((keyword) => message.includes(keyword));

      if (message.includes('[PROGRESS]')) {
        const progressMsg = message.replace('[PROGRESS]', '').trim();
        sendToRenderer('status', { message: progressMsg, severity: 'running' });
        if (message.includes('Downloading')) {
          progressCounter = Math.min(progressCounter + 15, 85);
        } else if (message.toLowerCase().includes('complete')) {
          progressCounter = 95;
        }
        sendToRenderer('progress', { percent: progressCounter, speed: progressMsg });
      } else {
        sendToRenderer(isDetailsMessage ? 'details' : 'log', { message, severity: 'info' });
        if (message.toLowerCase().includes('downloading')) {
          progressCounter = Math.min(progressCounter + 10, 90);
          sendToRenderer('progress', { percent: progressCounter, speed: 'Downloading files...' });
        }
      }
    });

    ps.stderr.on('data', (data) => {
      const message = data.toString().trim();
      if (!message) return;
      sendToRenderer('log', { message, severity: 'error' });
    });

    ps.on('close', (code) => {
      currentProcess = null;
      if (code === 0) {
        sendToRenderer('progress', { percent: 100, speed: 'Complete' });
        sendToRenderer('complete', { success: true, message: 'Office package created successfully!', outputPath: actualPackagePath, code });
      } else {
        sendToRenderer('complete', { success: false, message: `Process exited with code ${code}`, code });
      }
    });

    ps.on('error', (err) => {
      sendToRenderer('log', { message: `Process error: ${err.message}`, severity: 'error' });
      sendToRenderer('complete', { success: false, message: `Failed to start PowerShell: ${err.message}`, code: -1 });
    });

    return { success: true, message: 'Process started' };
  });

  ipcMain.handle('stop:deployment', () => {
    if (currentProcess && !currentProcess.killed) {
      currentProcess.kill('SIGTERM');
      setTimeout(() => {
        if (currentProcess && !currentProcess.killed) {
          currentProcess.kill('SIGKILL');
        }
      }, 3000);
      return { success: true, message: 'Stop signal sent' };
    }
    return { success: false, message: 'No process running' };
  });
}

// Initialize IPC handlers
setupIPCHandlers();

