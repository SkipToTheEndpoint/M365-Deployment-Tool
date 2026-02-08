const express = require('express');
const cors = require('cors');
const path = require('path');
const { spawn } = require('child_process');
const WebSocket = require('ws');
const http = require('http');
const fs = require('fs');

const app = express();
const server = http.createServer(app);
const wss = new WebSocket.Server({ server, path: '/ws' });

const PORT = process.env.PORT || 3000;
const SCRIPT_PATH = path.join(__dirname, 'Invoke-M365AppsHelper.ps1');

// Track current process
let currentProcess = null;

console.log(`[Server] Script path: ${SCRIPT_PATH}`);
console.log(`[Server] Script exists: ${fs.existsSync(SCRIPT_PATH)}`);

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'src')));

// Store active connections
const clients = new Set();

// WebSocket connection handler
wss.on('connection', (ws) => {
  console.log('[WebSocket] Client connected');
  clients.add(ws);

  ws.send(JSON.stringify({
    type: 'status',
    message: 'Ready',
    severity: 'idle'
  }));

  ws.on('close', () => {
    console.log('[WebSocket] Client disconnected');
    clients.delete(ws);
  });

  ws.on('error', (error) => {
    console.error('[WebSocket] Error:', error.message);
  });
});

// Broadcast to all connected clients
function broadcast(message) {
  clients.forEach((client) => {
    if (client.readyState === WebSocket.OPEN) {
      try {
        client.send(JSON.stringify(message));
      } catch (error) {
        console.error('[Broadcast] Error:', error.message);
      }
    }
  });
}

// API: Execute PowerShell script
app.post('/api/execute', (req, res) => {
  const {
    configXml,
    logName = path.join(process.env.APPDATA || '', 'M365AppsHelper', 'Output', 'Logs', 'Invoke-M365AppsHelper.log'),
    noZip = false,
    onlineMode = false,
    skipAPICheck = false,
    pmpcCustomApp = false,
    createIntuneWin = false,
    apiRetryDelaySeconds = 3,
    apiMaxExtendedAttempts = 10,
    setupUrl = 'https://officecdn.microsoft.com/pr/wsus/setup.exe',
    officeVersionUrl = 'https://clients.config.office.net/releases/v1.0/OfficeReleases',
    officeIconUrl = 'https://patchmypc.com/scupcatalog/downloads/icons/Microsoft.png',
  } = req.body;

  console.log('[Execute] Received from frontend:');
  console.log('  noZip:', noZip, '(type:', typeof noZip, ')');
  console.log('  onlineMode:', onlineMode, '(type:', typeof onlineMode, ')');
  console.log('  pmpcCustomApp:', pmpcCustomApp, '(type:', typeof pmpcCustomApp, ')');
  console.log('  createIntuneWin:', createIntuneWin, '(type:', typeof createIntuneWin, ')');

  // NoZip should ONLY be passed if PMPCCustomApp is selected AND noZip checkbox is checked AND OnlineMode is NOT active
  const effectiveNoZip = Boolean(pmpcCustomApp && noZip && !onlineMode);
  // CreateIntuneWin can work with OnlineMode, so don't add !onlineMode condition
  const effectiveCreateIntuneWin = Boolean(createIntuneWin);

  console.log('[Execute] After processing:');
  console.log('  effectiveNoZip:', effectiveNoZip);
  console.log('  effectiveCreateIntuneWin:', effectiveCreateIntuneWin);
  console.log('  Starting deployment with config:', configXml);
  console.log('[Execute] PMPCCustomApp:', pmpcCustomApp, '| OnlineMode:', onlineMode, '| EffectiveNoZip:', effectiveNoZip);

  broadcast({
    type: 'log',
    message: 'Starting Office deployment package creation...',
    severity: 'info',
    timestamp: new Date().toISOString()
  });

  // Build PowerShell command - use absolute paths to avoid issues
  const workingDir = path.dirname(SCRIPT_PATH);

  // Helper to resolve paths only when they are not already absolute
  const resolvePathIfNeeded = (p) => {
    if (!p || p === 'undefined' || !p.trim()) return null;
    return path.isAbsolute(p) ? p : path.resolve(workingDir, p);
  };

  let args = [
    '-NoProfile',
    '-ExecutionPolicy', 'Bypass',
    '-File', SCRIPT_PATH
  ];

  const cfgPath = resolvePathIfNeeded(configXml);
  const logPath = resolvePathIfNeeded(logName);

  if (cfgPath) {
    args.push('-ConfigXML');
    args.push(cfgPath);
  }
  if (logPath) {
    args.push('-LogName');
    args.push(logPath);
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

  console.log('[Execute] Full command:');
  console.log('  Executable:', 'pwsh.exe');
  console.log('  Arguments:', JSON.stringify(args, null, 2));

  // Spawn PowerShell process (use pwsh.exe for PowerShell 7+, fallback to powershell.exe)
  const psExecutable = 'pwsh.exe'; // PowerShell 7+ required for null-coalescing operators
  const ps = spawn(psExecutable, args, {
    stdio: ['pipe', 'pipe', 'pipe'],
    shell: false,
    cwd: path.dirname(SCRIPT_PATH)
  });

  // Track current process
  currentProcess = ps;

  let output = '';
  let errors = '';
  let progressCounter = 0;
  let actualPackagePath = null; // Store the actual package folder path

  ps.stdout.on('data', (data) => {
    const message = data.toString().trim();
    if (message) {
      output += message + '\n';
      console.log('[PS Output]', message);
      
      // Capture the actual package path from success messages
      const packagePathMatch = message.match(/M365 Office deployment package created (?:with zip compression |without zip compression )?at (.+?)$/i);
      if (packagePathMatch && packagePathMatch[1]) {
        actualPackagePath = packagePathMatch[1].trim();
        console.log('[Package Path] Captured:', actualPackagePath);
      }
      
      // Determine if this is a details/instruction message (for PatchMyPC custom app info, configuration details, etc.)
      const detailsKeywords = ['PatchMyPC', 'Custom App', 'Configuration', 'Channel:', 'Version:', 'contains:', 'Package Contents', 'custom app information'];
      const isDetailsMessage = detailsKeywords.some(keyword => message.includes(keyword));
      
      // Check for progress markers
      if (message.includes('[PROGRESS]')) {
        // Extract progress message
        const progressMsg = message.replace('[PROGRESS]', '').trim();
        
        // Send status update
        broadcast({
          type: 'status',
          message: progressMsg,
          severity: 'running'
        });
        
        // Update progress bar
        if (message.includes('Downloading')) {
          progressCounter = Math.min(progressCounter + 15, 85);
        } else if (message.includes('complete') || message.includes('Complete')) {
          progressCounter = 95;
        }
        
        broadcast({
          type: 'progress',
          percent: progressCounter,
          speed: progressMsg
        });
        
        console.log('[Progress]', progressMsg, `${progressCounter}%`);
      } else {
        // Send message to appropriate log section
        broadcast({
          type: isDetailsMessage ? 'details' : 'log',
          message,
          severity: 'info',
          timestamp: new Date().toISOString()
        });
        
        // Also detect regular downloading messages for progress
        if (message.toLowerCase().includes('downloading')) {
          progressCounter += 10;
          if (progressCounter <= 90) {
            broadcast({
              type: 'progress',
              percent: progressCounter,
              speed: 'Downloading files...'
            });
          }
        }
      }
    }
  });

  ps.stderr.on('data', (data) => {
    const message = data.toString().trim();
    if (message) {
      errors += message + '\n';
      console.log('[PS Error]', message);
      broadcast({
        type: 'log',
        message,
        severity: 'error',
        timestamp: new Date().toISOString()
      });
    }
  });

  ps.on('close', (code) => {
    console.log('[Execute] Process closed with code:', code);
    if (code === 0) {
      // Send final progress
      broadcast({
        type: 'progress',
        percent: 100,
        speed: 'Complete'
      });
      broadcast({
        type: 'complete',
        success: true,
        message: 'Office package created successfully!',
        outputPath: actualPackagePath, // Captured path from script logs if available
        code
      });
    } else {
      broadcast({
        type: 'complete',
        success: false,
        message: `Process exited with code ${code}`,
        code,
        errors
      });
    }
  });

  ps.on('error', (err) => {
    console.error('[Execute] Process error:', err.message);
    broadcast({
      type: 'log',
      message: `Process error: ${err.message}`,
      severity: 'error',
      timestamp: new Date().toISOString()
    });
    broadcast({
      type: 'complete',
      success: false,
      message: `Failed to start PowerShell: ${err.message}`,
      code: -1,
      errors: err.message
    });
  });

  res.json({ success: true, message: 'Process started' });
});

// API: Stop running process
app.post('/api/stop', (req, res) => {
  if (currentProcess && !currentProcess.killed) {
    console.log('[Stop] Terminating process PID:', currentProcess.pid);
    broadcast({
      type: 'log',
      message: 'Stop signal received, terminating process...',
      severity: 'warning',
      timestamp: new Date().toISOString()
    });
    
    currentProcess.kill('SIGTERM');
    
    // Force kill after 3 seconds if not already dead
    setTimeout(() => {
      if (currentProcess && !currentProcess.killed) {
        console.log('[Stop] Force killing process');
        currentProcess.kill('SIGKILL');
      }
    }, 3000);
    
    res.json({ success: true, message: 'Stop signal sent' });
  } else {
    res.status(400).json({ success: false, message: 'No process running' });
  }
});

// API: Get available XML files
app.get('/api/xml-files', (req, res) => {
  try {
    const currentDir = process.cwd();
    const files = fs.readdirSync(currentDir)
      .filter(f => f.endsWith('.xml'))
      .map(f => ({ name: f, path: path.join(currentDir, f) }));
    res.json(files);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// Serve index.html for all other routes (SPA)
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'src', 'index.html'));
});

server.listen(PORT, () => {
  console.log(`\n✓ M365 Office UI running on http://localhost:${PORT}`);
  console.log('✓ WebSocket server ready for real-time updates\n');
});

// Graceful shutdown
process.on('SIGINT', () => {
  console.log('\nShutting down...');
  server.close(() => {
    process.exit(0);
  });
});

