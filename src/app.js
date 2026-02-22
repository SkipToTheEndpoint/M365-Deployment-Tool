let isExecutionComplete = false;
let reconnectAttempts = 0;
let actualOutputPath = null;

document.addEventListener('DOMContentLoaded', async () => {
  setupIpcListeners();
  setupEventListeners();
  loadXMLFiles();
  
  const noZipContainer = document.getElementById('noZipContainer');
  noZipContainer.style.display = 'none';
  
  // Initialize button state
  updateStartButtonState();
  
  // Load and set default paths
  if (window.electronAPI) {
    try {
      const defaultPaths = await window.electronAPI.getDefaultPaths();
      document.getElementById('outputPath').value = defaultPaths.outputPath;
    } catch (error) {
      console.error('[Paths] Error loading default paths:', error);
    }
  }
  
  // Initialize checkbox states after settings are loaded
  initializeCheckboxStates();
});

// IPC listeners for main-process events
function setupIpcListeners() {
  if (!window.electronAPI) {
    console.warn('[IPC] electronAPI unavailable; running in non-Electron context');
    return;
  }

  // Log messages
  window.electronAPI.onLog((data) => {
    addLogEntry(data.message, data.severity || 'info', false);
  });

  // Detail messages
  window.electronAPI.onDetails((data) => {
    addLogEntry(data.message, data.severity || 'info', true);
  });

  // Status updates
  window.electronAPI.onStatus((data) => {
    updateStatus(data.message, data.severity || 'info');
  });

  // Progress updates
  window.electronAPI.onProgress((data) => {
    updateProgress(data.percent ?? 0, data.speed || '');
  });

  // Completion events
  window.electronAPI.onComplete((data) => {
    isExecutionComplete = true;
    completeExecution(data.success, data.message, data.outputPath);
  });
}

function updateStartButtonState() {
  const runBtn = document.getElementById('runBtn');
  const configXmlPath = document.getElementById('configXmlPath').dataset.path || '';
  const hasXml = configXmlPath && configXmlPath !== 'undefined' && configXmlPath.trim();
  
  if (!hasXml) {
    runBtn.disabled = true;
    runBtn.title = 'Please select a Configuration XML file first';
  } else {
    runBtn.disabled = false;
    runBtn.title = '';
  }
}

function initializeCheckboxStates() {
  const createIntuneWinCheckbox = document.getElementById('createIntuneWin');
  const pmpcCustomAppCheckbox = document.getElementById('pmpcCustomApp');
  const onlineModeCheckbox = document.getElementById('onlineMode');
  const skipAPICheckCheckbox = document.getElementById('skipAPICheck');
  const noZipContainer = document.getElementById('noZipContainer');
  
  // If createIntuneWin is checked, disable pmpcCustomApp
  if (createIntuneWinCheckbox.checked) {
    pmpcCustomAppCheckbox.checked = false;
    pmpcCustomAppCheckbox.disabled = true;
    const pmpcCustomAppLabel = pmpcCustomAppCheckbox.closest('.toggle-group').querySelector('.toggle-help');
    pmpcCustomAppLabel.textContent = 'Not available when creating Intune .intunewin package (mutually exclusive deployment methods)';
    pmpcCustomAppLabel.classList.add('toggle-help-unavailable');
    noZipContainer.style.display = 'none';
  }
  
  // If pmpcCustomApp is checked, disable createIntuneWin
  if (pmpcCustomAppCheckbox.checked && !pmpcCustomAppCheckbox.disabled) {
    createIntuneWinCheckbox.checked = false;
    createIntuneWinCheckbox.disabled = true;
    const createIntuneWinLabel = createIntuneWinCheckbox.closest('.toggle-group').querySelector('.toggle-help');
    createIntuneWinLabel.textContent = 'Not available when creating Patch My PC Custom App';
    createIntuneWinLabel.classList.add('toggle-help-unavailable');
    
    if (!onlineModeCheckbox.checked) {
      noZipContainer.style.display = 'flex';
    }
  }
  
  // If onlineMode is checked, disable skipAPICheck
  if (onlineModeCheckbox.checked) {
    skipAPICheckCheckbox.disabled = true;
    skipAPICheckCheckbox.checked = false;
    const skipAPICheckLabel = skipAPICheckCheckbox.closest('.toggle-group').querySelector('.toggle-help');
    skipAPICheckLabel.textContent = 'Not available when the option to "Create package without downloading Office data files" is checked';
    skipAPICheckLabel.classList.add('toggle-help-unavailable');
  }
  
  // If skipAPICheck is checked, disable onlineMode
  if (skipAPICheckCheckbox.checked) {
    onlineModeCheckbox.disabled = true;
    onlineModeCheckbox.checked = false;
    const onlineModeLabel = onlineModeCheckbox.closest('.toggle-group').querySelector('.toggle-help');
    onlineModeLabel.textContent = 'Not available with Skip API Check';
    onlineModeLabel.classList.add('toggle-help-unavailable');
  }
}

// Setup event listeners
function setupEventListeners() {
  // File picker for XML
  document.getElementById('browseXml').addEventListener('click', async () => {
    console.log('[Browse XML] electronAPI available:', !!window.electronAPI);
    if (window.electronAPI && window.electronAPI.openFileDialog) {
      try {
        const filePath = await window.electronAPI.openFileDialog();
        console.log('[Browse XML] Selected:', filePath);
        if (filePath) {
          document.getElementById('configXmlPath').value = filePath;
          document.getElementById('configXmlPath').dataset.path = filePath;
          updateStartButtonState(); // Update button state when XML is selected
        }
      } catch (error) {
        console.error('[Browse XML] Error:', error);
      }
    } else {
      console.error('[Browse XML] electronAPI not available');
    }
  });

  // Path selector for output folder
  document.getElementById('browseOutput').addEventListener('click', async () => {
    if (window.electronAPI && window.electronAPI.openFolderDialog) {
      try {
        const folderPath = await window.electronAPI.openFolderDialog();
        if (folderPath) {
          document.getElementById('outputPath').value = folderPath;
        }
      } catch (error) {
        console.error('[Browse Output] Error:', error);
      }
    }
  });


  document.getElementById('pmpcCustomApp').addEventListener('change', (e) => {
    const noZipContainer = document.getElementById('noZipContainer');
    const noZipCheckbox = document.getElementById('noZip');
    const onlineModeCheckbox = document.getElementById('onlineMode');
    const createIntuneWinCheckbox = document.getElementById('createIntuneWin');
    const createIntuneWinLabel = createIntuneWinCheckbox.closest('.toggle-group').querySelector('.toggle-help');
    
    if (e.target.checked) {
      noZipCheckbox.checked = false;
      createIntuneWinCheckbox.checked = false;
      createIntuneWinCheckbox.disabled = true;
      createIntuneWinLabel.textContent = 'Not available when creating Patch My PC Custom App';
      createIntuneWinLabel.classList.add('toggle-help-unavailable');
      
      if (!onlineModeCheckbox.checked) {
        noZipContainer.style.display = 'flex';
      }
    } else {
      noZipContainer.style.display = 'none';
      noZipCheckbox.checked = false;
      createIntuneWinCheckbox.disabled = false;
      createIntuneWinLabel.textContent = 'Downloads the Win32 Content Prep Tool and builds a .intunewin package';
      createIntuneWinLabel.classList.remove('toggle-help-unavailable');
    }
  });

  document.getElementById('onlineMode').addEventListener('change', (e) => {
    const skipAPICheckCheckbox = document.getElementById('skipAPICheck');
    const skipAPICheckLabel = skipAPICheckCheckbox.closest('.toggle-group').querySelector('.toggle-help');
    const noZipContainer = document.getElementById('noZipContainer');
    const noZipCheckbox = document.getElementById('noZip');
    const pmpcCustomAppCheckbox = document.getElementById('pmpcCustomApp');
    const createIntuneWinCheckbox = document.getElementById('createIntuneWin');
    
    if (e.target.checked) {
      skipAPICheckCheckbox.disabled = true;
      skipAPICheckCheckbox.checked = false;
      skipAPICheckLabel.textContent = 'Not available when the option to "Create package without downloading Office data files" is checked';
      skipAPICheckLabel.classList.add('toggle-help-unavailable');

      const createIntuneWinLabel = createIntuneWinCheckbox.closest('.toggle-group').querySelector('.toggle-help');
      // Only update createIntuneWin if pmpcCustomApp is not checked
      if (!pmpcCustomAppCheckbox.checked) {
        createIntuneWinLabel.textContent = 'Downloads the Win32 Content Prep Tool and builds a .intunewin package. ';
        createIntuneWinLabel.classList.remove('toggle-help-unavailable');
      } else {
        // Keep createIntuneWin unavailable if pmpcCustomApp is checked
        createIntuneWinLabel.textContent = 'Not available when PMPC Custom App is selected';
        createIntuneWinLabel.classList.add('toggle-help-unavailable');
      }
      
      noZipContainer.style.display = 'none';
      noZipCheckbox.checked = false;
    } else {
      skipAPICheckCheckbox.disabled = false;
      skipAPICheckLabel.textContent = 'Skip Office version validation (only when downloading Office files)';
      skipAPICheckLabel.classList.remove('toggle-help-unavailable');
      skipAPICheckLabel.classList.remove('toggle-help-unavailable');

      createIntuneWinCheckbox.closest('.toggle-group').querySelector('.toggle-help').textContent = 'Downloads the Win32 Content Prep Tool and builds a .intunewin package.';
      
      if (pmpcCustomAppCheckbox.checked) {
        noZipContainer.style.display = 'flex';
        noZipCheckbox.checked = false;
      }
    }
  });

  document.getElementById('createIntuneWin').addEventListener('change', (e) => {
    const pmpcCustomAppCheckbox = document.getElementById('pmpcCustomApp');
    const pmpcCustomAppLabel = pmpcCustomAppCheckbox.closest('.toggle-group').querySelector('.toggle-help');
    
    if (e.target.checked) {
      pmpcCustomAppCheckbox.checked = false;
      pmpcCustomAppCheckbox.disabled = true;
      pmpcCustomAppLabel.textContent = 'Not available when creating Intune .intunewin package (mutually exclusive deployment methods)';
      pmpcCustomAppLabel.classList.add('toggle-help-unavailable');
      
      const noZipContainer = document.getElementById('noZipContainer');
      noZipContainer.style.display = 'none';
      document.getElementById('noZip').checked = false;
    } else {
      pmpcCustomAppCheckbox.disabled = false;
      pmpcCustomAppLabel.textContent = 'Creates a Patch My PC Custom App package';
      pmpcCustomAppLabel.classList.remove('toggle-help-unavailable');
    }
  });

  document.getElementById('skipAPICheck').addEventListener('change', (e) => {
    const onlineModeCheckbox = document.getElementById('onlineMode');
    const onlineModeLabel = onlineModeCheckbox.closest('.toggle-group').querySelector('.toggle-help');
    
    if (e.target.checked) {
      onlineModeCheckbox.disabled = true;
      onlineModeCheckbox.checked = false;
      onlineModeLabel.textContent = 'Not available with Skip API Check';
      onlineModeLabel.classList.add('toggle-help-unavailable');
    } else {
      onlineModeCheckbox.disabled = false;
      onlineModeLabel.textContent = 'Creates a lightweight package with no Office data files (~200MB setup.exe + configs). If this option is unchecked, the Office data files are downloaded into the package (~3-4GB)';
      onlineModeLabel.classList.remove('toggle-help-unavailable');
    }
  });

  // Run button
  document.getElementById('runBtn').addEventListener('click', executeDeployment);

  // Stop button
  document.getElementById('stopBtn').addEventListener('click', stopExecution);

  // Reset button
  document.getElementById('resetBtn').addEventListener('click', resetForm);

  // Clear log
  document.getElementById('clearLogBtn').addEventListener('click', () => {
    document.getElementById('logOutput').innerHTML = '';
    addLogEntry('Log cleared', 'info');
  });

  // Copy log
  document.getElementById('copyLogBtn').addEventListener('click', copyLog);

  // Result actions
  document.getElementById('openOutputBtn').addEventListener('click', openOutputFolder);
}

async function loadXMLFiles() {
  try {
    if (window.electronAPI && window.electronAPI.listXMLFiles) {
      const files = await window.electronAPI.listXMLFiles();
      if (files && files.length > 0) {
        addLogEntry(`Found ${files.length} XML file(s) in current directory`, 'info');
      }
    }
  } catch (error) {
    console.error('Error loading XML files:', error);
  }
}

// Execute deployment
async function executeDeployment() {
  // Reset execution state
  isExecutionComplete = false;
  reconnectAttempts = 0;
  const configXmlPath = document.getElementById('configXmlPath').dataset.path || '';
  const configXml = configXmlPath && configXmlPath !== 'undefined' ? configXmlPath : '';
  const outputPath = document.getElementById('outputPath').value;
  const logName = document.getElementById('logName').value;
  const onlineMode = document.getElementById('onlineMode').checked;
  const skipAPICheck = document.getElementById('skipAPICheck').checked;
  const pmpcCustomApp = document.getElementById('pmpcCustomApp').checked;
  const createIntuneWin = document.getElementById('createIntuneWin').checked;
  
  let noZip = document.getElementById('noZip').checked;
  // If not using pmpcCustomApp, force noZip true
  if (!pmpcCustomApp) {
    noZip = true;
  }
  
  console.log('[Frontend] Sending to server:');
  console.log('  noZip:', noZip);
  console.log('  onlineMode:', onlineMode);
  console.log('  pmpcCustomApp:', pmpcCustomApp);
  console.log('  createIntuneWin:', createIntuneWin);
  
  const apiRetryDelaySeconds = parseInt(document.getElementById('apiRetryDelay').value);
  const apiMaxExtendedAttempts = parseInt(document.getElementById('apiMaxAttempts').value);
  const setupUrl = document.getElementById('setupUrl').value;
  const officeIconUrl = document.getElementById('officeIconUrl').value;
  const officeVersionUrl = document.getElementById('officeVersionUrl').value;

  // Validation
  if (!configXml) {
    addLogEntry('Please select a configuration XML file', 'error');
    return;
  }

  const runBtn = document.getElementById('runBtn');
  const stopBtn = document.getElementById('stopBtn');
  const originalText = runBtn.innerHTML;
  runBtn.disabled = true;
  runBtn.innerHTML = '<span class="btn-icon-text"></span> Running...';
  stopBtn.style.display = 'inline-block';

  // Update status
  updateStatus('Initializing package...', 'running');
  document.querySelector('.log-output').innerHTML = '';
  document.getElementById('progressContainer').classList.remove('hidden');
  document.getElementById('resultPanel').classList.add('hidden');

  try {
    if (!window.electronAPI || !window.electronAPI.executeDeployment) {
      throw new Error('Electron IPC not available');
    }

    const result = await window.electronAPI.executeDeployment({
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
      officeIconUrl,
      officeVersionUrl
    });

    if (!result?.success) {
      throw new Error(result?.message || 'Failed to start deployment');
    }
    addLogEntry(result.message || 'Deployment started', 'info');
  } catch (error) {
    addLogEntry(`Error starting deployment: ${error.message}`, 'error');
    updateStatus('Error', 'error');
    runBtn.disabled = false;
    runBtn.innerHTML = originalText;
  }
}

function addLogEntry(message, severity = 'info', isDetails = false) {
  const preferredId = isDetails ? 'detailsOutput' : 'logOutput';
  let logOutput = document.getElementById(preferredId);
  if (!logOutput) {
    logOutput = document.getElementById('logOutput');
  }
  const entry = document.createElement('div');
  entry.className = `log-entry log-${severity}`;

  const msg = document.createElement('span');
  msg.className = 'log-message';
  msg.textContent = message;

  entry.appendChild(msg);
  logOutput.appendChild(entry);

  logOutput.scrollTop = logOutput.scrollHeight;
}

// Update status
function updateStatus(message, type = 'info') {
  const indicator = document.getElementById('statusIndicator');
  const statusText = document.querySelector('.status-text');

  indicator.className = `status-indicator status-${type}`;
  statusText.textContent = message;

  const typeMap = {
    'info': 'info',
    'running': 'running',
    'success': 'success',
    'error': 'error'
  };

  indicator.className = `status-indicator status-${typeMap[type] || type}`;
}

// Update progress
function updateProgress(percent, speed = '') {
  const progressFill = document.getElementById('progressFill');
  progressFill.style.width = `${Math.min(percent, 100)}%`;

  document.getElementById('progressPercent').textContent = `${Math.round(percent)}%`;
  if (speed) {
    document.getElementById('progressTime').textContent = speed;
  }
}

// Complete execution
function completeExecution(success, message, outputPath) {
  const runBtn = document.getElementById('runBtn');
  const stopBtn = document.getElementById('stopBtn');
  runBtn.disabled = false;
  runBtn.innerHTML = '<span class="btn-icon-text"></span> Start Packaging';
  stopBtn.style.display = 'none';

  document.getElementById('progressContainer').classList.add('hidden');

  if (success) {
    actualOutputPath = outputPath;
    
    updateStatus('Package created successfully', 'success');
    addLogEntry('M365 Apps package created successfully!', 'success');

    // Show result panel
    const resultPanel = document.getElementById('resultPanel');
    const resultContent = document.getElementById('resultContent');
    resultContent.innerHTML = `
      <div style="margin-bottom: 8px;">
        <strong>Output Location:</strong><br>
        <code style="background: var(--bg-primary); padding: 6px 8px; border-radius: 3px; display: block; margin-top: 4px; font-size: 11px; white-space: pre-wrap; word-break: break-all;">${outputPath}</code>
      </div>
    `;
    resultPanel.classList.remove('hidden');
  } else {
    updateStatus('Deployment failed', 'error');
    addLogEntry('âœ• ' + message, 'error');
  }
}

// Reset form
function resetForm() {
  // Reset execution state
  isExecutionComplete = false;
  reconnectAttempts = 0;
  actualOutputPath = null; // Reset stored path
  
  document.getElementById('configXmlPath').value = '';
  document.getElementById('configXmlPath').dataset.path = '';
  
  if (window.electronAPI) {
    window.electronAPI.getDefaultPaths().then((defaultPaths) => {
      document.getElementById('outputPath').value = defaultPaths.outputPath;
    }).catch(() => {
      document.getElementById('outputPath').value = '';
    });
  } else {
    document.getElementById('outputPath').value = '';
  }
  
  document.getElementById('logName').value = 'Invoke-M365AppsHelper.log';
  document.getElementById('noZip').checked = false;
  document.getElementById('onlineMode').checked = false;
  document.getElementById('skipAPICheck').checked = false;
  document.getElementById('apiRetryDelay').value = '3';
  document.getElementById('apiMaxAttempts').value = '10';
  document.getElementById('setupUrl').value = 'https://officecdn.microsoft.com/pr/wsus/setup.exe';
  document.getElementById('officeIconUrl').value = 'https://www.svgrepo.com/show/452062/microsoft.svg';
  document.getElementById('officeVersionUrl').value = 'https://clients.config.office.net/releases/v1.0/OfficeReleases';

  document.getElementById('logOutput').innerHTML = '';
  addLogEntry('Form reset. Ready for new deployment.', 'info');

  document.getElementById('progressContainer').classList.add('hidden');
  document.getElementById('resultPanel').classList.add('hidden');

  updateStatus('Ready', 'idle');
}

// Stop execution
async function stopExecution() {
  isExecutionComplete = true;
  try {
    if (!window.electronAPI || !window.electronAPI.stopDeployment) {
      throw new Error('Electron IPC not available');
    }
    const result = await window.electronAPI.stopDeployment();
    if (result?.success) {
      addLogEntry('Stop signal sent to process', 'warning');
      const runBtn = document.getElementById('runBtn');
      runBtn.disabled = false;
      runBtn.innerHTML = '<span class="btn-icon-text">â–¶</span> Start Deployment';
    } else {
      addLogEntry(result?.message || 'No process running', 'warning');
    }
  } catch (error) {
    console.error('[Stop] Error:', error);
    addLogEntry('Error stopping process: ' + error.message, 'error');
  }
}

function copyLog() {
  const logOutput = document.getElementById('logOutput');
  const logText = Array.from(logOutput.querySelectorAll('.log-message'))
    .map(el => el.textContent)
    .join('\n');

  navigator.clipboard.writeText(logText).then(() => {
    const btn = document.getElementById('copyLogBtn');
    const originalTitle = btn.title;
    btn.title = 'Copied!';
    btn.style.color = 'var(--success)';
    setTimeout(() => {
      btn.title = originalTitle;
      btn.style.color = '';
    }, 2000);
  }).catch(err => {
    addLogEntry('Failed to copy log', 'error');
  });
}

async function openOutputFolder() {
  const folderPath = actualOutputPath || document.getElementById('outputPath').value;
  
  if (window.electronAPI && window.electronAPI.openFolder) {
    try {
      const result = await window.electronAPI.openFolder(folderPath);
      if (result.success) {
        addLogEntry(`Opened folder: ${folderPath}`, 'success');
      } else {
        addLogEntry(`Failed to open folder: ${result.error}`, 'error');
      }
    } catch (error) {
      addLogEntry(`Error opening folder: ${error.message}`, 'error');
    }
  } else {
    // Fallback for non-Electron environments
    addLogEntry(`Output location: ${folderPath}`, 'info');
  }
}

