const { SerialPort } = require('serialport');
const { ReadlineParser } = require('@serialport/parser-readline');
const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const fs = require('fs');
const path = require('path');

let mainWindow;
let headerEditorWindow;
let lastScannedWindow = null;

let port;
const BAUD_RATE = 115200; // MUST match Arduino's Serial.begin()

// List of ports to check (edit this!)
const PORTS_TO_CHECK = ['COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6'];

let currentPortIndex = 0;
let isConnected = false;
let csvFilePath = '';
let logMessages = [];
let defaultHeaders = ['Timestamp', 'Data1', 'Data2', 'Data3', 'Data4', 'Data5'];
let currentHeaders = [...defaultHeaders];

// MARK: Load saved headers on startup
function loadSavedHeaders() {
  const configPath = path.join(__dirname, 'headers_config.json');
  try {
    if (fs.existsSync(configPath)) {
      const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
      if (config.headers && Array.isArray(config.headers)) {
        currentHeaders = config.headers;
        log('Loaded saved headers from config');
      }
    }
  } catch (error) {
    log('Using default headers (config load failed)');
    currentHeaders = [...defaultHeaders];
  }
}

function openHeaderEditor() {
  if (headerEditorWindow) {
    headerEditorWindow.focus();
    return;
  }

  headerEditorWindow = new BrowserWindow({
    width: 600,
    height: 500,
    parent: mainWindow,
    modal: true,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
    title: 'CSV Header Editor',
    autoHideMenuBar: true,
  });

  headerEditorWindow.loadFile('header-editor.html');

  headerEditorWindow.on('closed', () => {
    headerEditorWindow = null;
  });
}

// IPC handlers for header management
ipcMain.handle('get-headers', () => {
  return currentHeaders;
});

ipcMain.handle('save-headers', (event, headers) => {
  try {
    // Validate headers
    if (!Array.isArray(headers) || headers.length < 2) {
      return { success: false, error: 'Headers must be an array with at least 2 items (Timestamp + 1 data field)' };
    }

    // Ensure first three headers are always the system headers
    headers[0] = 'Timestamp';

    currentHeaders = headers;

    // Save headers to a config file
    const configPath = path.join(__dirname, 'headers_config.json');
    fs.writeFileSync(configPath, JSON.stringify({ headers: currentHeaders }, null, 2));

    log(`Headers updated: ${currentHeaders.length} columns configured`);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('reset-headers', () => {
  currentHeaders = [...defaultHeaders];
  const configPath = path.join(__dirname, 'headers_config.json');
  try {
    if (fs.existsSync(configPath)) {
      fs.unlinkSync(configPath);
    }
  } catch (error) {
    console.error('Error removing config file:', error);
  }
  log('Headers reset to default');
  return { success: true };
});

ipcMain.handle('open-header-editor', () => {
  openHeaderEditor();
});

// MARK: Open last scanned data window
function openLastScannedWindow() {
  if (lastScannedWindow) {
    lastScannedWindow.focus();
    return;
  }

  lastScannedWindow = new BrowserWindow({
    width: 600,
    height: 200,
    frame: true,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
    autoHideMenuBar: true,
    title: 'Last Scanned Data',
  });

  lastScannedWindow.loadFile('last-scanned.html');

  lastScannedWindow.on('closed', () => {
    lastScannedWindow = null;
  });
}

function createLastScannedWindow() {
  if (lastScannedWindow && !lastScannedWindow.isDestroyed()) {
    lastScannedWindow.focus();
    return;
  }

  lastScannedWindow = new BrowserWindow({
    width: 400,
    height: 200,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
    autoHideMenuBar: true
  });

  lastScannedWindow.loadFile('last-scanned.html');

  lastScannedWindow.on('closed', () => {
    lastScannedWindow = null;
  });
}

ipcMain.handle('open-last-scanned-window', () => {
  createLastScannedWindow();
});

//MARK: IPC handlers for save functionality
ipcMain.handle('save-csv-as', async () => {
  if (!csvFilePath || !fs.existsSync(csvFilePath)) {
    return { success: false, error: 'No CSV file available to save' };
  }

  try {
    const result = await dialog.showSaveDialog(mainWindow, {
      title: 'Save CSV File As',
      defaultPath: path.basename(csvFilePath),
      filters: [
        { name: 'CSV Files', extensions: ['csv'] },
        { name: 'All Files', extensions: ['*'] }
      ]
    });

    if (!result.canceled && result.filePath) {
      // Copy the current CSV file to the selected location
      fs.copyFileSync(csvFilePath, result.filePath);
      return { success: true, filePath: result.filePath };
    }

    return { success: false, error: 'Save cancelled' };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('browse-and-save-csv', async () => {
  try {
    // Get all CSV files in the current directory
    const csvFiles = getAllCSVFiles();

    if (csvFiles.length === 0) {
      return { success: false, error: 'No CSV files found in application directory' };
    }

    // Show dialog to select which CSV to save
    const selectResult = await dialog.showMessageBox(mainWindow, {
      type: 'question',
      title: 'Select CSV File to Save',
      message: 'Choose a CSV file to save:',
      detail: csvFiles.map((file, index) =>
        `${index + 1}. ${file.name} (${formatFileSize(file.size)}, ${file.entries} entries, ${file.date})`
      ).join('\n'),
      buttons: csvFiles.map((file, index) => `${index + 1}. ${file.name}`).concat(['Cancel']),
      defaultId: 0,
      cancelId: csvFiles.length
    });

    if (selectResult.response === csvFiles.length) {
      return { success: false, error: 'Selection cancelled' };
    }

    const selectedFile = csvFiles[selectResult.response];

    // Show save dialog for the selected file
    const saveResult = await dialog.showSaveDialog(mainWindow, {
      title: `Save ${selectedFile.name} As`,
      defaultPath: selectedFile.name,
      filters: [
        { name: 'CSV Files', extensions: ['csv'] },
        { name: 'All Files', extensions: ['*'] }
      ]
    });

    if (!saveResult.canceled && saveResult.filePath) {
      fs.copyFileSync(selectedFile.fullPath, saveResult.filePath);
      return {
        success: true,
        filePath: saveResult.filePath,
        originalFile: selectedFile.name
      };
    }

    return { success: false, error: 'Save cancelled' };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

function getAllCSVFiles() {
  try {
    const files = fs.readdirSync(__dirname);
    const csvFiles = files
      .filter(file => file.startsWith('arduino_logs_') && file.endsWith('.csv'))
      .map(file => {
        const fullPath = path.join(__dirname, file);
        const stats = fs.statSync(fullPath);
        const content = fs.readFileSync(fullPath, 'utf8');
        const lines = content.split('\n').filter(line => line.trim() !== '');

        return {
          name: file,
          fullPath: fullPath,
          size: stats.size,
          entries: Math.max(0, lines.length - 1), // Subtract header
          date: stats.mtime.toLocaleDateString(),
          modified: stats.mtime
        };
      })
      .sort((a, b) => b.modified - a.modified); // Sort by newest first

    return csvFiles;
  } catch (error) {
    console.error('Error reading CSV files:', error);
    return [];
  }
}

function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

ipcMain.handle('get-all-csv-files', () => {
  return getAllCSVFiles();
});

ipcMain.handle('get-csv-info', () => {
  if (!csvFilePath || !fs.existsSync(csvFilePath)) {
    return { exists: false };
  }

  try {
    const stats = fs.statSync(csvFilePath);
    const content = fs.readFileSync(csvFilePath, 'utf8');
    const lines = content.split('\n').filter(line => line.trim() !== '');

    return {
      exists: true,
      fileName: path.basename(csvFilePath),
      fullPath: csvFilePath,
      size: stats.size,
      entries: Math.max(0, lines.length - 1), // Subtract header
      lastModified: stats.mtime
    };
  } catch (error) {
    return { exists: false, error: error.message };
  }
});

// Initialize CSV file
function initializeCSV() {
  const date = new Date().toISOString().split('T')[0]; // YYYY-MM-DD format
  csvFilePath = path.join(__dirname, `arduino_logs_${date}.csv`);

  // Check if file exists, if not create with header
  try {
    if (!fs.existsSync(csvFilePath)) {
      const header = currentHeaders.join(',') + '\n';
      fs.writeFileSync(csvFilePath, header);
      log(`CSV log file created: ${path.basename(csvFilePath)}`);
    } else {
      log(`Using existing CSV log file: ${path.basename(csvFilePath)}`);
    }
  } catch (err) {
    if (err.code === 'EBUSY' || err.code === 'EPERM') {
      log(`Warning: CSV file is locked by another application. Logging will continue when file is available.`);
    } else {
      log(`Error initializing CSV: ${err.message}`);
    }
  }
}

function writeToCSV(message) {
  if (!csvFilePath) return;

  const timestamp = new Date().toISOString();

  // Construct a pipe-delimited row: Timestamp + message split by "|"
  const messageParts = message.split('|').map(p => p.trim());
  const csvFields = [timestamp, ...messageParts];
  const csvLine = csvFields.join(',') + '\n';

  // Retry mechanism for file access
  const maxRetries = 3;
  let retryCount = 0;

  function attemptWrite() {
    fs.appendFile(csvFilePath, csvLine, (err) => {
      if (err) {
        if (err.code === 'EBUSY' || err.code === 'ENOENT' || err.code === 'EPERM') {
          retryCount++;
          if (retryCount < maxRetries) {
            console.log(`CSV write failed (${err.code}), retrying in ${retryCount * 100}ms... (${retryCount}/${maxRetries})`);
            setTimeout(attemptWrite, retryCount * 100);
          } else {
            console.error(`Failed to write to CSV after ${maxRetries} attempts:`, err.message);
          }
        } else {
          console.error('Error writing to CSV:', err.message);
        }
      }
    });
  }

  attemptWrite();
}

function log(message) {
  console.log(message);
  logMessages.push(message);

  if (mainWindow && mainWindow.webContents) {
    mainWindow.webContents.send('log', logMessages);
  }

}
function sendData(data) {
  if (mainWindow && mainWindow.webContents) {
    mainWindow.webContents.send('data', data);
  }

  if (lastScannedWindow && lastScannedWindow.webContents) {
    lastScannedWindow.webContents.send('last-scanned-data', data);
  }

  const currentPort = currentPortIndex < PORTS_TO_CHECK.length ? PORTS_TO_CHECK[currentPortIndex] : '';
  writeToCSV(data);
}

function tryNextPort() {
  if (isConnected) return; // Don't try new ports if already connected

  if (currentPortIndex >= PORTS_TO_CHECK.length) {
    currentPortIndex = 0; // Restart from first port
    log('No Arduino found. Retrying in 3 seconds...');
    setTimeout(tryNextPort, 3000);
    return;
  }

  const portPath = PORTS_TO_CHECK[currentPortIndex];
  log(`Checking ${portPath}...`);

  // Close existing port if open
  if (port && port.isOpen) {
    port.close();
  }

  // Try to open the port
  port = new SerialPort({
    path: portPath,
    baudRate: BAUD_RATE,
    autoOpen: false
  });

  // Set a timeout to detect dead ports
  const timeout = setTimeout(() => {
    if (port && port.isOpen) {
      port.close();
    }
    currentPortIndex++;
    setTimeout(tryNextPort, 1000); // Try next port after 1 sec
  }, 3000); // Increased timeout

  port.open((err) => {
    if (err) {
      clearTimeout(timeout);
      log(`Failed to open ${portPath}: ${err.message}`);
      currentPortIndex++;
      setTimeout(tryNextPort, 1000);
      return;
    }

    log(`Port ${portPath} opened, waiting for data...`);

    // Create parser
    const parser = port.pipe(new ReadlineParser({ delimiter: '\n' }));

    // If data arrives, we found the Arduino!
    parser.once('data', (data) => {
      clearTimeout(timeout);
      isConnected = true;
      log(`✅ ARDUINO FOUND ON ${portPath}!`);

      // Log all future data
      parser.on('data', (data) => {
        const trimmed = data.trim();
        if (/^\d/.test(trimmed)) { // ✅ Starts with a digit
          sendData(trimmed);
        } else {
          log(`Ignored non-numeric data: "${trimmed}"`);
        }
      });
    });

    // Handle port errors
    port.on('error', (err) => {
      clearTimeout(timeout);
      log(`Port error on ${portPath}: ${err.message}`);
      if (port && port.isOpen) {
        port.close();
      }
      currentPortIndex++;
      setTimeout(tryNextPort, 1000);
    });

    // Handle port close
    port.on('close', () => {
      if (isConnected) {
        log(`Connection to ${portPath} lost. Reconnecting...`);
        isConnected = false;
        currentPortIndex = 0;
        setTimeout(tryNextPort, 2000);
      }
    });
  });
}

function createWindow() {
  // Initialize CSV logging
  initializeCSV();

  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false // Required for older Electron versions with nodeIntegration
    },
    autoHideMenuBar: true,
  });

  mainWindow.loadFile('index.html');

  // Start checking ports once window is ready
  mainWindow.webContents.once('did-finish-load', () => {
    tryNextPort();
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  loadSavedHeaders();
  createWindow();
  openLastScannedWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
      openLastScannedWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('before-quit', () => {
  if (port && port.isOpen) {
    port.close();
  }
});

