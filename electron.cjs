const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const path = require('path')
const fs = require('fs')

const isDev = process.env.NODE_ENV === 'development' || !app.isPackaged
const DATA_PATH = path.join(app.getPath('userData'), 'attendx-data.json')

function safeReadJson(filePath) {
  try {
    if (!fs.existsSync(filePath)) return null
    const raw = fs.readFileSync(filePath, 'utf-8')
    if (!raw.trim()) return null
    return JSON.parse(raw)
  } catch (e) {
    console.error('Load error:', e)
    try {
      const backupPath = `${filePath}.broken-${Date.now()}`
      fs.renameSync(filePath, backupPath)
      console.error('Broken data file moved to:', backupPath)
    } catch (backupError) {
      console.error('Could not backup broken data file:', backupError)
    }
    return null
  }
}

function createWindow() {
  const preloadPath = path.join(__dirname, 'preload.cjs')

  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: preloadPath,
      devTools: true,
    },
    title: 'AttendX',
    autoHideMenuBar: true,
    show: false,
  })

  win.once('ready-to-show', () => {
    win.show()
  })

  win.webContents.on('did-fail-load', (_event, errorCode, errorDescription, validatedURL) => {
    const message = `AttendX failed to load.\n\nURL: ${validatedURL}\nError ${errorCode}: ${errorDescription}`
    console.error(message)
    dialog.showErrorBox('AttendX load error', message)
  })

  win.webContents.on('render-process-gone', (_event, details) => {
    const message = `AttendX renderer crashed.\n\nReason: ${details.reason}\nExit code: ${details.exitCode}`
    console.error(message)
    dialog.showErrorBox('AttendX renderer crashed', message)
  })

  win.webContents.on('console-message', (_event, level, message, line, sourceId) => {
    console.log(`[renderer:${level}] ${message} (${sourceId}:${line})`)
  })

  if (isDev) {
    win.loadURL('http://localhost:5173')
    win.webContents.openDevTools({ mode: 'detach' })
  } else {
    const indexPath = path.join(app.getAppPath(), 'dist', 'index.html')
    console.log('Loading production index:', indexPath)

    if (!fs.existsSync(indexPath)) {
      const fallbackPath = path.join(__dirname, 'dist', 'index.html')
      const message = `dist/index.html was not found.\n\nTried:\n${indexPath}\n\nFallback:\n${fallbackPath}\n\nRun npm run build before packaging and confirm dist/index.html exists.`
      console.error(message)
      dialog.showErrorBox('AttendX missing build files', message)
      win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(`<h2>AttendX build files missing</h2><pre>${message}</pre>`))
      return
    }

    win.loadFile(indexPath)
  }
}

ipcMain.handle('load-data', () => safeReadJson(DATA_PATH))

ipcMain.handle('save-data', (_, data) => {
  try {
    fs.mkdirSync(path.dirname(DATA_PATH), { recursive: true })
    fs.writeFileSync(DATA_PATH, JSON.stringify(data, null, 2), 'utf-8')
    return true
  } catch (e) {
    console.error('Save error:', e)
    return false
  }
})

ipcMain.handle('reset-data', () => {
  try {
    if (fs.existsSync(DATA_PATH)) fs.unlinkSync(DATA_PATH)
    return true
  } catch (e) {
    console.error('Reset data error:', e)
    return false
  }
})

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})
