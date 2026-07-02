const { app, BrowserWindow, ipcMain, nativeTheme } = require('electron');
const path = require('path');
const fs = require('fs');
const updater = require('./updater');

// Report a dark color scheme to embedded web content (Outlook on the web honors
// prefers-color-scheme, so this makes the Outlook tab render in dark mode).
nativeTheme.themeSource = 'dark';

// Lightweight JSON settings store (replaces electron-store to avoid packaging issues)
const store = {
  _file: null,
  _data: {},
  _init() {
    if (this._file) return;
    this._file = path.join(app.getPath('userData'), 'settings.json');
    try { this._data = JSON.parse(fs.readFileSync(this._file, 'utf8')); }
    catch (_) { this._data = {}; }
  },
  get(key, fallback = null) {
    this._init();
    return key in this._data ? this._data[key] : fallback;
  },
  set(key, value) {
    this._init();
    this._data[key] = value;
    try { fs.writeFileSync(this._file, JSON.stringify(this._data, null, 2)); } catch (_) {}
  },
};
const APP_ICON = path.join(__dirname, 'assets', 'icon.ico');
let win;

app.whenReady().then(() => {
  win = new BrowserWindow({
    width: 1400,
    height: 860,
    minWidth: 1000,
    minHeight: 640,
    frame: false,
    backgroundColor: '#0d0f14',
    show: false,
    title: 'Dynamics Desk',
    icon: APP_ICON,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      webviewTag: true          // allows <webview> in renderer
    }
  });

  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  win.once('ready-to-show', () => win.show());
});

app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });

ipcMain.handle('get-settings', () => store.get('settings', null));
ipcMain.handle('save-settings', (_, s) => { store.set('settings', s); return true; });

ipcMain.handle('get-version', () => app.getVersion());
ipcMain.handle('updater-check', () => updater.checkForUpdates());
ipcMain.handle('updater-apply', (e, asset) =>
  updater.downloadAndApply(asset, p => e.sender.send('updater-progress', p)));
ipcMain.handle('minimize',  () => win.minimize());
ipcMain.handle('maximize',  () => win.isMaximized() ? win.unmaximize() : win.maximize());
ipcMain.handle('close',     () => win.close());

ipcMain.handle('open-record', (_, { url, title }) => {
  const recWin = new BrowserWindow({
    width: 1280, height: 860, minWidth: 900, minHeight: 600,
    title: title || 'Record',
    backgroundColor: '#0d0f14',
    icon: APP_ICON,
    show: false,
    webPreferences: { partition: 'persist:dynamics', contextIsolation: true, nodeIntegration: false }
  });
  recWin.loadURL(url);
  recWin.once('ready-to-show', () => recWin.show());
  recWin.setMenuBarVisibility(false);
});

ipcMain.handle('open-workorder', (_, { workOrderId, orgUrl, title }) => {
  const woWin = new BrowserWindow({
    width: 1200, height: 840, minWidth: 900, minHeight: 600,
    frame: false, backgroundColor: '#0d0f14',
    title: title || 'Work Order', show: false,
    icon: APP_ICON,
    webPreferences: {
      preload: path.join(__dirname, 'preload-workorder.js'),
      contextIsolation: true, nodeIntegration: false, webviewTag: true
    }
  });
  woWin.loadFile(path.join(__dirname, 'renderer', 'workorder.html'), {
    hash: `bid=${encodeURIComponent(workOrderId)}&org=${encodeURIComponent(orgUrl)}`
  });
  woWin.once('ready-to-show', () => woWin.show());
  woWin.setMenuBarVisibility(false);
});

ipcMain.handle('open-workorder-direct', (_, { workOrderId, orgUrl, title }) => {
  const woWin = new BrowserWindow({
    width: 1200, height: 840, minWidth: 900, minHeight: 600,
    frame: false, backgroundColor: '#0d0f14',
    title: title || 'Work Order', show: false,
    icon: APP_ICON,
    webPreferences: {
      preload: path.join(__dirname, 'preload-workorder.js'),
      contextIsolation: true, nodeIntegration: false, webviewTag: true
    }
  });
  woWin.loadFile(path.join(__dirname, 'renderer', 'workorder.html'), {
    hash: `wo=${encodeURIComponent(workOrderId)}&org=${encodeURIComponent(orgUrl)}`
  });
  woWin.once('ready-to-show', () => woWin.show());
  woWin.setMenuBarVisibility(false);
});

ipcMain.handle('open-contact', (_, { contactId, orgUrl, title }) => {
  const cWin = new BrowserWindow({
    width: 760, height: 640, minWidth: 600, minHeight: 480,
    frame: false, backgroundColor: '#0d0f14',
    title: title || 'Contact', show: false,
    icon: APP_ICON,
    webPreferences: {
      preload: path.join(__dirname, 'preload-workorder.js'),
      contextIsolation: true, nodeIntegration: false, webviewTag: true
    }
  });
  cWin.loadFile(path.join(__dirname, 'renderer', 'contact.html'), {
    hash: `cid=${encodeURIComponent(contactId)}&org=${encodeURIComponent(orgUrl)}`
  });
  cWin.once('ready-to-show', () => cWin.show());
  cWin.setMenuBarVisibility(false);
});

ipcMain.handle('open-team-member', (_, { name, orgUrl, title }) => {
  const tWin = new BrowserWindow({
    width: 760, height: 640, minWidth: 600, minHeight: 480,
    frame: false, backgroundColor: '#0d0f14',
    title: title || 'Team Member', show: false,
    icon: APP_ICON,
    webPreferences: {
      preload: path.join(__dirname, 'preload-workorder.js'),
      contextIsolation: true, nodeIntegration: false, webviewTag: true
    }
  });
  tWin.loadFile(path.join(__dirname, 'renderer', 'team-member.html'), {
    hash: `name=${encodeURIComponent(name)}&org=${encodeURIComponent(orgUrl)}`
  });
  tWin.once('ready-to-show', () => tWin.show());
  tWin.setMenuBarVisibility(false);
});

const { shell } = require('electron');
ipcMain.handle('open-external', (_, url) => shell.openExternal(url));
ipcMain.handle('wo-minimize', e => BrowserWindow.fromWebContents(e.sender)?.minimize());
ipcMain.handle('wo-maximize', e => {
  const w = BrowserWindow.fromWebContents(e.sender);
  w?.isMaximized() ? w.unmaximize() : w?.maximize();
});
ipcMain.handle('wo-close', e => BrowserWindow.fromWebContents(e.sender)?.close());
