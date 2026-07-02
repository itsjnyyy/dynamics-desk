// Lightweight self-updater for Dynamics Desk.
// Checks a private GitHub repo's latest release, downloads the packaged zip,
// swaps the installed app folder in place, and relaunches. Windows-only.
const { app } = require('electron');
const path = require('path');
const fs = require('fs');
const https = require('https');
const { spawn } = require('child_process');

const CONFIG = (() => {
  try { return JSON.parse(fs.readFileSync(path.join(__dirname, 'update-config.json'), 'utf8')); }
  catch (_) { return null; }
})();

function isConfigured() {
  return !!(CONFIG && CONFIG.owner && CONFIG.repo && CONFIG.token &&
           !CONFIG.token.startsWith('PASTE'));
}

// semver-ish compare: returns 1 if a>b, -1 if a<b, 0 if equal
function cmpVersions(a, b) {
  const pa = String(a).replace(/^v/, '').split('.').map(n => parseInt(n, 10) || 0);
  const pb = String(b).replace(/^v/, '').split('.').map(n => parseInt(n, 10) || 0);
  for (let i = 0; i < 3; i++) {
    if ((pa[i] || 0) > (pb[i] || 0)) return 1;
    if ((pa[i] || 0) < (pb[i] || 0)) return -1;
  }
  return 0;
}

function fetchJson(url) {
  return new Promise((resolve, reject) => {
    const req = https.request(url, {
      method: 'GET',
      headers: {
        'User-Agent': 'DynamicsDesk-Updater',
        'Authorization': `token ${CONFIG.token}`,
        'Accept': 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
      },
    }, res => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        try { resolve({ status: res.statusCode, json: JSON.parse(Buffer.concat(chunks).toString()) }); }
        catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

// Download a URL to a file, following redirects. Auth header only sent to api.github.com.
function downloadToFile(url, dest, onProgress, redirectsLeft = 5) {
  return new Promise((resolve, reject) => {
    const isGitHubApi = url.startsWith('https://api.github.com/');
    const headers = { 'User-Agent': 'DynamicsDesk-Updater' };
    if (isGitHubApi) {
      headers['Authorization'] = `token ${CONFIG.token}`;
      headers['Accept'] = 'application/octet-stream';
      headers['X-GitHub-Api-Version'] = '2022-11-28';
    }
    const req = https.request(url, { method: 'GET', headers }, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        res.resume();
        if (redirectsLeft <= 0) return reject(new Error('Too many redirects'));
        return resolve(downloadToFile(res.headers.location, dest, onProgress, redirectsLeft - 1));
      }
      if (res.statusCode !== 200) { res.resume(); return reject(new Error('HTTP ' + res.statusCode)); }
      const total = parseInt(res.headers['content-length'] || '0', 10);
      let received = 0;
      const file = fs.createWriteStream(dest);
      res.on('data', c => {
        received += c.length;
        if (onProgress && total) onProgress(Math.round((received / total) * 100));
      });
      res.pipe(file);
      file.on('finish', () => file.close(() => resolve()));
      file.on('error', reject);
    });
    req.on('error', reject);
    req.end();
  });
}

async function checkForUpdates() {
  if (!isConfigured()) return { ok: false, reason: 'not-configured' };
  try {
    const url = `https://api.github.com/repos/${CONFIG.owner}/${CONFIG.repo}/releases/latest`;
    const { status, json } = await fetchJson(url);
    if (status !== 200) return { ok: false, reason: 'http-' + status };
    const latest = json.tag_name || '';
    const current = app.getVersion();
    const asset = (json.assets || []).find(a => a.name.toLowerCase().endsWith('.zip'));
    return {
      ok: true,
      updateAvailable: cmpVersions(latest, current) > 0 && !!asset,
      current,
      latest,
      notes: json.body || '',
      asset: asset ? { url: asset.url, name: asset.name, size: asset.size } : null,
    };
  } catch (e) {
    return { ok: false, reason: e.message };
  }
}

function extractZip(zip, destDir) {
  // Use the system tar (bsdtar) — it reads our tar-created zips, which Expand-Archive cannot.
  return new Promise((resolve, reject) => {
    const tar = spawn('tar.exe', ['-x', '-f', zip], { cwd: destDir, windowsHide: true });
    let err = '';
    tar.stderr.on('data', d => { err += d.toString(); });
    tar.on('error', reject);
    tar.on('exit', code => code === 0 ? resolve() : reject(new Error('Extraction failed (' + code + '): ' + err.trim())));
  });
}

// Find the extracted app folder (the one containing "Dynamics Desk.exe")
function findAppDir(root) {
  const direct = path.join(root, 'Dynamics Desk.exe');
  if (fs.existsSync(direct)) return root;
  for (const entry of fs.readdirSync(root, { withFileTypes: true })) {
    if (entry.isDirectory()) {
      const candidate = path.join(root, entry.name);
      if (fs.existsSync(path.join(candidate, 'Dynamics Desk.exe'))) return candidate;
    }
  }
  return null;
}

async function downloadAndApply(asset, onProgress) {
  if (!app.isPackaged) throw new Error('Updates can only be applied to the installed app.');
  if (!asset || !asset.url) throw new Error('No update asset available.');

  const updatesDir = path.join(app.getPath('userData'), 'updates');
  fs.mkdirSync(updatesDir, { recursive: true });
  // Best-effort cleanup of any leftover staging dirs from previous runs (AV can briefly
  // lock freshly-extracted files, so retry rather than fail).
  try {
    for (const e of fs.readdirSync(updatesDir)) {
      if (e.startsWith('staging')) {
        fs.rmSync(path.join(updatesDir, e), { recursive: true, force: true, maxRetries: 8, retryDelay: 250 });
      } else if (e.endsWith('.zip') || e.endsWith('.bat') || e.endsWith('.vbs')) {
        try { fs.rmSync(path.join(updatesDir, e), { force: true, maxRetries: 4, retryDelay: 200 }); } catch (_) {}
      }
    }
  } catch (_) {}
  // Use a fresh, uniquely-named staging dir so we never have to delete a locked one.
  const stagingDir = path.join(updatesDir, 'staging-' + Date.now());
  fs.mkdirSync(stagingDir, { recursive: true });

  const zipPath = path.join(updatesDir, asset.name);
  await downloadToFile(asset.url, zipPath, onProgress);
  await extractZip(zipPath, stagingDir);

  const newAppDir = findAppDir(stagingDir);
  if (!newAppDir) throw new Error('Downloaded update looks invalid (app not found).');

  const installDir = path.dirname(process.execPath);
  const exePath = path.join(installDir, 'Dynamics Desk.exe');
  const pid = process.pid;
  const batchPath = path.join(updatesDir, 'apply-update.bat');
  // Wait for THIS specific main process (by PID) to exit — robust against Electron's
  // multiple "Dynamics Desk.exe" helper processes — then swap the folder and relaunch.
  const batch = [
    '@echo off',
    ':wait',
    `tasklist /fi "PID eq ${pid}" 2>nul | find "${pid}" >nul`,
    'if not errorlevel 1 ( timeout /t 1 /nobreak >nul & goto wait )',
    'timeout /t 1 /nobreak >nul',
    `robocopy "${newAppDir}" "${installDir}" /MIR /R:2 /W:1 >nul`,
    `start "" "${exePath}"`,
    `rmdir /s /q "${stagingDir}"`,
    'del "%~f0"',
  ].join('\r\n');
  fs.writeFileSync(batchPath, batch, 'utf8');

  // Launch the batch fully hidden (no console window) and detached, via a VBScript shim.
  const vbsPath = path.join(updatesDir, 'apply-update.vbs');
  fs.writeFileSync(vbsPath,
    `CreateObject("Wscript.Shell").Run "cmd /c ""${batchPath}""", 0, False\r\n`, 'utf8');
  spawn('wscript.exe', [vbsPath], { detached: true, stdio: 'ignore', windowsHide: true }).unref();
  setTimeout(() => app.quit(), 400);
  return { ok: true };
}

module.exports = { checkForUpdates, downloadAndApply, isConfigured };
