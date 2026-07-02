// Build, zip, and publish a new release to the private GitHub releases repo.
// Usage: npm run release
// Token: set GH_TOKEN env var, or create scripts/release-token.txt with the PAT.
import { execSync, spawnSync } from 'node:child_process';
import { readFileSync, existsSync, statSync, rmSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import https from 'node:https';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = JSON.parse(readFileSync(path.join(root, 'package.json'), 'utf8'));
const cfg = JSON.parse(readFileSync(path.join(root, 'update-config.json'), 'utf8'));
const version = pkg.version;
const tag = 'v' + version;

const token = process.env.GH_TOKEN
  || (existsSync(path.join(root, 'scripts', 'release-token.txt'))
      ? readFileSync(path.join(root, 'scripts', 'release-token.txt'), 'utf8').trim()
      : null);
if (!token) { console.error('No token. Set GH_TOKEN or create scripts/release-token.txt'); process.exit(1); }

const appDir = path.join(root, 'dist-packaged', 'Dynamics Desk-win32-x64');
const zipName = `Dynamics-Desk-${tag}.zip`;
const zipPath = path.join(root, 'dist-packaged', zipName);

function ghRequest(method, url, { headers = {}, body = null } = {}) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const req = https.request({
      method, hostname: u.hostname, path: u.pathname + u.search,
      headers: {
        'User-Agent': 'DynamicsDesk-Publish',
        'Authorization': `token ${token}`,
        'Accept': 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
        ...headers,
      },
    }, res => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        const text = Buffer.concat(chunks).toString();
        if (res.statusCode >= 200 && res.statusCode < 300) resolve(JSON.parse(text || '{}'));
        else reject(new Error(`HTTP ${res.statusCode}: ${text}`));
      });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

(async () => {
  // 1. Build the packaged app
  console.log(`\n▶ Building Dynamics Desk ${tag}…`);
  rmSync(path.join(root, 'dist-packaged'), { recursive: true, force: true });
  execSync('npx electron-packager . "Dynamics Desk" --platform=win32 --arch=x64 ' +
    '--icon=assets/icon.ico --out=dist-packaged --overwrite ' +
    '--ignore=dist-packaged --ignore=dist --ignore=scripts',
    { cwd: root, stdio: 'inherit' });

  // 2. Zip it (use tar — handles Electron's pre-1980 timestamps that break Compress-Archive)
  console.log('▶ Zipping…');
  rmSync(zipPath, { force: true });
  const distDir = path.join(root, 'dist-packaged');
  spawnSync('tar.exe', ['--force-local', '-a', '-c', '-f', zipPath, '-C', distDir, 'Dynamics Desk-win32-x64'],
    { stdio: 'inherit' });
  if (!existsSync(zipPath)) { console.error('Zip failed'); process.exit(1); }
  console.log(`  ${zipName} (${(statSync(zipPath).size / 1048576).toFixed(1)} MB)`);

  // 3. Compile the Inno Setup installer (for fresh installs; auto-updater uses the zip)
  console.log('▶ Building installer…');
  const iscc = path.join(process.env.LOCALAPPDATA || '', 'Programs', 'Inno Setup 6', 'ISCC.exe');
  const setupName = `Dynamics-Desk-Setup-${tag}.exe`;
  const setupPath = path.join(root, 'dist-installer', setupName);
  if (existsSync(iscc)) {
    const r = spawnSync(iscc, [`/DAppVersion=${version}`, path.join(root, 'installer.iss')],
      { cwd: root, stdio: 'inherit' });
    if (r.status !== 0 || !existsSync(setupPath)) { console.error('Installer build failed'); process.exit(1); }
    console.log(`  ${setupName} (${(statSync(setupPath).size / 1048576).toFixed(1)} MB)`);
  } else {
    console.warn('  Inno Setup not found — skipping installer (zip only).');
  }

  // 4. Create the GitHub release
  console.log(`▶ Creating release ${tag} on ${cfg.owner}/${cfg.repo}…`);
  let release;
  try {
    release = await ghRequest('POST', `https://api.github.com/repos/${cfg.owner}/${cfg.repo}/releases`,
      { body: JSON.stringify({ tag_name: tag, name: tag, body: `Dynamics Desk ${tag}`, draft: false, prerelease: false }) });
  } catch (e) {
    console.error('Could not create release (does the tag already exist?):\n', e.message);
    process.exit(1);
  }

  // 5. Upload assets — installer first (fresh installs), then zip (auto-updater)
  const assets = [];
  if (existsSync(setupPath)) assets.push({ file: setupPath, name: setupName, type: 'application/vnd.microsoft.portable-executable' });
  assets.push({ file: zipPath, name: zipName, type: 'application/zip' });
  for (const a of assets) {
    console.log(`▶ Uploading ${a.name}…`);
    const data = readFileSync(a.file);
    const uploadUrl = `https://uploads.github.com/repos/${cfg.owner}/${cfg.repo}/releases/${release.id}/assets?name=${encodeURIComponent(a.name)}`;
    await ghRequest('POST', uploadUrl, {
      headers: { 'Content-Type': a.type, 'Content-Length': data.length },
      body: data,
    });
  }

  console.log(`\n✅ Published ${tag}. Installer + auto-update zip are live.\n`);
})().catch(e => { console.error(e); process.exit(1); });
