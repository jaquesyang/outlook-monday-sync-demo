import { spawn, spawnSync, type ChildProcess } from 'child_process';
import { readFileSync, writeFileSync, existsSync } from 'fs';
import { resolve } from 'path';

const ROOT = resolve(__dirname, '..');
const MANIFEST_DIR = resolve(ROOT, 'manifest');
const ENV_FILE = resolve(ROOT, '.env');
const NEXT_CONFIG = resolve(ROOT, 'next.config.ts');

const TUNNEL_URL_RE = /https:\/\/[a-z0-9-]+\.trycloudflare\.com/g;

let tunnelUrl: string | null = null;

function updateManifests(url: string) {
  const files = ['manifest.xml.min', 'manifest.xml.legacy', 'manifest.xml.win'];
  for (const f of files) {
    const p = resolve(MANIFEST_DIR, f);
    if (!existsSync(p)) continue;
    let content = readFileSync(p, 'utf-8');
    const matches = content.match(TUNNEL_URL_RE);
    if (matches && matches[0] !== url) {
      content = content.split(matches[0]).join(url);
      writeFileSync(p, content);
      console.log(`  Updated ${f}: ${matches[0]} → ${url}`);
    }
  }
}

function updateEnv(url: string) {
  if (!existsSync(ENV_FILE)) {
    console.log('  .env not found, skipping APP_BASE_URL update');
    return;
  }
  let content = readFileSync(ENV_FILE, 'utf-8');
  const old = content.match(/APP_BASE_URL=(.+)/)?.[1];
  if (old && old.trim() !== url) {
    content = content.replace(/APP_BASE_URL=.*/, `APP_BASE_URL=${url}`);
    writeFileSync(ENV_FILE, content);
    console.log(`  Updated .env: APP_BASE_URL=${url}`);
  }
}

function updateNextConfig(url: string) {
  const hostname = new URL(url).hostname;
  let content = readFileSync(NEXT_CONFIG, 'utf-8');
  const match = content.match(/allowedDevOrigins:\s*\[([^\]]*)\]/);
  if (match) {
    const existing = match[1]
      .split(',')
      .map((s) => s.trim().replace(/['"]/g, ''))
      .filter(Boolean);
    if (!existing.includes(hostname)) {
      const newArr = [...existing, hostname].map((s) => `'${s}'`).join(', ');
      content = content.replace(/allowedDevOrigins:\s*\[[^\]]*\]/, `allowedDevOrigins: [${newArr}]`);
      writeFileSync(NEXT_CONFIG, content);
      console.log(`  Updated next.config.ts: added '${hostname}' to allowedDevOrigins`);
    }
  }
}

function onTunnelUrl(url: string) {
  if (tunnelUrl === url) return;
  tunnelUrl = url;
  console.log(`\nTunnel ready: ${url}\n`);
  updateManifests(url);
  updateEnv(url);
  updateNextConfig(url);
  console.log('\nCopy manifest to Outlook:');
  console.log(`  cp manifest/manifest.xml.min ~/Library/Containers/com.microsoft.Outlook/Data/Documents/WEF/manifest.xml`);
  console.log('');
}

function runTunnel(): ChildProcess {
  console.log('Starting Cloudflare Tunnel → https://localhost:3000 ...');
  const proc = spawn('cloudflared', ['tunnel', '--url', 'https://localhost:3000'], {
    stdio: ['ignore', 'pipe', 'pipe'],
  });

  proc.stdout?.on('data', (data: Buffer) => {
    const line = data.toString();
    const match = line.match(/https:\/\/[a-z0-9-]+\.trycloudflare\.com/);
    if (match) onTunnelUrl(match[0]);
    process.stdout.write(line);
  });

  proc.stderr?.on('data', (data: Buffer) => {
    const line = data.toString();
    const match = line.match(/https:\/\/[a-z0-9-]+\.trycloudflare\.com/);
    if (match) onTunnelUrl(match[0]);
    process.stderr.write(line);
  });

  proc.on('exit', (code) => {
    console.log(`\nTunnel exited with code ${code}`);
    process.exit(code ?? 0);
  });

  return proc;
}

function runDev(): ChildProcess {
  console.log('Starting Next.js dev server ...');
  const proc = spawn('npx', ['next', 'dev', '--turbopack', '--experimental-https'], {
    cwd: ROOT,
    stdio: 'inherit',
  });

  proc.on('exit', (code) => {
    console.log(`\nDev server exited with code ${code}`);
    process.exit(code ?? 0);
  });

  return proc;
}

function checkCloudflared() {
  try {
    const result = spawnSync('cloudflared', ['--version'], { encoding: 'utf-8' });
    if (result.error) throw result.error;
  } catch {
    console.error('cloudflared not found. Install with:');
    console.error('  brew install cloudflared');
    process.exit(1);
  }
}

checkCloudflared();
runDev();
runTunnel();
