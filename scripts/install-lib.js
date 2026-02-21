#!/usr/bin/env node
const { spawnSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const packageRoot = path.resolve(__dirname, '..');
const localAppData = process.env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Local');
const installDir = path.join(localAppData, 'NetQualityTray');
const startupRegPath = 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run';
const startupRegValue = 'NQA';

function copyPath(source, destination) {
  const stat = fs.statSync(source);
  if (stat.isDirectory()) {
    fs.cpSync(source, destination, { recursive: true, force: true });
    return;
  }
  fs.mkdirSync(path.dirname(destination), { recursive: true });
  fs.copyFileSync(source, destination);
}

function setStartup(exePath) {
  const quotedExe = `\"${exePath}\"`;
  const result = spawnSync('reg', ['add', startupRegPath, '/v', startupRegValue, '/t', 'REG_SZ', '/d', quotedExe, '/f'], {
    encoding: 'utf8',
  });

  if (result.status !== 0) {
    const errorText = result.stderr || result.stdout || 'Unknown registry error';
    throw new Error(`Failed to set startup registry value: ${errorText.trim()}`);
  }
}

function removeStartup() {
  const result = spawnSync('reg', ['delete', startupRegPath, '/v', startupRegValue, '/f'], {
    encoding: 'utf8',
  });

  if (result.status !== 0) {
    const text = `${result.stderr || ''}${result.stdout || ''}`.trim();
    if (text && !/unable to find|cannot find/i.test(text)) {
      throw new Error(`Failed to remove startup registry value: ${text}`);
    }
  }
}

function ensureInstalled(options = {}) {
  const setStartupEntry = options.setStartupEntry !== false;

  if (process.platform !== 'win32') {
    throw new Error('NetQualityTray currently supports Windows only.');
  }

  fs.mkdirSync(installDir, { recursive: true });

  const requiredFiles = [
    'NQA.exe',
    'settings.json',
    'NetworkQualityTray.ps1',
    'Start-NetworkQualityTray.vbs',
    'Build-NativeApp.ps1',
  ];

  for (const rel of requiredFiles) {
    const source = path.join(packageRoot, rel);
    if (fs.existsSync(source)) {
      copyPath(source, path.join(installDir, rel));
    }
  }

  const optionalDirs = ['assets', 'NativeApp'];
  for (const rel of optionalDirs) {
    const source = path.join(packageRoot, rel);
    if (fs.existsSync(source)) {
      copyPath(source, path.join(installDir, rel));
    }
  }

  const exePath = path.join(installDir, 'NQA.exe');
  if (!fs.existsSync(exePath)) {
    throw new Error(`NQA.exe was not found after install: ${exePath}`);
  }

  if (setStartupEntry) {
    setStartup(exePath);
  }

  return { installDir, exePath };
}

module.exports = {
  ensureInstalled,
  removeStartup,
  installDir,
};
