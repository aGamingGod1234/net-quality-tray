#!/usr/bin/env node
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const { ensureInstalled, removeStartup, installDir } = require('../scripts/install-lib');

function printHelp() {
  console.log('Usage: nqa-tray [start|install|status|open-folder|remove-startup|help]');
  console.log('  start          Install (if needed) and launch the tray app (default)');
  console.log('  install        Install files + startup entry only');
  console.log('  status         Show install status');
  console.log('  open-folder    Open install folder in Explorer');
  console.log('  remove-startup Remove startup registry entry');
}

function launchExe(exePath) {
  const child = spawn(exePath, [], {
    detached: true,
    stdio: 'ignore',
    windowsHide: true,
  });
  child.unref();
}

function run() {
  if (process.platform !== 'win32') {
    console.error('nqa-tray currently supports Windows only.');
    process.exit(1);
  }

  const command = (process.argv[2] || 'start').toLowerCase();

  if (command === 'help' || command === '--help' || command === '-h') {
    printHelp();
    return;
  }

  if (command === 'status') {
    const exePath = path.join(installDir, 'NQA.exe');
    console.log(`Install directory: ${installDir}`);
    console.log(`NQA.exe present: ${fs.existsSync(exePath)}`);
    return;
  }

  if (command === 'install') {
    const details = ensureInstalled({ setStartupEntry: true });
    console.log(`Installed to ${details.installDir}`);
    return;
  }

  if (command === 'open-folder') {
    spawn('explorer.exe', [installDir], { detached: true, stdio: 'ignore', windowsHide: true }).unref();
    return;
  }

  if (command === 'remove-startup') {
    removeStartup();
    console.log('Startup entry removed.');
    return;
  }

  if (command === 'start') {
    const details = ensureInstalled({ setStartupEntry: true });
    launchExe(details.exePath);
    console.log(`NQA started from ${details.exePath}`);
    return;
  }

  console.error(`Unknown command: ${command}`);
  printHelp();
  process.exit(1);
}

try {
  run();
} catch (error) {
  console.error(`nqa-tray error: ${error.message}`);
  process.exit(1);
}
