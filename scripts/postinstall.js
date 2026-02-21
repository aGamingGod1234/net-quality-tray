#!/usr/bin/env node
const { ensureInstalled } = require('./install-lib');

try {
  const details = ensureInstalled({ setStartupEntry: true });
  console.log(`NetQualityTray installed to ${details.installDir}`);
} catch (error) {
  console.error(`NetQualityTray postinstall failed: ${error.message}`);
  process.exit(1);
}
