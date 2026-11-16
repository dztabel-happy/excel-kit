#!/usr/bin/env node
'use strict';

const { spawn } = require('node:child_process');
const { join, dirname } = require('node:path');
const { constants } = require('node:os');
const { version } = require('../package.json');

function platformPackage(platform, arch) {
  // The Windows x64 binary also runs under Windows ARM64 emulation.
  if (platform === 'win32' && (arch === 'x64' || arch === 'arm64')) return '@dztabel/excel-kit-win32-x64';
  if (platform === 'darwin' && (arch === 'arm64' || arch === 'x64')) return `@dztabel/excel-kit-darwin-${arch}`;
  throw new Error(`Unsupported platform ${platform}/${arch}. This release supports macOS arm64/x64 and Windows x64 (including Windows ARM64 emulation).`);
}

function main() {
  let executable;
  try {
    const name = platformPackage(process.platform, process.arch);
    let manifestPath;
    try { manifestPath = require.resolve(`${name}/package.json`); }
    catch { throw new Error(`Missing ${name}@${version}. Reinstall with npm install --include=optional @dztabel/excel-kit@${version} (add -g for a global installation).`); }
    const installed = require(manifestPath);
    if (installed.version !== version) throw new Error(`Expected ${name}@${version}, found ${installed.version}. Reinstall @dztabel/excel-kit@${version}.`);
    executable = join(dirname(manifestPath), 'bin', process.platform === 'win32' ? 'excel-kit.exe' : 'excel-kit');
  } catch (error) {
    process.stderr.write(`excel-kit: ${error.message}\n`);
    process.exitCode = 1;
    return;
  }

  const child = spawn(executable, process.argv.slice(2), { stdio: 'inherit', windowsHide: true });
  const signals = ['SIGINT', 'SIGTERM'];
  const handlers = new Map(signals.map(signal => [signal, () => { if (child.exitCode === null) child.kill(signal); }]));
  for (const [signal, handler] of handlers) process.on(signal, handler);
  function unhook() { for (const [signal, handler] of handlers) process.removeListener(signal, handler); }
  child.once('error', error => {
    unhook();
    process.stderr.write(`excel-kit: Cannot start the installed program: ${error.message}\n`);
    process.exitCode = 1;
  });
  child.once('exit', (code, signal) => {
    unhook();
    if (signal && process.platform !== 'win32') process.kill(process.pid, signal);
    else process.exitCode = code ?? (128 + (constants.signals[signal] || 1));
  });
}

if (require.main === module) main();
module.exports = { platformPackage };
