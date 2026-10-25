#!/usr/bin/env node
/**
 * Standalone OneDrive & Repository Realtime Watcher for `Stock list.xlsx`
 * Automatically detects changes in OneDrive `Stock list.xlsx`, updates the repository,
 * rebuilds the Multi-Agent verified JSON dataset, and optionally commits & pushes to `main`.
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { execSync } from 'node:child_process';

const ROOT_DIR = process.cwd();
const REPO_STOCK_PATH = path.join(ROOT_DIR, 'Stock list.xlsx');

// Candidate OneDrive paths (configurable via ONE_DRIVE_STOCK_PATH or ONE_DRIVE_DIR)
const candidatePaths = [
  process.env.ONE_DRIVE_STOCK_PATH,
  process.env.ONE_DRIVE_DIR ? path.join(process.env.ONE_DRIVE_DIR, 'Stock list.xlsx') : null,
  path.join(os.homedir(), 'OneDrive', 'Stock list.xlsx'),
  path.join(os.homedir(), 'OneDrive', 'stock list', 'Stock list.xlsx'),
  path.join(os.homedir(), 'OneDrive - LPVN', 'Stock list.xlsx'),
  REPO_STOCK_PATH
].filter(Boolean);

console.log('==============================================================');
console.log('🚀 LPVN OneDrive -> Repository Realtime Sync Service Started');
console.log('==============================================================');

let lastMtimeMap = new Map();

function syncFromSource(sourcePath) {
  try {
    const stat = fs.statSync(sourcePath);
    const prevMtime = lastMtimeMap.get(sourcePath) || 0;
    if (stat.mtimeMs <= prevMtime) return;
    lastMtimeMap.set(sourcePath, stat.mtimeMs);

    console.log(`[${new Date().toISOString()}] Detected update in: ${sourcePath}`);
    if (path.resolve(sourcePath) !== path.resolve(REPO_STOCK_PATH)) {
      fs.copyFileSync(sourcePath, REPO_STOCK_PATH);
      console.log(`  -> Copied updated Stock list.xlsx into repo: ${REPO_STOCK_PATH}`);
    }

    execSync('python3 scripts/build_data.py', { cwd: ROOT_DIR, stdio: 'inherit' });

    if (process.env.AUTO_GIT_PUSH === 'true') {
      try {
        execSync(
          'git add "Stock list.xlsx" public/data/initial-data.json && git commit -m "chore(onedrive): realtime auto-update Stock list.xlsx" && git push origin main',
          { cwd: ROOT_DIR, stdio: 'inherit' }
        );
        console.log('  -> Committed and pushed updated Stock list.xlsx to origin/main');
      } catch {
        console.log('  -> No git changes to push or push skipped.');
      }
    }
  } catch (err) {
    console.error('Sync error:', err.message);
  }
}

setInterval(() => {
  for (const p of candidatePaths) {
    if (p && fs.existsSync(p)) {
      syncFromSource(p);
    }
  }
}, 2000);
