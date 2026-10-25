#!/usr/bin/env node
/**
 * Builds the production web bundle and publishes static assets to the repository root
 * so GitHub Pages (serving from branch `main` `/`) immediately hosts the compiled app.
 */
import fs from 'node:fs';
import path from 'node:path';
import { execSync } from 'node:child_process';

const rootDir = process.cwd();
const templatePath = path.join(rootDir, 'index.template.html');
const indexPath = path.join(rootDir, 'index.html');
const distDir = path.join(rootDir, 'dist');
const rootAssetsDir = path.join(rootDir, 'assets');

// 1. Restore Vite source entry template into index.html before building
fs.copyFileSync(templatePath, indexPath);

// 2. Rebuild verified data from Stock list.xlsx & NOV. SHIPMENT -OCT22.xlsx
execSync('python3 scripts/build_data.py', { cwd: rootDir, stdio: 'inherit' });

// 3. Run Vite production build
execSync('npx vite build', { cwd: rootDir, stdio: 'inherit' });

// 4. Copy compiled dist/index.html & dist/assets/* to root for GitHub Pages (main /)
if (fs.existsSync(rootAssetsDir)) {
  fs.rmSync(rootAssetsDir, { recursive: true, force: true });
}
fs.cpSync(path.join(distDir, 'assets'), rootAssetsDir, { recursive: true });
fs.copyFileSync(path.join(distDir, 'index.html'), indexPath);
fs.writeFileSync(path.join(rootDir, '.nojekyll'), '');

console.log('✅ Built and prepared static GitHub Pages bundle at repository root.');
