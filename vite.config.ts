import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import fs from 'node:fs';
import path from 'node:path';
import { execSync } from 'node:child_process';

// Built-in Repository Sync Middleware for Vite
function realtimeRepoSyncPlugin() {
  return {
    name: 'realtime-repo-sync',
    configureServer(server: any) {
      const rootDir = process.cwd();
      const stockPath = path.join(rootDir, 'Stock list.xlsx');
      const shipPath = path.join(rootDir, 'NOV. SHIPMENT -OCT22.xlsx');

      // Watch Stock list.xlsx & NOV. SHIPMENT -OCT22.xlsx for local changes
      let debounceTimer: ReturnType<typeof setTimeout> | null = null;
      const triggerRebuildData = (changedFile: string) => {
        if (debounceTimer) clearTimeout(debounceTimer);
        debounceTimer = setTimeout(() => {
          try {
            execSync('python3 scripts/build_data.py', { cwd: rootDir, stdio: 'inherit' });
            server.ws.send({
              type: 'custom',
              event: 'repo-data-updated',
              data: { file: changedFile, timestamp: new Date().toISOString() }
            });
          } catch (err) {
            console.error('Error rebuilding dataset:', err);
          }
        }, 400);
      };

      for (const fpath of [stockPath, shipPath]) {
        if (fs.existsSync(fpath)) {
          fs.watchFile(fpath, { interval: 1500 }, (curr, prev) => {
            if (curr.mtimeMs !== prev.mtimeMs) {
              triggerRebuildData(path.basename(fpath));
            }
          });
        }
      }

      server.middlewares.use('/api/sync-file', (req: any, res: any, next: any) => {
        if (req.method !== 'POST') return next();
        const chunks: Buffer[] = [];
        req.on('data', (chunk: Buffer) => chunks.push(chunk));
        req.on('end', () => {
          try {
            const body = JSON.parse(Buffer.concat(chunks).toString('utf-8'));
            const { fileType, fileName, base64Data, autoGitPush } = body;
            const targetName =
              fileType === 'stock'
                ? 'Stock list.xlsx'
                : fileType === 'monthly'
                ? 'NOV. SHIPMENT -OCT22.xlsx'
                : fileName || 'Weekly_Plan.xlsx';
            const targetPath = path.join(rootDir, targetName);
            const buf = Buffer.from(base64Data, 'base64');
            fs.writeFileSync(targetPath, buf);

            const publicFilesDir = path.join(rootDir, 'public', 'files');
            fs.mkdirSync(publicFilesDir, { recursive: true });
            fs.writeFileSync(path.join(publicFilesDir, targetName), buf);

            execSync('python3 scripts/build_data.py', { cwd: rootDir });

            let gitStatus = 'Saved to repo workspace & rebuilt dataset';
            if (autoGitPush) {
              try {
                execSync(
                  `git add "${targetName}" public/data/initial-data.json && git commit -m "chore(sync): update ${targetName} from user import" && git push origin main`,
                  { cwd: rootDir }
                );
                gitStatus = 'Committed & pushed to origin/main';
              } catch (gitErr: any) {
                gitStatus = 'Saved locally (git commit skipped or no diff)';
              }
            }

            res.setHeader('Content-Type', 'application/json');
            res.end(
              JSON.stringify({
                ok: true,
                targetName,
                gitStatus,
                timestamp: new Date().toISOString()
              })
            );
          } catch (err: any) {
            res.statusCode = 500;
            res.setHeader('Content-Type', 'application/json');
            res.end(JSON.stringify({ ok: false, error: err?.message || 'Unknown error' }));
          }
        });
      });
    }
  };
}

export default defineConfig({
  base: './',
  plugins: [react(), tailwindcss(), realtimeRepoSyncPlugin()],
  build: {
    target: 'esnext',
    chunkSizeWarningLimit: 2000
  }
});
