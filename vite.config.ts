import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

// Office Add-ins require the task pane to be served over HTTPS, even in dev.
// Certs are generated once via `npx office-addin-dev-certs install`.
const certDir = path.join(os.homedir(), '.office-addin-dev-certs')

export default defineConfig(({ command }) => ({
  plugins: [react()],
  // GitHub Pages serves this as a project page at /Jun-Office-Linter/, not
  // at the domain root, so production asset URLs need that prefix. Dev
  // stays at '/' since Vite's own server has no such subpath.
  base: command === 'build' ? '/Jun-Office-Linter/' : '/',
  // Only wired up for `vite`/`vite dev`: reading dev certs at config-eval
  // time would break `vite build` anywhere the certs don't exist, e.g. CI.
  server:
    command === 'build'
      ? undefined
      : {
          port: 3000,
          strictPort: true,
          https: {
            key: fs.readFileSync(path.join(certDir, 'localhost.key')),
            cert: fs.readFileSync(path.join(certDir, 'localhost.crt')),
          },
        },
}))
