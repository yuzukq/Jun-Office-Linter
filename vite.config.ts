import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

// Office Add-ins require the task pane to be served over HTTPS, even in dev.
// Certs are generated once via `npx office-addin-dev-certs install`.
const certDir = path.join(os.homedir(), '.office-addin-dev-certs')

export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    strictPort: true,
    https: {
      key: fs.readFileSync(path.join(certDir, 'localhost.key')),
      cert: fs.readFileSync(path.join(certDir, 'localhost.crt')),
    },
  },
})
