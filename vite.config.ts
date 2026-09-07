import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

// Office Add-insはタスクペインをHTTPSで配信する必要がある(開発時も含む)。
// 証明書は `npx office-addin-dev-certs install` で一度だけ生成する。
const certDir = path.join(os.homedir(), '.office-addin-dev-certs')

export default defineConfig(({ command }) => ({
  plugins: [react()],
  // GitHub Pagesはこれをドメイン直下ではなく /Jun-Office-Linter/ という
  // プロジェクトページとして配信するため、本番のアセットURLにはこの
  // プレフィックスが必要。Vite自身のdevサーバーにはこのようなサブパスは
  // ないため、開発時は '/' のまま。
  base: command === 'build' ? '/Jun-Office-Linter/' : '/',
  // `vite`/`vite dev` のときだけ設定する: 設定ファイル評価時にdev証明書を
  // 読み込むと、証明書が存在しない環境（CIなど）では `vite build` が
  // 壊れてしまう。
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
