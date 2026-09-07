import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './styles/app.css'
import { App } from './App'
import { syncOfficeTheme } from './styles/theme'

// マウントは必ず Office.onReady を待ってから行う——モジュール直下でマウント
// するのは「ブラウザでは動くのにWord/PowerPointでは動かない」バグの
// 最も典型的な原因。
Office.onReady(() => {
  syncOfficeTheme()
  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <App />
    </StrictMode>,
  )
})
