import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './styles/app.css'
import { App } from './App'
import { syncOfficeTheme } from './styles/theme'

// Mounting must wait for Office.onReady — mounting at module top level is
// the most common "works in the browser, breaks in Word/PowerPoint" bug.
Office.onReady(() => {
  syncOfficeTheme()
  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <App />
    </StrictMode>,
  )
})
