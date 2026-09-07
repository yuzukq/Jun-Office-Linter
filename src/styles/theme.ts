/**
 * Maps Office.context.officeTheme onto the CSS custom properties from
 * tokens.css, so the pane's background/text match the host app's theme
 * (including dark mode) instead of rendering a white rectangle inside a
 * dark Word/PowerPoint window. The `--shu` accent is left untouched —
 * it's the app's signature color and should stay constant across themes.
 */
export function syncOfficeTheme(): void {
  const officeTheme = (Office.context as { officeTheme?: Office.OfficeTheme }).officeTheme
  if (!officeTheme) return

  const root = document.documentElement.style
  if (officeTheme.bodyBackgroundColor) root.setProperty('--paper', officeTheme.bodyBackgroundColor)
  if (officeTheme.bodyForegroundColor) root.setProperty('--ink', officeTheme.bodyForegroundColor)
  if (officeTheme.controlBackgroundColor) root.setProperty('--control-bg', officeTheme.controlBackgroundColor)
  if (officeTheme.controlForegroundColor) root.setProperty('--muted', officeTheme.controlForegroundColor)
}
