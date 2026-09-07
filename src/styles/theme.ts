/**
 * Office.context.officeTheme を tokens.css のCSSカスタムプロパティに
 * マッピングし、ダークモードのWord/PowerPointウィンドウの中で
 * ペインだけ白い四角形になってしまわないよう、背景・文字色をホストアプリの
 * テーマに合わせる。`--shu` のアクセントカラーだけは変更しない
 * ——これはこのアプリの署名的な色なので、テーマが変わっても一定に保ちたい。
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
