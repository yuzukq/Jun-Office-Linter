import type { LintRule } from './types'

/**
 * Seed rules for common IME misconversions. Edit freely from the task pane;
 * these only exist so the pane isn't empty on first launch.
 */
export const defaultRules: LintRule[] = [
  {
    id: 'giji-gizi',
    wrong: '疑似',
    correct: '擬似',
    note: '「疑似」は誤変換になりやすい語です。文脈上「擬似」が適切か確認してください。',
  },
  {
    id: 'shokkaku-shokkaku',
    wrong: '触角',
    correct: '触覚',
    note: '「触角」（昆虫などの器官）と「触覚」（感覚）は同音で誤変換されやすい語です。文脈により意図した語が異なる場合は例外リストに追加してください。',
  },
]
