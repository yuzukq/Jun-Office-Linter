import { defaultRules } from './defaultRules'
import type { LintRule } from './types'

const STORAGE_KEY = 'jun-office-linter.rules.v1'

/**
 * ルールは文書単位の設定ではなくlocalStorageに保存する。研究テーマ全体で使う
 * 用語集は、文書ごとにリセットされるのではなく、Word/PowerPointのどのファイルを
 * 開いてもユーザーについてきてほしいため。
 */
export function loadRules(): LintRule[] {
  const raw = localStorage.getItem(STORAGE_KEY)
  if (!raw) return defaultRules

  try {
    const parsed = JSON.parse(raw)
    if (Array.isArray(parsed)) return parsed as LintRule[]
  } catch {
    // 保存データが壊れている場合はデフォルトにフォールバック
  }
  return defaultRules
}

export function saveRules(rules: LintRule[]): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(rules))
}

export function exportRulesAsJson(rules: LintRule[]): string {
  return JSON.stringify(rules, null, 2)
}

export function importRulesFromJson(json: string): LintRule[] {
  const parsed = JSON.parse(json)
  if (!Array.isArray(parsed)) throw new Error('JSON must be an array of rules')

  for (const rule of parsed) {
    if (typeof rule.id !== 'string' || typeof rule.wrong !== 'string' || typeof rule.correct !== 'string') {
      throw new Error('Each rule needs at least id, wrong, and correct as strings')
    }
  }

  return parsed as LintRule[]
}

export function createRule(wrong: string, correct: string, note?: string): LintRule {
  return {
    id: `rule-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    wrong,
    correct,
    note,
  }
}
