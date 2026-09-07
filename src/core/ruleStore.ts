import { defaultRules } from './defaultRules'
import type { LintRule } from './types'

const STORAGE_KEY = 'jun-office-linter.rules.v1'

/**
 * Rules live in localStorage rather than per-document settings: a
 * research-wide term list should follow the user across every Word/PowerPoint
 * file, not reset per document.
 */
export function loadRules(): LintRule[] {
  const raw = localStorage.getItem(STORAGE_KEY)
  if (!raw) return defaultRules

  try {
    const parsed = JSON.parse(raw)
    if (Array.isArray(parsed)) return parsed as LintRule[]
  } catch {
    // fall through to defaults on corrupt storage
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
