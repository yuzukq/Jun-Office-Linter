import { useRef, useState } from 'react'
import type { LintRule } from '../core/types'
import { createRule, exportRulesAsJson, importRulesFromJson } from '../core/ruleStore'
import { CorrectionMark } from './CorrectionMark'

export function DictionaryEditor({
  rules,
  onChange,
}: {
  rules: LintRule[]
  onChange: (rules: LintRule[]) => void
}) {
  const [wrong, setWrong] = useState('')
  const [correct, setCorrect] = useState('')
  const [note, setNote] = useState('')
  const [importError, setImportError] = useState<string | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

  function handleAdd() {
    if (!wrong.trim() || !correct.trim()) return
    onChange([...rules, createRule(wrong.trim(), correct.trim(), note.trim() || undefined)])
    setWrong('')
    setCorrect('')
    setNote('')
  }

  function handleDelete(id: string) {
    onChange(rules.filter((r) => r.id !== id))
  }

  function handleExport() {
    const blob = new Blob([exportRulesAsJson(rules)], { type: 'application/json' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'jun-office-linter-dictionary.json'
    a.click()
    URL.revokeObjectURL(url)
  }

  function handleImportClick() {
    setImportError(null)
    fileInputRef.current?.click()
  }

  async function handleFileSelected(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    e.target.value = ''
    if (!file) return
    try {
      const text = await file.text()
      onChange(importRulesFromJson(text))
    } catch (err) {
      setImportError(err instanceof Error ? err.message : 'JSONの読み込みに失敗しました')
    }
  }

  return (
    <div>
      <div className="summary-bar">
        <span className="summary-count">{rules.length} 件の禁止語</span>
        <div style={{ display: 'flex', gap: 'var(--space-2)' }}>
          <button className="fix-all-button" onClick={handleImportClick}>
            インポート
          </button>
          <button className="fix-all-button" onClick={handleExport}>
            エクスポート
          </button>
        </div>
      </div>
      <input
        ref={fileInputRef}
        type="file"
        accept="application/json"
        style={{ display: 'none' }}
        onChange={handleFileSelected}
      />
      {importError && (
        <div className="rule-group-note" style={{ color: 'var(--shu)' }}>
          {importError}
        </div>
      )}

      <div className="rule-group">
        {rules.map((rule) => (
          <div className="occurrence-row" key={rule.id} style={{ cursor: 'default' }}>
            <span className="occurrence-snippet" style={{ overflow: 'visible', whiteSpace: 'normal' }}>
              <CorrectionMark wrong={rule.wrong} correct={rule.correct} />
              {rule.note && <div className="rule-group-note" style={{ padding: '4px 0 0' }}>{rule.note}</div>}
            </span>
            <span
              className="occurrence-fix-button"
              role="button"
              tabIndex={0}
              onClick={() => handleDelete(rule.id)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' || e.key === ' ') handleDelete(rule.id)
              }}
            >
              削除
            </span>
          </div>
        ))}
        {rules.length === 0 && (
          <div className="empty-state" style={{ padding: 'var(--space-3)' }}>
            禁止語が登録されていません。下のフォームから追加してください。
          </div>
        )}
      </div>

      <div className="rule-group" style={{ padding: 'var(--space-3)' }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 'var(--space-2)' }}>
          <input
            placeholder="誤って変換されやすい語（例：疑似）"
            value={wrong}
            onChange={(e) => setWrong(e.target.value)}
            style={inputStyle}
          />
          <input
            placeholder="正しい語（例：擬似）"
            value={correct}
            onChange={(e) => setCorrect(e.target.value)}
            style={inputStyle}
          />
          <input
            placeholder="メモ（任意）"
            value={note}
            onChange={(e) => setNote(e.target.value)}
            style={inputStyle}
          />
          <button className="scan-button" onClick={handleAdd} disabled={!wrong.trim() || !correct.trim()}>
            追加
          </button>
        </div>
      </div>
    </div>
  )
}

const inputStyle: React.CSSProperties = {
  fontFamily: 'inherit',
  fontSize: 13,
  padding: '6px 8px',
  border: '1px solid var(--line)',
  borderRadius: 6,
  background: 'var(--paper)',
  color: 'var(--ink)',
}
