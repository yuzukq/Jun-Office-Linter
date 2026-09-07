import { useEffect, useMemo, useState } from 'react'
import type { HostAdapter, Occurrence } from './adapters/types'
import { WordAdapter } from './adapters/wordAdapter'
import { PowerPointAdapter } from './adapters/powerPointAdapter'
import { IssueList } from './components/IssueList'
import { DictionaryEditor } from './components/DictionaryEditor'
import { loadRules, saveRules } from './core/ruleStore'
import type { LintRule } from './core/types'

type Tab = 'check' | 'dictionary'

function createAdapter(): HostAdapter | null {
  switch (Office.context.host) {
    case Office.HostType.Word:
      return new WordAdapter()
    case Office.HostType.PowerPoint:
      return new PowerPointAdapter()
    default:
      return null
  }
}

function describeError(err: unknown): string {
  if (err && typeof err === 'object' && 'message' in err) return String((err as { message: unknown }).message)
  return String(err)
}

export function App() {
  const [tab, setTab] = useState<Tab>('check')
  const [rules, setRules] = useState<LintRule[]>(() => loadRules())
  const [occurrences, setOccurrences] = useState<Occurrence[]>([])
  const [scanning, setScanning] = useState(false)
  const [scanned, setScanned] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const adapter = useMemo(() => createAdapter(), [])

  useEffect(() => {
    saveRules(rules)
  }, [rules])

  useEffect(() => {
    return () => {
      void adapter?.dispose()
    }
  }, [adapter])

  async function runScan() {
    if (!adapter) return
    setScanning(true)
    setError(null)
    try {
      const found = await adapter.scan(rules)
      setOccurrences(found)
      setScanned(true)
    } catch (err) {
      setError(`スキャンに失敗しました: ${describeError(err)}`)
    } finally {
      setScanning(false)
    }
  }

  async function handleSelect(id: string) {
    try {
      await adapter?.selectOccurrence(id)
    } catch (err) {
      setError(`該当箇所を選択できませんでした: ${describeError(err)}`)
    }
  }

  async function handleApplyOne(id: string) {
    try {
      await adapter?.applyOccurrence(id)
      // 一覧をクライアント側でフィルタするのではなく再スキャンする: 修正を
      // 適用すると同じシェイプ/段落内の他の未処理のマッチの文字オフセットが
      // ずれる可能性があるため、正しい状態は常に再スキャンで得る。
      await runScan()
    } catch (err) {
      setError(`修正を適用できませんでした: ${describeError(err)}`)
    }
  }

  async function handleApplyAll(ruleId: string) {
    if (!adapter) return
    try {
      await adapter.applyAllForRule(ruleId)
      await runScan()
    } catch (err) {
      setError(`一括修正に失敗しました: ${describeError(err)}`)
    }
  }

  if (!adapter) {
    return (
      <div className="app-shell">
        <div className="empty-state">このアドインは Word または PowerPoint でのみ動作します。</div>
      </div>
    )
  }

  return (
    <div className="app-shell">
      {/* ここにアプリ側のタイトルは置かない: Office自体のタスクペインの
          UI枠が、このコンテンツの上に既にアドインの表示名を出している。 */}
      <div className="tabs">
        <button className={`tab-button ${tab === 'check' ? 'active' : ''}`} onClick={() => setTab('check')}>
          チェック
        </button>
        <button
          className={`tab-button ${tab === 'dictionary' ? 'active' : ''}`}
          onClick={() => setTab('dictionary')}
        >
          辞書
        </button>
      </div>

      {error && (
        <div className="rule-group-note" style={{ color: 'var(--shu)', padding: 'var(--space-3) var(--space-4) 0' }}>
          {error}
        </div>
      )}

      {tab === 'check' ? (
        <>
          <div className="summary-bar">
            <span className="summary-count">
              {scanned ? `${occurrences.length} 件見つかりました` : 'まだスキャンしていません'}
            </span>
            <button className="scan-button" onClick={runScan} disabled={scanning}>
              {scanning ? 'スキャン中…' : 'ドキュメントをスキャン'}
            </button>
          </div>
          <div className="scroll-area">
            {scanned && occurrences.length === 0 && <div className="empty-state">問題は見つかりませんでした。</div>}
            {!scanned && occurrences.length === 0 && (
              <div className="empty-state">
                「ドキュメントをスキャン」を押すと、登録した禁止語をチェックします。
              </div>
            )}
            <IssueList
              occurrences={occurrences}
              onSelect={handleSelect}
              onApplyOne={handleApplyOne}
              onApplyAll={handleApplyAll}
            />
          </div>
        </>
      ) : (
        <div className="scroll-area">
          <DictionaryEditor rules={rules} onChange={setRules} />
        </div>
      )}
    </div>
  )
}
