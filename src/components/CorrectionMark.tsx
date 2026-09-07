/** このアプリの署名的なビジュアル: 誤った語を校正記号の丸で囲み、矢印で修正候補を示す。 */
export function CorrectionMark({ wrong, correct }: { wrong: string; correct: string }) {
  return (
    <span className="correction-mark">
      <span className="wrong">{wrong}</span>
      <span className="arrow" aria-hidden="true">
        →
      </span>
      <span className="correct">{correct}</span>
    </span>
  )
}
