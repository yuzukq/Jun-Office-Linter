/**
 * macOS由来の日本語テキストはNFD（濁点などが基底文字＋結合文字に分解された形）で
 * 渡ってくることがある。マッチング前にNFCへ正規化しておくことで、ルール文字列と
 * ドキュメント本文を正しく比較できるようにする。
 */
export function normalize(text: string): string {
  return text.normalize('NFC')
}
