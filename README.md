# Jun Office Linter

専門用語の誤変換（例: 疑似→擬似）を Word / PowerPoint のタスクペインから検出し、修正候補を提案するアドイン

## できること

- 自分で登録した「禁止語 → 正しい語」の辞書に基づき、文書内の該当箇所を検出
- 該当箇所をクリックしてドキュメント内にジャンプ（選択状態にする）
- 1件だけ修正 / そのルールの該当箇所をすべて修正、のどちらも可能
- 辞書はブラウザの `localStorage` に保存され、Word/PowerPoint どちらでも同じ辞書を共有（JSONでインポート/エクスポート可）
- 例外リスト（`exceptions`）で「この語を含む正しい複合語」を誤検出から除外可能

## 既知の制限

- Word: 本文・ヘッダー/フッターはスキャンしますが、脚注・文末脚注・floating テキストボックス内のテキストは対象外です。
- PowerPoint: スライド上のシェイプのテキストはスキャンしますが、スピーカーノートおよび2階層以上ネストしたグループ内シェイプは対象外です。
- 誤検出除外の例外リストは文字列一致ベースの簡易実装です。複雑な文脈判定はできません。

## 普段使い（GitHub Pages）

このアドインは完全に静的なファイルなので、`main` ブランチに push すると GitHub Actions (`.github/workflows/deploy.yml`) が自動でビルドして GitHub Pages
(`https://yuzukq.github.io/Jun-Office-Linter/`) に公開します。`manifest.xml` はこの固定URLを指しているので、**一度サイドロードすれば、それ以降はローカルでNodeやVite devサーバーを起動する必要はありません**。PowerPointの裏でNodeが動いている、という状態を維持しなくても常に使えます。

初回サイドロードは以下のどちらかで行います。

- Word/PowerPoint の「挿入」タブ → 「アドイン」→「マイアドインのアップロード」から `manifest.xml` を選択
- または `npm run reload:word` / `npm run reload:powerpoint`（開発サーバーは起動せず、現在の `manifest.xml` を再サイドロードするだけ）

`manifest.xml` を更新した後や、古い（localhostを指す）マニフェストが残っている場合は、上記のどちらかで読み直してください。Officeはマニフェストをローカルにコピーして保持するため、リポジトリ側を更新しただけでは自動的には反映されません。

辞書を編集してもコードの再デプロイは不要です（`localStorage` に保存されるだけなので）。コード自体（検出ロジックやUI）を変更したときだけ push すれば、数十秒後には最新版が全ホストに反映されます。

手動でマニフェストの妥当性だけ確認したい場合:

```sh
npm run validate
```

## 開発（コードを変更する場合）

コードを触ってすぐ確認したいときは、ローカルの Vite dev サーバーを使うと GitHub Pages へのデプロイを待たずに動作確認できます。

```sh
npm install
npx office-addin-dev-certs install   # 初回のみ。ローカル用HTTPS証明書を信頼済みにする
```

```sh
npm run start:word         # Word で試す場合
npm run start:powerpoint   # PowerPoint で試す場合
```

これらは `office-addin-debugging` 経由で、開発サーバー起動・マニフェストのサイドロード・対象アプリの起動をまとめて行います。ただしこのとき `manifest.xml` の `SourceLocation` は GitHub Pages の固定URLを指したままなので、ローカルの変更を確認するには一時的に `manifest.xml` 内の URL を `https://localhost:3000` に書き換えてから実行してください（コミットはしないこと）。

終了するときは:

```sh
npm run stop
```

## 辞書の編集

タスクペイン上部の「辞書」タブから、禁止語・正しい語・メモを登録できます。研究テーマ全体で使う用語集なので、文書単位ではなくブラウザ（Officeアプリ）単位で保存されます。別のPCでも使いたい場合は「エクスポート」でJSONを書き出し、そちらで「インポート」してください。

デフォルトの辞書は `src/core/defaultRules.ts` にあります。

## プロジェクト構成

```
src/
  core/        辞書の型・マッチングロジック・localStorage永続化（ホスト非依存）
  adapters/    Word / PowerPoint それぞれの Office.js 連携（scan/select/apply）
  components/  タスクペインUI
  styles/      デザイントークン・Officeテーマ同期
manifest.xml   Word (Document) / PowerPoint (Presentation) 両方を宣言
```

`core` は Word/PowerPoint どちらにも依存しない純粋なロジックなので、将来的に `.docx`/`.pptx` をオフラインで一括チェックするCLIを作る場合もここを再利用できます。

## デプロイの仕組み

`npm run build` で `dist/` に静的ファイルが出力され、`main` への push をトリガーに GitHub Actions がこれをビルドして GitHub Pages に公開します（Microsoft 365アカウントでのAppSource公開は不要です。個人利用のためリポジトリは public にしていますが、辞書データ自体はコードに含まれずローカルの `localStorage` にのみ保存されます）。
