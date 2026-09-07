# Jun Office Linter

専門用語の誤変換（例: 疑似→擬似）を Word / PowerPoint のタスクペインから検出し、修正候補を提案するアドイン

## できること

- 自分で登録した「禁止語 → 正しい語」の辞書に基づき、文書内の該当箇所を検出
- 該当箇所をクリックしてドキュメント内にジャンプ（選択状態にする）
- 1件だけ修正 / そのルールの該当箇所をすべて修正、のどちらも可能
- 辞書はブラウザの `localStorage` に保存され、Word/PowerPoint どちらでも同じ辞書を共有（JSONでインポート/エクスポート可）
- 例外リスト（`exceptions`）で「この語を含む正しい複合語」を誤検出から除外可能

## インストール

このリポジトリを clone したり、Node/npm を用意したりする必要はありません。

1. [Releases](https://github.com/yuzukq/Jun-Office-Linter/releases/latest) から `manifest.xml` をダウンロード
2. Word または PowerPoint を開く
3. 「挿入」タブ →「アドイン」→「マイアドインのアップロード」から、ダウンロードした `manifest.xml` を選択

これだけで、リボンに「誤変換チェック」ボタンが追加されます。アドイン本体は GitHub Pages
(`https://yuzukq.github.io/Jun-Office-Linter/`) から配信され、`main` への push のたびに自動更新されるので、
一度サイドロードすればそれ以降は何もする必要がありません（ローカルでNode/Viteを動かし続ける必要はありません）。

コードを更新したのにタスクペインの表示が古いままの場合は、Word/PowerPoint を完全に終了（Cmd+Q）してから開き直してください。WebViewが以前の内容をキャッシュしていることがあります。マニフェスト自体（URLなど）を更新した場合は、Releasesから最新の `manifest.xml` を取り直してアップロードし直してください。

## 既知の制限

- Word: 本文・ヘッダー/フッターはスキャンしますが、脚注・文末脚注・floating テキストボックス内のテキストは対象外です。
- PowerPoint: スライド上のシェイプのテキストはスキャンしますが、スピーカーノートおよび2階層以上ネストしたグループ内シェイプは対象外です。
- 誤検出除外の例外リストは文字列一致ベースの簡易実装です。複雑な文脈判定はできません。

## 辞書の編集

タスクペイン上部の「辞書」タブから、禁止語・正しい語・メモを登録できます。研究テーマ全体で使う用語集なので、文書単位ではなくブラウザ（Officeアプリ）単位で保存されます。別のPCでも使いたい場合は「エクスポート」でJSONを書き出し、そちらで「インポート」してください。

デフォルトの辞書は `src/core/defaultRules.ts` にあります。

## 開発（コードを変更する場合）

```sh
npm install
npx office-addin-dev-certs install   # 初回のみ。ローカル用HTTPS証明書を信頼済みにする
```

コードを触ってすぐ確認したいときは、ローカルの Vite dev サーバーを使うと GitHub Pages へのデプロイを待たずに動作確認できます。

```sh
npm run start:word         # Word で試す場合
npm run start:powerpoint   # PowerPoint で試す場合
```

これらは `office-addin-debugging` 経由で、開発サーバー起動・マニフェストのサイドロード・対象アプリの起動をまとめて行います。ただしこのとき `manifest.xml` の `SourceLocation` は GitHub Pages の固定URLを指したままなので、ローカルの変更を確認するには一時的に `manifest.xml` 内の URL を `https://localhost:3000` に書き換えてから実行してください（コミットはしないこと）。

終了するときは:

```sh
npm run stop
```

現在の（GitHub Pagesを指す）`manifest.xml` を、開発サーバーなしで再サイドロードしたいだけの場合:

```sh
npm run reload:word
npm run reload:powerpoint
```

マニフェストの妥当性だけ確認したい場合:

```sh
npm run validate
```

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

`main` への push をトリガーに、GitHub Actions (`.github/workflows/deploy.yml`) が以下を自動で行います（Microsoft 365アカウントでのAppSource公開は不要です。個人利用のためリポジトリは public にしていますが、辞書データ自体はコードに含まれずローカルの `localStorage` にのみ保存されます）。

- `npm run build` で `dist/` を生成し、GitHub Pages に公開
- `manifest.xml` を [Releases](https://github.com/yuzukq/Jun-Office-Linter/releases/latest) の `latest` リリースに添付（毎回上書き更新）
