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

## セットアップ

```sh
npm install
npx office-addin-dev-certs install   # 初回のみ。ローカル用HTTPS証明書を信頼済みにする
```

## 開発サーバーの起動とサイドロード

タスクペインは HTTPS で配信する必要があるため、`npm run dev` で起動する Vite サーバーは `https://localhost:3000` で待ち受けます。

Word で試す場合:

```sh
npm run start:word
```

PowerPoint で試す場合:

```sh
npm run start:powerpoint
```

これらは `office-addin-debugging` 経由で、開発サーバー起動・マニフェストのサイドロード・対象アプリの起動をまとめて行います（手動で `~/Library/Containers/.../Data/Documents/wef` にマニフェストをコピーするより確実です）。

終了するときは:

```sh
npm run stop
```

手動でマニフェストの妥当性だけ確認したい場合:

```sh
npm run validate
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

## 本番ビルド

```sh
npm run build
```

`dist/` に静的ファイルが出力されます。個人利用の範囲であれば、ビルド後のファイルを配置するホスティング先を用意し、`manifest.xml` 内の `https://localhost:3000` をそのURLに置き換えれば、開発サーバーなしで常時使えるようになります（Microsoft 365アカウントでのAppSource公開は不要です）。
