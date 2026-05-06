# 2026 RDSP Applicant Tracker

MS Formsから出力した応募者Excelを読み込み、以下を管理する最小実装です。

- 応募者情報（氏名・メール・生年月日・国籍）
- 書類提出状況（①パスポートコピー ②在籍証明書）
- 提出期限（共通設定 + 個別修正）
- 各応募者のOneDrive提出先リンク
- 進捗管理用Excelの再出力

このアプリは以下2種類のExcelを自動判別して読み込みます。

- MS Formsから直接ダウンロードした応募者Excel
  - 氏名は `First Name (e.g. John)`、`Middle Name (e.g. Alan)`、`Last Name (e.g. Smith)` を結合
  - メールは回答項目の `E-mail` を優先し、空欄の場合のみForms自動列の `メール` を使用
  - 生年月日は `Date of Birth (yyyy/MM/dd)` を読み込み、Excel日付シリアル値でも `yyyy/mm/dd` に変換
  - 国籍は `Nationality (Corresponds your passport / e.g. JAPAN)` を使用
- このアプリから出力した進捗管理Excel（再読み込みでパスポートコピー・在籍証明書の提出状況、提出期限、OneDriveリンクを復元）

## ローカル実行

```bash
npm install
npm run dev
```

## ビルド確認

```bash
npm run build
```

## StackBlitz

1. GitHubリポジトリをStackBlitzで `Import from GitHub`
2. 自動で `npm install` 後、`npm run dev` が起動

## Vercel

1. VercelでGitHub連携し本リポジトリをImport
2. Framework Preset: **Vite**
3. Build Command: `npm run build`
4. Output Directory: `dist`

