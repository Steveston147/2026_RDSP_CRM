# 2026 RDSP Applicant Tracker

MS Formsから出力した応募者Excelを読み込み、以下を管理する最小実装です。

- 応募者情報（氏名・メール・生年月日・国籍）
- 書類提出状況（①パスポートコピー ②在籍証明書）
- 提出期限（共通設定 + 個別修正）
- 各応募者のOneDrive提出先リンク
- 進捗管理用Excelの再出力

このアプリは以下2種類のExcelを自動判別して読み込みます。

- MS Formsから直接ダウンロードした応募者Excel
- このアプリから出力した進捗管理Excel（再読み込みで提出状況や期限・リンクを復元）

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

