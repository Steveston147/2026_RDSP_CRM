# 2026 RDSP Applicant Tracker

MS Formsから出力した応募者Excelを直接読み込み、以下を管理する最小実装です。

- 応募者情報（氏名・メール・生年月日・国籍）
- 書類提出状況（①パスポートコピー ②在籍証明書）
- 提出期限（共通設定 + 個別修正）
- 各応募者のOneDrive提出先リンク
- 進捗管理用Excelの再出力と再読み込み

## Excelインポート仕様

MS Formsから出力したExcelは、以下の列を読み込みます。

- 氏名: `First Name` / `Middle Name` / `Last Name` を結合します。Middle Nameが空欄の場合も余分なスペースは入りません。
- メール: `E-mail` を優先し、空欄の場合のみ `メール` を使用します。
- 生年月日: `Date of Birth (yyyy/MM/dd)` を読み込みます。Excelシリアル日付は `yyyy/mm/dd` 表示に変換します。
- 国籍: `Nationality (Corresponds your passport / e.g. JAPAN)` を読み込みます。

MS Forms Excelの初回インポートでは、`パスポートコピー`、`在籍証明書`、`提出期限`、`OneDriveリンク` は空の進捗として開始します。

このアプリから出力した進捗Excelを再度インポートした場合は、`パスポートコピー`、`在籍証明書`、`提出期限`、`OneDriveリンク` を復元します。

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

