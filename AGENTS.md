# Repository Guidelines

## Project Structure & Module Organization
このリポジトリは Google Apps Script（GAS）で構成された単一スクリプト型のプロジェクトです。主要な実装は `コード.js` にあり、`onOpen()` によるカスタムメニュー作成、カレンダー書き込み、祝日追加・削除、毎週予定の一括操作をまとめて管理しています。`appsscript.json` はマニフェストで、タイムゾーンと V8 ランタイムを定義します。`.clasp.json` は `clasp` の接続設定です。テスト用ディレクトリやビルド成果物はありません。

## Build, Test, and Development Commands
- `clasp login` : Google Apps Script CLI を認証します。
- `clasp status` : ローカルとリモートの差分を確認します。
- `clasp pull` : エディタ側の変更を取り込んで競合を防ぎます。
- `clasp push` : `コード.js` と `appsscript.json` を Apps Script プロジェクトへ反映します。
- `clasp open` : Apps Script エディタを開きます。

ビルド工程はありません。反映後は対象スプレッドシートを再読み込みし、`年間行事予定` メニューから動作確認してください。

## Coding Style & Naming Conventions
既存コードに合わせて 2 スペースインデント、セミコロンあり、`var` ベースで統一してください。GAS のエントリーポイントは `function` 宣言を使い、関数名は `writeScheduleToCalendarSpecificMonth` のような camelCase を採用します。ユーザー向けメッセージ、コメント、メニュー名は日本語で維持してください。予定文字列は `会議<10:00-12:00>`、複数予定は `授業参観,職員会議` の形式を前提に実装します。

## Testing Guidelines
自動テストは未整備です。変更時はコピーしたスプレッドシートとテスト用カレンダーで手動検証してください。少なくとも、全期間登録、前期・後期登録、月別登録、祝日追加・削除、毎週予定追加・削除を確認します。破壊的処理を含むため、本番カレンダーで直接検証しないでください。

## Commit & Pull Request Guidelines
公開履歴では、`READMEを更新し、特定の月だけを選択して書き込み可能な機能を追加` のように、変更内容をそのまま説明する日本語コミットメッセージが使われています。コミットは 1 変更点 1 目的で分け、機能追加・修正・ドキュメント更新を混在させすぎないでください。PR には変更したメニュー項目、影響するセル範囲（例: `A1`, `E1`, `A3:X33`）、確認手順、UI 変更がある場合はスクリーンショットを添えてください。

## Security & Configuration Tips
`scriptId`、カレンダー ID、共有用でないスプレッドシート URL は不用意に差し替えたり公開したりしないでください。`clasp push` 前に `clasp pull` を実行し、別環境の変更を上書きしない運用を推奨します。
