# CLAUDE.md — GetYoutubePlayCount

Google Apps Script で YouTube 動画の再生数・チャンネル内順位を記録するツール。

## 基本情報

- 言語: JavaScript (GAS)
- デプロイ: `npm run push`（GASのみ）/ `npm run deploy`（GAS + git push）
- mainブランチ: `main`（直push禁止）

## アーキテクチャ

GAS の6分制限を避けるため2トリガーに分割:
- `main()` — 毎時: 再生数取得・記録
- `updateAllCharts()` — 6時間ごと: グラフ・比較シート更新

## ファイル構成

- `config.gs` — CONFIG定数（ユーザー編集箇所）
- `main.gs` — エントリーポイントのみ
- `youtube.gs` — YouTube APIラッパー・ランク計算
- `video.gs` — 動画処理・シート操作・個別グラフ
- `comparison.gs` — 比較シート・グラフ
- `utils.gs` — ユーティリティ
- `admin.gs` — 手動実行用管理関数

## コード規約

- プライベート関数はトレイリングアンダースコア（例: `fetchVideoData_`）
- 重いシート操作後は `SpreadsheetApp.flush()`
- タイムアウトしやすい処理は `retryOnTimeout_()` でラップ
- GASはグローバルスコープ共有のため、ファイルをまたいだ呼び出し可能
