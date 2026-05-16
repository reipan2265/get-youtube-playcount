# Changelog

<!-- GAS version は deploy.yml の clasp version で自動作成されます。
     各リリースの GAS バージョン番号は GitHub Actions のジョブサマリーで確認できます。 -->

## [1.2.1] - 2026-05-17
### Changed
- `deploy.yml`: `fetch-depth: 0` を廃止し `fetch-depth: 1` + `fetch-tags: true` に変更（全履歴取得を解消）
- `auto-merge.yml`: `synchronize` トリガーを削除（push毎の不要な起動を抑制）
- `claude-review.yml`: timeout-minutes を 15 → 10 に短縮

## [1.2.0] - 2026-05-16
### Changed
- `CONFIG.PLAYLIST_ID`（単一文字列）を `CONFIG.PLAYLIST_IDS`（配列）に変更し複数プレイリストをサポート
- `fetchPlaylistVideoIds_()` が `PLAYLIST_IDS` の全プレイリストを順に取得するよう変更
- `_settings` シートのキーを `playlist_id` → `playlist_ids`（カンマ区切り複数対応）に変更
- `loadConfig_()` が `playlist_ids` キーを読んで `PLAYLIST_IDS` 配列を上書きするよう変更

## [1.0.5] - 2026-05-13
### Changed
- チャンネル内順位グラフをデータテーブル下の縦積みから行1の横一列並びに変更

## [1.0.4] - 2026-05-13
### Fixed
- `main()` 実行後にシートタブが投稿日昇順に並び替えられない問題を修正（`sortVideoSheetsByPublishDate_` を毎時呼び出すよう変更）
### Added
- 通常PRの自動マージワークフロー（`auto-merge.yml`）を追加

## [1.0.2] - 2026-05-01
### Changed
- `シート1` を `PRESERVE_SHEET_NAMES` から除外
### Added
- `deleteSheet1()` 管理関数を追加

## [1.0.1] - 2026-04-24
### Fixed
- シート並び替え関数 `sortVideoSheetsByPublishDate_` で `PRESERVE_SHEET_NAMES` の定義順が無視されるバグを修正
