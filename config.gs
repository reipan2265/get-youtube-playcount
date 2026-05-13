// ==========================================
// 設定エリア（ここだけ編集してOK）
// ==========================================
const CONFIG = {
  // 対象プレイリストの ID（不要な場合は空文字）
  PLAYLIST_ID: 'PLriG7RRWaKk-YG8N7y4Fr8C15NJqnkLYG',

  // プレイリスト外で個別追加したい動画 ID
  EXTRA_VIDEO_IDS: ['Z_BpyttvaKI', 'WGrgo8-8XwY'],

  // 推移のみ記録（再生数比較シートには含めない）
  WATCH_ONLY_VIDEO_IDS: ['sd-4mwj1UDY'],

  // ライブ配信アーカイブとして扱う動画 ID（チャンネル内順位をライブ内で比較）
  // YouTube APIではプレミア公開と区別できないため、ここで明示指定する
  LIVE_VIDEO_IDS: ['sd-4mwj1UDY'],

  // 全動画比較シートのシート名
  COMP_SHEET_NAME: '再生数比較',

  // チャンネル内順位履歴シートのシート名
  RANK_SHEET_NAME: 'チャンネル内順位',

  // 削除・リセット対象から除外するシート名
  PRESERVE_SHEET_NAMES: ['再生数比較', 'チャンネル内順位', '_abs_helper', '_elapsed_helper', '_rank_helper'],

  // 比較グラフのサイズ（ピクセル）
  CHART: {
    WIDTH:  2210,
    HEIGHT:  850,
  },

  // 増加量サマリーの表示期間数（直近 + この数だけ前の期間を表示）
  SUMMARY_WINDOWS: 5,

  // テスト用: true にすると同一タイムスタンプのスキップを無視して強制書き込みする
  // 通常運用では必ず false にすること
  FORCE_WRITE: false,

  // データ間引き設定
  // keepEveryHours: null = 全件保持（トリガー間隔ごとに1件 = 実質1時間ごと）
  //                 数値 = その間隔（時間）ごとに1件保持
  SAMPLING: {
    MIN_ROWS_TO_SAMPLE: 10,
    RULES: [
      { maxDays:        30, keepEveryHours: null },  // 30日以内:  全件
      { maxDays:        90, keepEveryHours:    6 },  // ~90日:   6時間ごと
      { maxDays:       180, keepEveryHours:   12 },  // ~180日: 12時間ごと
      { maxDays:       365, keepEveryHours:   24 },  // ~365日:  1日ごと
      { maxDays: Infinity,  keepEveryHours:  168 },  // 365日超:   週1
    ],
  },
};

// ==========================================
// 定数
// ==========================================
const MS_PER_DAY  = 24 * 60 * 60 * 1000;
const MS_PER_HOUR =      60 * 60 * 1000;

// ==========================================
// 動的設定ローダー
// ==========================================

/**
 * `_settings` シートから動的設定を読み込み CONFIG オブジェクトを上書きする。
 * シートが存在しない場合は何もしない（config.gs のハードコード値が使われる）。
 * WebUI から設定を変更すると次回のトリガー実行で反映される。
 */
function loadConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('_settings');
  if (!sh || sh.getLastRow() < 2) return;

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const kv   = Object.fromEntries(rows.filter(([k]) => k !== '').map(([k, v]) => [String(k), v]));

  const split = s => String(s || '').split(',').map(x => x.trim()).filter(Boolean);

  if (kv['playlist_id']          != null) CONFIG.PLAYLIST_ID          = String(kv['playlist_id']);
  if (kv['extra_video_ids'])              CONFIG.EXTRA_VIDEO_IDS       = split(kv['extra_video_ids']);
  if (kv['watch_only_video_ids'])         CONFIG.WATCH_ONLY_VIDEO_IDS  = split(kv['watch_only_video_ids']);
  if (kv['live_video_ids'])               CONFIG.LIVE_VIDEO_IDS        = split(kv['live_video_ids']);
}
