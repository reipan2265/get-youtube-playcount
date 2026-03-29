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

  // 全動画比較シートのシート名
  COMP_SHEET_NAME: '再生数比較',

  // 削除・リセット対象から除外するシート名
  PRESERVE_SHEET_NAMES: ['再生数比較', 'シート1', '_abs_helper', '_elapsed_helper', '_rank_helper'],

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
