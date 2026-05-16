// ==========================================
// 管理用ユーティリティ（手動実行）
// ==========================================

/**
 * updateChannelRanks() のキャッシュ（動画メタ情報・順位マップ）を削除する。
 * 次回 main() 実行時にメタ情報が再保存され、次回 updateChannelRanks() で順位が再計算される。
 */
function clearRankCache() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('video_metadata');
  props.deleteProperty('rank_map');
  props.deleteProperty('last_rank_update');
  props.deleteProperty('rank_sheet_col_map');
  console.log('ランクキャッシュ（video_metadata / rank_map / last_rank_update / rank_sheet_col_map）を削除しました。次回 main() 実行時に順位を再計算します。');
}

/** clearRankCache() の旧名エイリアス。 */
function resetRankTimer() { clearRankCache(); }


/**
 * 全動画シートの再生数非単調増加行（成長曲線の誤挿入等）を削除する。
 * グラフがジグザグになっているシートを修正する際に手動実行する。
 */
function fixNonMonotonicData() {
  console.log('非単調増加データのクリーンアップを開始します...');
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const preserveSet = new Set(CONFIG.PRESERVE_SHEET_NAMES);
  let totalRemoved = 0;

  ss.getSheets()
    .filter(sh => !preserveSet.has(sh.getName()) && !sh.getName().startsWith('_'))
    .forEach(sh => {
      totalRemoved += removeNonMonotonicRows_(sh);
      SpreadsheetApp.flush();
    });

  console.log(`完了。合計 ${totalRemoved} 行を削除しました。`);
}

/**
 * 動画シートをすべて削除してリセットする（PRESERVE_SHEET_NAMES は保持）。
 * 同時に成長曲線補完の実行済みフラグも削除する。
 * ⚠️ データが失われるため慎重に使用すること。
 */
function resetSheets() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  ss.getSheets()
    .filter(sh => !CONFIG.PRESERVE_SHEET_NAMES.includes(sh.getName()))
    .forEach(sh => {
      props.deleteProperty(`curve_filled_${sh.getName()}`);
      props.deleteProperty(`summary_fmt_${sh.getName()}`);
      ss.deleteSheet(sh);
    });
  console.log('動画シートをリセットしました');
}

function deleteSheet1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('シート1');
  if (sheet) { ss.deleteSheet(sheet); console.log('シート1を削除しました'); }
  else        { console.log('シート1は存在しません'); }
}

/**
 * CONFIG.LIVE_VIDEO_IDS に含まれる動画のチャンネル内順位（全動画混合）をログ出力する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 */
function debugLiveRanking() {
  const metaMap    = loadVideoMetadataFromProps_();
  const liveIds    = new Set(CONFIG.LIVE_VIDEO_IDS);
  const channelIds = [...new Set(Object.values(metaMap).map(m => m.channelId).filter(Boolean))];

  channelIds.forEach(channelId => {
    console.log(`\n=== チャンネル ${channelId} ===`);
    const allIds     = fetchChannelVideoIds_(channelId);
    const viewCounts = fetchViewCountsOnly_(allIds);
    const sorted     = allIds
      .filter(id => viewCounts[id] != null)
      .sort((a, b) => viewCounts[b] - viewCounts[a]);

    CONFIG.LIVE_VIDEO_IDS.forEach(id => {
      const idx = sorted.indexOf(id);
      const rank = idx >= 0 ? idx + 1 : null;
      console.log(`  ${id}: ${rank}位 / ${allIds.length}本中 (${((viewCounts[id] ?? 0) / 10000).toFixed(1)}万回)`);
    });
  });
}

/**
 * 比較シートのみを再生成する（動画データシートは変更しない）。
 * グラフやレイアウトを修正したい場合に使用する。
 */
function rebuildComparisonSheet() {
  loadConfig_();
  console.log('比較シートを再構築します...');
  updateComparisonSheet_(SpreadsheetApp.getActiveSpreadsheet());
  console.log('完了。');
}

/**
 * 動画シートを投稿日時の昇順（古い順が左、新しい順が右）に並び替える。
 * PRESERVE_SHEET_NAMES のシートは先頭（左側）に固定する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function sortVideoSheetsByPublishDate_(ss) {
  const sheets = ss.getSheets();
  const sheetMap    = new Map(sheets.map(s => [s.getName(), s]));
  const preserved   = CONFIG.PRESERVE_SHEET_NAMES.map(name => sheetMap.get(name)).filter(Boolean);
  const videoSheets = sheets.filter(s => !CONFIG.PRESERVE_SHEET_NAMES.includes(s.getName()) && !s.getName().startsWith('_'));

  videoSheets.sort((a, b) => {
    const dateA = a.getRange('A2').getValue();
    const dateB = b.getRange('A2').getValue();
    if (!(dateA instanceof Date)) return 1;
    if (!(dateB instanceof Date)) return -1;
    return dateA.getTime() - dateB.getTime();
  });

  [...preserved, ...videoSheets].forEach((sheet, index) => {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(index + 1);
  });
}

/**
 * `_settings` シートを作成し、現在の CONFIG 値を初期データとして書き込む。
 * すでに存在する場合はヘッダー行のみ確認し、データは上書きしない。
 * WebUI 連携の初期セットアップ時に手動実行する。
 */
function ensureSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('_settings');

  if (!sh) {
    sh = ss.insertSheet('_settings');
    console.log('_settings シートを作成しました。');
  }

  // ヘッダー行
  sh.getRange(1, 1, 1, 3).setValues([['key', 'value', 'memo']]);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold');

  // 既にデータ行がある場合は playlist_id → playlist_ids のキー移行のみ実施してスキップ
  if (sh.getLastRow() > 1) {
    const rows      = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    const oldKeyIdx = rows.findIndex(([k]) => k === 'playlist_id');
    const newKeyIdx = rows.findIndex(([k]) => k === 'playlist_ids');
    if (oldKeyIdx >= 0 && newKeyIdx < 0) {
      sh.getRange(oldKeyIdx + 2, 1).setValue('playlist_ids');
      sh.getRange(oldKeyIdx + 2, 3).setValue('プレイリスト ID（カンマ区切りで複数指定可）');
      console.log('_settings の playlist_id キーを playlist_ids に移行しました。');
    }
    console.log('_settings シートにはすでにデータが存在します。上書きをスキップしました。');
    return;
  }

  const rows = [
    ['playlist_ids',         CONFIG.PLAYLIST_IDS.join(','),                'プレイリスト ID（カンマ区切りで複数指定可）'],
    ['extra_video_ids',      CONFIG.EXTRA_VIDEO_IDS.join(','),             'プレイリスト外の追加動画 ID（カンマ区切り）'],
    ['watch_only_video_ids', CONFIG.WATCH_ONLY_VIDEO_IDS.join(','),        '推移のみ記録・比較シートに含めない動画 ID'],
    ['live_video_ids',       CONFIG.LIVE_VIDEO_IDS.join(','),              'ライブ配信アーカイブとして扱う動画 ID'],
    ['updated_at',           new Date().toISOString(),                     'WebUI からの最終更新日時'],
  ];
  sh.getRange(2, 1, rows.length, 3).setValues(rows);
  console.log('_settings シートに初期値を書き込みました。');
}
