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

/**
 * チャンネル内でライブと判定された動画の上位20件を再生数降順でログ出力する。
 * isLiveVideo_() の判定結果を確認するためのデバッグ用関数。
 * 対象チャンネルID は Script Properties の video_metadata から自動取得する。
 */
function debugLiveRanking() {
  const metaMap = loadVideoMetadataFromProps_();
  const channelIds = [...new Set(Object.values(metaMap).map(m => m.channelId).filter(Boolean))];

  channelIds.forEach(channelId => {
    console.log(`\n=== チャンネル ${channelId} ===`);
    const allIds  = fetchChannelVideoIds_(channelId);
    const stats   = fetchVideoStatsAndType_(allIds);

    const lives = Object.entries(stats)
      .filter(([, v]) => v.isLive)
      .sort(([, a], [, b]) => b.viewCount - a.viewCount)
      .slice(0, 20);

    console.log(`ライブ判定数: ${Object.values(stats).filter(v => v.isLive).length} / ${allIds.length}`);
    lives.forEach(([id, v], i) => {
      console.log(`  ${i + 1}位: ${id}  ${(v.viewCount / 10000).toFixed(1)}万回`);
    });
  });
}

/**
 * 比較シートのみを再生成する（動画データシートは変更しない）。
 * グラフやレイアウトを修正したい場合に使用する。
 */
function rebuildComparisonSheet() {
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
  const preserved   = sheets.filter(s =>  CONFIG.PRESERVE_SHEET_NAMES.includes(s.getName()));
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
