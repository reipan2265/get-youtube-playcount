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
  console.log('ランクキャッシュ（video_metadata / rank_map / last_rank_update）を削除しました。次回 main() 実行時に順位を再計算します。');
}

/** clearRankCache() の旧名エイリアス。 */
function resetRankTimer() { clearRankCache(); }


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
