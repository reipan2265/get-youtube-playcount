// ==========================================
// エントリーポイント
// ==========================================

/**
 * トリガーから呼び出すメイン処理。
 * 全動画の再生数を取得・記録し、比較シートを更新する。
 */
function main() {
  console.log('再生数取得を開始します...');

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  // トリガーの遅延（〜数分）を吸収するため正時に切り捨てる
  // 例: 13:01:05 に起動 → 13:00:00 として記録
  const now = new Date();
  now.setMinutes(0, 0, 0);

  const videoIds     = collectVideoIds_();
  const watchOnlySet = new Set(CONFIG.WATCH_ONLY_VIDEO_IDS);
  const excludeSheets = new Set();
  console.log(`対象: ${videoIds.length} 本`);

  const videoDataMap = fetchAllVideoData_(videoIds);

  // チャンネル内順位は12時間に1回だけ算出（API負荷軽減）
  // タイムスタンプは計算成功後に記録（失敗時は次回リトライさせる）
  let rankMap = {};
  if (shouldUpdateRank_()) {
    rankMap = computeChannelRankMap_(videoDataMap);
    if (Object.keys(rankMap).length > 0) {
      PropertiesService.getScriptProperties().setProperty('last_rank_update', String(Date.now()));
      updateRankHistorySheet_(ss, rankMap, videoDataMap, now);
    } else {
      console.warn('順位計算結果が空のためタイムスタンプを更新しません（次回リトライ）');
    }
  }

  videoIds.forEach((id, index) => {
    const sheetName = processVideo_(ss, id, index, videoIds.length, now, rankMap[id] ?? null, videoDataMap[id] ?? null);
    if (sheetName && watchOnlySet.has(id)) excludeSheets.add(sheetName);
    SpreadsheetApp.flush();
  });

  // 比較シートのテーブルを更新（チャートは updateAllCharts() で別途更新）
  updateComparisonTableOnly_(ss, excludeSheets);
  console.log('データ更新完了。');
}

/**
 * グラフ・比較シート・シート並び替えを更新する。
 * main() とは別トリガー（例: 6時間ごと）で実行することで実行時間超過を回避する。
 */
function updateAllCharts() {
  console.log('グラフ・比較シート更新を開始します...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const preserveSet = new Set(CONFIG.PRESERVE_SHEET_NAMES);
  ss.getSheets()
    .filter(sh => !preserveSet.has(sh.getName()) && !sh.getName().startsWith('_'))
    .forEach(sh => {
      updateIndividualChart_(sh);
      SpreadsheetApp.flush();
    });

  console.log('比較シートを更新します...');
  updateComparisonSheet_(ss);
  sortVideoSheetsByPublishDate_(ss);
  console.log('完了。');
}
