// ==========================================
// エントリーポイント
// ==========================================

/**
 * トリガーから呼び出すメイン処理（毎時）。
 * 全動画の再生数を取得・記録する。順位計算は updateChannelRanks() に分離。
 */
function main() {
  console.log('再生数取得を開始します...');

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  // トリガーの遅延（〜数分）を吸収するため正時に切り捨てる
  // 例: 13:01:05 に起動 → 13:00:00 として記録
  const now = new Date();
  now.setMinutes(0, 0, 0);

  const videoIds      = collectVideoIds_();
  const watchOnlySet  = new Set(CONFIG.WATCH_ONLY_VIDEO_IDS);
  const excludeSheets = new Set();
  console.log(`対象: ${videoIds.length} 本`);

  const videoDataMap = fetchAllVideoData_(videoIds);

  // 動画メタ情報（channelId等）を保存して updateChannelRanks() で再利用できるようにする
  saveVideoMetadataToProps_(videoDataMap);

  // updateChannelRanks() が保存した最新の rankMap を読み込んで動画シートのC列に書き込む
  const rankMap = loadRankMapFromProps_();

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
 * チャンネル内順位を更新する（1日2回トリガー推奨）。
 * main() とは独立して実行し、チャンネル全動画を1回だけ取得して順位を算出する。
 * 結果は Script Properties と「チャンネル内順位」シートに保存する。
 */
function updateChannelRanks() {
  console.log('チャンネル内順位の更新を開始します...');

  const metaMap = loadVideoMetadataFromProps_();
  if (Object.keys(metaMap).length === 0) {
    console.warn('動画メタデータが未保存です。先に main() を実行してください。');
    return;
  }

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const now     = new Date();
  now.setMinutes(0, 0, 0);

  const videoIds      = collectVideoIds_();
  const channelGroups = buildChannelGroups_(metaMap, videoIds);

  const { rankMap, viewCountMap } = computeRanksByChannelGroups_(channelGroups);
  if (Object.keys(rankMap).length === 0) {
    console.warn('順位計算結果が空でした。');
    return;
  }

  saveRankMapToProps_(rankMap);
  updateRankHistorySheet_(ss, rankMap, metaMap, viewCountMap, now);
  console.log('チャンネル内順位の更新完了。');
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
