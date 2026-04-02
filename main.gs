// ==========================================
// エントリーポイント
// ==========================================

/**
 * トリガーから呼び出すメイン処理（毎時）。
 * 再生数を取得・記録し、12時間ごとに updateChannelRanks_() を呼び出す。
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

  // 動画メタ情報（channelId等）を保存（updateChannelRanks_() で再利用）
  saveVideoMetadataToProps_(videoDataMap);

  // 12時間ごとに順位を更新する
  if (shouldUpdateRank_()) {
    updateChannelRanks_(ss, now);
  }

  // 最新の rankMap を読み込んで動画シートのC列に書き込む
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
 * チャンネル内順位を計算して「チャンネル内順位」シートに記録する。
 * main() から 12 時間ごとに呼び出される。手動実行も可能。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @param {Date} [now]
 */
function updateChannelRanks(ss, now) {
  updateChannelRanks_(
    ss  ?? SpreadsheetApp.getActiveSpreadsheet(),
    now ?? (() => { const d = new Date(); d.setMinutes(0, 0, 0); return d; })(),
  );
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Date} now
 */
function updateChannelRanks_(ss, now) {
  console.log('チャンネル内順位の更新を開始します...');

  const metaMap = loadVideoMetadataFromProps_();
  if (Object.keys(metaMap).length === 0) {
    console.warn('動画メタデータが未保存です。先に main() を実行してください。');
    return;
  }

  const videoIds      = collectVideoIds_();
  const channelGroups = buildChannelGroups_(metaMap, videoIds);

  const { rankMap } = computeRanksByChannelGroups_(channelGroups);
  if (Object.keys(rankMap).length === 0) {
    console.warn('順位計算結果が空でした（次回リトライ）。');
    return;
  }

  saveRankMapToProps_(rankMap);
  PropertiesService.getScriptProperties().setProperty('last_rank_update', String(Date.now()));
  updateRankHistorySheet_(ss, rankMap, metaMap, now);
  console.log('チャンネル内順位の更新完了。');
}

/**
 * グラフ・比較シート・シート並び替えを更新する。
 * main() とは別トリガー（例: 6時間ごと）で実行することで実行時間超過を回避する。
 *
 * 実行順序：比較シート更新を先に行い、タイムアウトしても必ず反映されるようにする。
 * 個別グラフ更新は後回し（タイムアウトしても比較グラフへの影響なし）。
 */
function updateAllCharts() {
  console.log('グラフ・比較シート更新を開始します...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ① 比較シートを先に更新（最重要）
  console.log('比較シートを更新します...');
  updateComparisonSheet_(ss);
  sortVideoSheetsByPublishDate_(ss);

  // ② 順位推移グラフを更新
  console.log('順位推移グラフを更新します...');
  updateRankHistoryChart_(ss);
  SpreadsheetApp.flush();

  // ③ 個別グラフ更新（後回し・タイムアウトしても比較シートには影響しない）
  console.log('個別グラフを更新します...');
  const preserveSet = new Set(CONFIG.PRESERVE_SHEET_NAMES);
  ss.getSheets()
    .filter(sh => !preserveSet.has(sh.getName()) && !sh.getName().startsWith('_'))
    .forEach(sh => {
      updateIndividualChart_(sh);
      SpreadsheetApp.flush();
    });

  console.log('完了。');
}
