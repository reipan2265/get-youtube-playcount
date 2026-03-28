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

// ==========================================
// 1. エントリーポイント
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

  // 全動画の詳細を一括取得してチャンネル内順位を算出
  const videoDataMap = fetchAllVideoData_(videoIds);
  const rankMap      = computeChannelRankMap_(videoDataMap);

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

  // 個別動画グラフを更新
  const preserveSet = new Set(CONFIG.PRESERVE_SHEET_NAMES);
  ss.getSheets()
    .filter(sh => !preserveSet.has(sh.getName()))
    .forEach(sh => {
      updateIndividualChart_(sh);
      SpreadsheetApp.flush();
    });

  // 比較シートを更新
  console.log('比較シートを更新します...');
  updateComparisonSheet_(ss);
  sortVideoSheetsByPublishDate_(ss);
  console.log('完了。');
}

// ==========================================
// 2. 動画ごとの処理
// ==========================================
/**
 * 1本の動画を処理する（取得 → 記録 → 補完 → 間引き → グラフ更新）。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string}      id              YouTube 動画 ID
 * @param {number}      index           現在のインデックス（ログ用）
 * @param {number}      total           総件数（ログ用）
 * @param {Date}        now             実行日時
 * @param {number|null} rank            チャンネル内再生数順位（1が最多）
 * @param {object|null} preloadedVideo  fetchAllVideoData_ で取得済みの動画オブジェクト
 */
function processVideo_(ss, id, index, total, now, rank, preloadedVideo) {
  try {
    const video = preloadedVideo ?? fetchVideoData_(id);
    if (!video) {
      console.warn(`[${index + 1}/${total}] 動画が見つかりません (id: ${id})`);
      return null;
    }

    const { fullTitle, viewCount, publishedAt, channelId, channelTitle } = parseVideoData_(video);
    const sheetName = buildSheetName_(fullTitle, id);
    const sheet     = getOrCreateVideoSheet_(ss, sheetName, fullTitle, publishedAt, channelId, channelTitle);

    // 同一タイムスタンプが既に存在する場合はスキップ（冪等性の担保）
    // CONFIG.FORCE_WRITE = true のときはスキップしない
    const lastRow = sheet.getLastRow();
    if (!CONFIG.FORCE_WRITE && lastRow >= 4) {
      const existing = sheet.getRange(4, 1, lastRow - 3, 1).getValues();
      if (existing.some(r => r[0] instanceof Date && r[0].getTime() === now.getTime())) {
        console.log(`[${index + 1}/${total}] ${sheetName}: スキップ（同一タイムスタンプ既存）`);
        return sheetName;
      }
    }

    sheet.appendRow([now, viewCount, rank ?? '']);
    console.log(`[${index + 1}/${total}] ${sheetName}: ${viewCount.toLocaleString()} 回 (順位: ${rank ?? '?'})`);

    fillInitialGrowthCurve_(sheet, publishedAt);
    runSampling_(sheet, publishedAt);
    // グラフ更新は updateAllCharts() に分離（実行時間超過対策）
    sortVideoSheetDescending_(sheet);
    // ソート後のデータを1回読んで渡す（updateGrowthSummary_ 内の重複読み込みを省く）
    const lastRow_ = sheet.getLastRow();
    const allData_ = lastRow_ >= 4 ? sheet.getRange(4, 1, lastRow_ - 3, 2).getValues() : [];
    updateGrowthSummary_(sheet, viewCount, now, allData_);

    return sheetName;
  } catch (e) {
    console.error(`[${index + 1}/${total}] エラー (id: ${id}): ${e.message}\n${e.stack}`);
    return null;
  }
}

// ==========================================
// 3. YouTube API ラッパー
// ==========================================
/**
 * YouTube Data API から動画情報を取得する。
 * @param {string} id  YouTube 動画 ID
 * @returns {object|null}  API レスポンスの items[0]、見つからない場合は null
 */
function fetchVideoData_(id) {
  const res = YouTube.Videos.list('snippet,statistics', { id });
  return res.items?.[0] ?? null;
}

/**
 * API レスポンスから必要フィールドを抽出する。
 * @param {object} video  fetchVideoData_ の戻り値
 * @returns {{ fullTitle: string, viewCount: number, publishedAt: Date }}
 */
function parseVideoData_(video) {
  return {
    fullTitle   : video.snippet.title,
    viewCount   : Number(video.statistics.viewCount),
    publishedAt : new Date(video.snippet.publishedAt),
    channelId   : video.snippet.channelId,
    channelTitle: video.snippet.channelTitle,
  };
}

/**
 * 複数の動画 ID を一括取得して { videoId: item } のマップを返す。
 * YouTube Data API は1リクエストで最大50件対応。
 * @param {string[]} videoIds
 * @returns {Object<string, object>}
 */
function fetchAllVideoData_(videoIds) {
  const result = {};
  for (let i = 0; i < videoIds.length; i += 50) {
    const batch = videoIds.slice(i, i + 50);
    try {
      const res = YouTube.Videos.list('snippet,statistics', { id: batch.join(',') });
      res.items?.forEach(item => { result[item.id] = item; });
    } catch (e) {
      console.warn(`動画情報一括取得失敗 (offset ${i}): ${e.message}`);
    }
  }
  return result;
}

/**
 * チャンネルのアップロードプレイリストから全動画 ID を取得する。
 * アップロードプレイリスト ID = チャンネル ID の先頭 "UC" を "UU" に変換したもの。
 * @param {string} channelId  "UC..." 形式のチャンネル ID
 * @returns {string[]}
 */
function fetchChannelVideoIds_(channelId) {
  const uploadsPlaylistId = 'UU' + channelId.slice(2);
  const ids = [];
  let pageToken = '';
  try {
    do {
      const res = YouTube.PlaylistItems.list('snippet', {
        playlistId: uploadsPlaylistId,
        maxResults: 50,
        pageToken,
      });
      res.items?.forEach(item => ids.push(item.snippet.resourceId.videoId));
      pageToken = res.nextPageToken ?? '';
    } while (pageToken);
  } catch (e) {
    console.warn(`チャンネル動画一覧取得失敗 (${channelId}): ${e.message}`);
  }
  return ids;
}

/**
 * 動画 ID リストの再生数のみを一括取得して { videoId: viewCount } を返す。
 * @param {string[]} videoIds
 * @returns {Object<string, number>}
 */
function fetchViewCountsOnly_(videoIds) {
  const result = {};
  for (let i = 0; i < videoIds.length; i += 50) {
    const batch = videoIds.slice(i, i + 50);
    try {
      const res = YouTube.Videos.list('statistics', { id: batch.join(',') });
      res.items?.forEach(item => { result[item.id] = Number(item.statistics.viewCount); });
    } catch (e) {
      console.warn(`再生数一括取得失敗 (offset ${i}): ${e.message}`);
    }
  }
  return result;
}

/**
 * 追跡動画ごとに「チャンネル全動画の中での再生数順位」を返す。
 * チャンネルのアップロードプレイリストから全動画を取得し、再生数でランクを付与する。
 * @param {Object<string, object>} videoDataMap  { videoId: YouTubeVideoItem }
 * @returns {Object<string, number>}  { videoId: rank }
 */
function computeChannelRankMap_(videoDataMap) {
  // 追跡動画をチャンネルごとにグループ化
  const channelGroups = {}; // { channelId: [videoId, ...] }
  Object.entries(videoDataMap).forEach(([id, item]) => {
    const cid = item.snippet.channelId;
    if (!channelGroups[cid]) channelGroups[cid] = [];
    channelGroups[cid].push(id);
  });

  const rankMap = {};

  Object.entries(channelGroups).forEach(([channelId, trackedIds]) => {
    console.log(`チャンネル ${channelId} の全動画を取得中...`);
    const allIds = fetchChannelVideoIds_(channelId);
    console.log(`チャンネル全動画: ${allIds.length} 本`);

    const viewCounts = fetchViewCountsOnly_(allIds);

    // 再生数降順でソートしてランクを決定（ランク外の動画は null）
    const sorted = allIds
      .filter(id => viewCounts[id] != null)
      .sort((a, b) => viewCounts[b] - viewCounts[a]);

    trackedIds.forEach(id => {
      const idx = sorted.indexOf(id);
      rankMap[id] = idx >= 0 ? idx + 1 : null;
    });

    console.log(`ランク算出完了: ${trackedIds.map(id => `${id}=${rankMap[id]}`).join(', ')}`);
  });

  return rankMap;
}

/**
 * プレイリスト内の全動画 ID を取得する（ページネーション対応）。
 * @returns {string[]}
 */
function fetchPlaylistVideoIds_() {
  if (!CONFIG.PLAYLIST_ID) return [];

  const ids = [];
  let pageToken = '';

  try {
    do {
      const res = YouTube.PlaylistItems.list('snippet', {
        playlistId: CONFIG.PLAYLIST_ID,
        maxResults: 50,
        pageToken,
      });
      res.items.forEach(item => ids.push(item.snippet.resourceId.videoId));
      pageToken = res.nextPageToken ?? '';
    } while (pageToken);
  } catch (e) {
    console.warn(`プレイリスト取得失敗: ${e.message}`);
  }

  return ids;
}

/**
 * EXTRA_VIDEO_IDS, WATCH_ONLY_VIDEO_IDS, プレイリストを合わせて重複排除した動画 ID リストを返す。
 * @returns {string[]}
 */
function collectVideoIds_() {
  const playlistIds = fetchPlaylistVideoIds_();
  console.log(`プレイリストから ${playlistIds.length} 本を検出`);
  return [...new Set([...CONFIG.EXTRA_VIDEO_IDS, ...CONFIG.WATCH_ONLY_VIDEO_IDS, ...playlistIds])];
}

/**
 * WATCH_ONLY_VIDEO_IDS の動画IDから対応するシート名を解決する。
 * rebuildComparisonSheet など main() 経由でない呼び出し用。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {Set<string>}
 */
function resolveWatchOnlySheetNames_(ss) {
  const exclude = new Set();
  CONFIG.WATCH_ONLY_VIDEO_IDS.forEach(id => {
    try {
      const video = fetchVideoData_(id);
      if (video) {
        const title = video.snippet.title;
        exclude.add(buildSheetName_(title, id));
      }
    } catch (_) { /* API エラーは無視 */ }
  });
  return exclude;
}

// ==========================================
// 4. シート操作
// ==========================================
/**
 * 動画タイトルからシート名を生成する。
 * Google Sheets の制約: 31 文字以内、バックスラッシュ等の特殊文字禁止。
 * @param {string} fullTitle  動画タイトル
 * @param {string} fallback   タイトルが空の場合に使う動画 ID
 * @returns {string}
 */
function buildSheetName_(fullTitle, fallback) {
  const INVALID_CHARS = /[\\\/\?\*\:\[\]]/g;
  const MAX_LEN = 31;

  let name = fullTitle.replace(/【.*?】/g, '').trim().replace(INVALID_CHARS, '_');
  if (!name) return fallback;

  if (name.length > MAX_LEN) {
    const half = Math.floor((MAX_LEN - 1) / 2);
    name = name.substring(0, half) + '…' + name.substring(name.length - half);
  }

  return name;
}

/**
 * 動画シートを取得する。存在しなければ新規作成してヘッダーを書き込む。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @param {string} fullTitle
 * @param {Date}   publishedAt
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateVideoSheet_(ss, sheetName, fullTitle, publishedAt, channelId, channelTitle) {
  let sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    if (!sheet.getRange('A2').getValue()) {
      sheet.getRange('A2').setValue(publishedAt);
    }
    // 既存シートにチャンネル情報が未記入の場合のみ補完
    if (channelId && !sheet.getRange('B2').getValue()) {
      sheet.getRange('B2').setValue(channelId);
      sheet.getRange('C2').setValue(channelTitle || '');
    }
    return sheet;
  }

  console.log(`シートを作成: ${sheetName}`);
  sheet = ss.insertSheet(sheetName);

  sheet.getRange('A1').setValue(fullTitle).setFontWeight('bold');
  sheet.getRange('A2').setValue(publishedAt);
  sheet.getRange('B2').setValue(channelId   || '');
  sheet.getRange('C2').setValue(channelTitle || '');
  sheet.getRange('A3:C3').setValues([['日時', '再生数', '順位']]).setBackground('#eeeeee');
  sheet.setFrozenRows(3);
  sheet.appendRow([publishedAt, 0, '']); // 投稿日時点の起点レコード（再生数 = 0、順位は未計測）

  return sheet;
}

/**
 * 動画シートのデータ行（row4以降）をタイムスタンプ降順に並び替える。
 * 最新データが常に行4（先頭）になる。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function sortVideoSheetDescending_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 4) return;
  sheet.getRange(4, 1, lastRow - 3, 3).sort({ column: 1, ascending: false });
}

// ==========================================
// 5. データ間引き（サンプリング）
// ==========================================
/**
 * CONFIG.SAMPLING.RULES に従い、古いデータを間引いてシートの肥大化を抑制する。
 * バケツ方式で均等に間引くため、記録タイミングのズレに強い。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Date} publishedAt
 */
function runSampling_(sheet, publishedAt) {
  const dataCount = sheet.getLastRow() - 3;
  if (dataCount < CONFIG.SAMPLING.MIN_ROWS_TO_SAMPLE) return;

  const values     = sheet.getRange(4, 1, dataCount, 3).getValues();
  const seenBucket = new Set();

  const keepRows = values.filter((row, index) => {
    if (index === 0 || index === values.length - 1) return true; // 先頭・末尾は必ず保持
    if (!(row[0] instanceof Date)) return false;

    const ageDays = (row[0].getTime() - publishedAt.getTime()) / MS_PER_DAY;
    const rule    = CONFIG.SAMPLING.RULES.find(r => ageDays <= r.maxDays);

    if (!rule || rule.keepEveryHours === null) return true;

    const ageHours  = (row[0].getTime() - publishedAt.getTime()) / MS_PER_HOUR;
    const bucketKey = `${rule.keepEveryHours}_${Math.floor(ageHours / rule.keepEveryHours)}`;

    if (seenBucket.has(bucketKey)) return false;
    seenBucket.add(bucketKey);
    return true;
  });

  if (keepRows.length < CONFIG.SAMPLING.MIN_ROWS_TO_SAMPLE) {
    console.log(`間引きスキップ [${sheet.getName()}]: 間引き後 ${keepRows.length} 行 < 最低 ${CONFIG.SAMPLING.MIN_ROWS_TO_SAMPLE} 行`);
    return;
  }

  if (keepRows.length < values.length) {
    console.log(`間引き [${sheet.getName()}]: ${values.length} → ${keepRows.length} 行`);
    sheet.getRange(4, 1, sheet.getLastRow() - 3, 3).clearContent();
    sheet.getRange(4, 1, keepRows.length, 3).setValues(keepRows);
  }
}

// ==========================================
// 6. 個別動画グラフ
// ==========================================
/**
 * 各動画シートに再生数推移の折れ線グラフを作成または更新する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function updateIndividualChart_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  const dataRange = sheet.getRange(`A3:B${lastRow}`);
  const charts    = sheet.getCharts();
  const title     = sheet.getRange('A1').getValue() || sheet.getName();

  const builder = (charts.length > 0 ? charts[0].modify() : sheet.newChart())
    .asLineChart()
    .clearRanges()
    .addRange(dataRange)
    .setPosition(10, 4, 0, 0)
    .setOption('title', title)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
    .setOption('vAxis', { format: '#,###' })
    .setOption('pointSize', 2)
    .setOption('lineWidth', 2);

  if (charts.length > 0) {
    sheet.updateChart(builder.build());
  } else {
    sheet.insertChart(builder.build());
  }
}

// ==========================================
// 7. 比較シートの生成
// ==========================================
/**
 * 比較シートのテーブルデータのみを更新する（チャートは変更しない）。
 * main() から毎時呼び出し、最新再生数・増加数を常に最新に保つ。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Set<string>} excludeSheets  除外するシート名のセット
 */
function updateComparisonTableOnly_(ss, excludeSheets) {
  if (!excludeSheets) excludeSheets = resolveWatchOnlySheetNames_(ss);

  const compSheet = ss.getSheetByName(CONFIG.COMP_SHEET_NAME);
  if (!compSheet) return;

  const videoSheets = ss.getSheets().filter(s =>
    !CONFIG.PRESERVE_SHEET_NAMES.includes(s.getName()) && !excludeSheets.has(s.getName())
  );

  const { dataMap, publishDateMap, sortedTimestamps } = aggregateVideoData_(videoSheets);
  if (sortedTimestamps.length === 0) return;

  const { tableValues } = buildComparisonTable_(videoSheets, dataMap, publishDateMap, sortedTimestamps);

  const rows = tableValues.length;
  const cols = tableValues[0].length;

  // チャートはそのままでテーブル範囲のみ上書き
  compSheet.getRange(1, 1, rows, cols).setValues(tableValues);
  compSheet.getRange(2, 2, rows - 1, 1).setNumberFormat('#,###').setFontWeight('bold');
  compSheet.getRange(2, 3, rows - 1, 1).setNumberFormat('0');
  SpreadsheetApp.flush();
  console.log('比較シートのテーブルを更新しました。');
}

/**
 * 全動画の再生数推移を1シートに集約し、比較グラフを2種類生成する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function updateComparisonSheet_(ss, excludeSheets) {
  SpreadsheetApp.flush();
  Utilities.sleep(2000);

  if (!excludeSheets) excludeSheets = resolveWatchOnlySheetNames_(ss);

  const compSheet   = ss.getSheetByName(CONFIG.COMP_SHEET_NAME) || ss.insertSheet(CONFIG.COMP_SHEET_NAME, 0);
  const videoSheets = ss.getSheets().filter(s =>
    !CONFIG.PRESERVE_SHEET_NAMES.includes(s.getName()) && !excludeSheets.has(s.getName())
  );

  const { dataMap, publishDateMap, elapsedMaps, sortedTimestamps } = aggregateVideoData_(videoSheets);
  if (sortedTimestamps.length === 0) {
    console.warn('比較シートに書き込むデータがありません');
    return;
  }

  const { tableValues, sortedNames, sortedTitles } = buildComparisonTable_(videoSheets, dataMap, publishDateMap, sortedTimestamps);

  renderComparisonSheet_(compSheet, tableValues);

  const absHelperSheet = getOrCreateHelperSheet_(ss, '_abs_helper');
  const absHelper = buildAbsoluteTimeHelperTable_(absHelperSheet, sortedNames, sortedTitles, dataMap, sortedTimestamps);
  buildComparisonChart_(compSheet, absHelperSheet, absHelper, tableValues.length);

  const elapsedHelperSheet = getOrCreateHelperSheet_(ss, '_elapsed_helper');
  buildElapsedDaysChart_(compSheet, elapsedHelperSheet, videoSheets, publishDateMap, elapsedMaps, tableValues.length);

  const rankHelperSheet = getOrCreateHelperSheet_(ss, '_rank_helper');
  buildChannelRankCharts_(ss, compSheet, rankHelperSheet, videoSheets, tableValues.length);
}

/**
 * 全動画シートのデータを集約して返す。
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} videoSheets
 * @returns {{ dataMap: object, publishDateMap: object, sortedTimestamps: string[] }}
 */
function aggregateVideoData_(videoSheets) {
  const dataMap        = {}; // { timestamp: { sheetName: viewCount } }
  const publishDateMap = {}; // { sheetName: Date }
  const elapsedMaps    = {}; // { sheetName: Map<elapsedHours, viewCount> } （経過日数グラフ用）
  const allTimestamps  = new Set();

  videoSheets.forEach(sh => {
    const name       = sh.getName();
    const pubDateVal = sh.getRange('A2').getValue();
    if (!pubDateVal) return;

    const pubDate = new Date(pubDateVal);
    publishDateMap[name] = pubDate;

    const pubTs = formatTimestamp_(pubDate);
    allTimestamps.add(pubTs);
    setNestedValue_(dataMap, pubTs, name, 0);

    const elapsed = new Map();
    elapsed.set(0, 0); // 投稿日 = 起点 0 再生
    elapsedMaps[name] = elapsed;

    const lastRow = sh.getLastRow();
    if (lastRow < 4) return;

    sh.getRange(4, 1, lastRow - 3, 2).getValues().forEach(row => {
      if (!(row[0] instanceof Date)) return;
      const ts = formatTimestamp_(row[0]);
      allTimestamps.add(ts);
      setNestedValue_(dataMap, ts, name, row[1]);

      // 経過日数マップも同時に構築（buildElapsedDaysChart_ の重複読み込みを省く）
      const elapsedMs = row[0].getTime() - pubDate.getTime();
      if (elapsedMs >= 0) {
        elapsed.set(getRoundedElapsedHours_(elapsedMs), row[1]);
      }
    });
  });

  return {
    dataMap,
    publishDateMap,
    elapsedMaps,
    sortedTimestamps: [...allTimestamps].sort().reverse(), // 最新が左になるよう降順
  };
}

/**
 * 比較シート用のテーブルデータ（2列: 動画名, 1日平均再生数）を構築する。
 * 動画行は最新再生数の降順でソートする。
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} videoSheets
 * @param {object}   dataMap
 * @param {object}   publishDateMap
 * @param {string[]} sortedTimestamps
 * @returns {{ tableValues: any[][], sortedNames: string[] }}
 */
function buildComparisonTable_(videoSheets, dataMap, publishDateMap, sortedTimestamps) {
  const now   = new Date();
  const nowMs = now.getTime();

  const videoRows = videoSheets
    .map(sh => {
      const name    = sh.getName();
      const pubDate = publishDateMap[name];
      if (!pubDate) return null;

      // dataMap から降順の [Date, viewCount] 配列を再構築（シートの再読み込みを省く）
      const allData = sortedTimestamps
        .filter(ts => dataMap[ts]?.[name] != null)
        .map(ts => [new Date(ts), dataMap[ts][name]]);

      const lastVal          = allData.length > 0 ? allData[0][1] : 0;
      const daysSincePublish = Math.floor((nowMs - pubDate.getTime()) / MS_PER_DAY);

      const fmt = (fromMs) =>
        formatIncreaseWithRate_(calcIncrease_(allData, lastVal, nowMs, fromMs, 0), lastVal);

      const fullTitle    = sh.getRange('A1').getValue() || name;
      const escapedName  = name.replace(/'/g, "''");
      const titleFormula = `='${escapedName}'!$A$1`;

      return {
        row: [
          titleFormula,
          lastVal,
          daysSincePublish,
          fmt(MS_PER_DAY),
          fmt(7 * MS_PER_DAY),
          fmt(30 * MS_PER_DAY),
        ],
        lastVal,
        name,
        fullTitle,
      };
    })
    .filter(Boolean)
    .sort((a, b) => b.lastVal - a.lastVal);

  const headerRow = ['動画名', '最新再生数', '経過日数', '24h増加', '7日増加', '30日増加'];
  return {
    tableValues: [headerRow, ...videoRows.map(r => r.row)],
    sortedNames:  videoRows.map(r => r.name),
    sortedTitles: videoRows.map(r => r.fullTitle),
  };
}

/**
 * 比較シートにテーブル（2列: 動画名, 1日平均再生数）を書き込む。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} compSheet
 * @param {any[][]} tableValues
 */
function renderComparisonSheet_(compSheet, tableValues) {
  compSheet.getCharts().forEach(c => compSheet.removeChart(c));
  compSheet.clear();

  const rows = tableValues.length;
  const cols = tableValues[0].length;

  // 余剰列を削除してセル数を最小化（ヘルパーテーブルは別シートに移動済み）
  const maxCols = compSheet.getMaxColumns();
  if (maxCols > cols) {
    compSheet.deleteColumns(cols + 1, maxCols - cols);
  }

  compSheet.getRange(1, 1, rows, cols).setValues(tableValues);
  // ヘッダー行
  compSheet.getRange(1, 1, 1, cols).setBackground('#eeeeee').setFontWeight('bold');
  // B列（最新再生数）: 数値フォーマット + 太字
  compSheet.getRange(2, 2, rows - 1, 1).setNumberFormat('#,###').setFontWeight('bold');
  // C列（経過日数）: 整数フォーマット
  compSheet.getRange(2, 3, rows - 1, 1).setNumberFormat('0');
  compSheet.setFrozenRows(1);
}

/**
 * ヘルパーシートを取得または新規作成する。既存の場合はクリアしてサイズを最小化する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name  シート名
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateHelperSheet_(ss, name) {
  return retryOnTimeout_(() => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.hideSheet();
    } else {
      sheet.getCharts().forEach(c => sheet.removeChart(c));
      sheet.clear();
      // 余剰行・列を削除してセル数を最小化
      if (sheet.getMaxColumns() > 1) {
        sheet.deleteColumns(2, sheet.getMaxColumns() - 1);
      }
      if (sheet.getMaxRows() > 1) {
        sheet.deleteRows(2, sheet.getMaxRows() - 1);
      }
    }
    return sheet;
  });
}

/**
 * 専用ヘルパーシートに絶対日時テーブルを書き込む。グラフのデータ源として使用する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} helperSheet  書き込み先ヘルパーシート
 * @param {string[]} sortedNames   動画名（ソート済み、凡例順一致用）
 * @param {object}   dataMap       { timestamp: { sheetName: viewCount } }
 * @param {string[]} sortedTimestamps  タイムスタンプ（降順）
 * @returns {{ numRows: number, numCols: number }}
 */
function buildAbsoluteTimeHelperTable_(helperSheet, sortedNames, sortedTitles, dataMap, sortedTimestamps) {
  // グラフ用に昇順にする
  const ascTimestamps = [...sortedTimestamps].reverse();

  const headerRow = ['日時', ...sortedTitles];
  const dataRows  = ascTimestamps.map(ts => [
    new Date(ts),
    ...sortedNames.map(name => dataMap[ts]?.[name] ?? null),
  ]);
  const tableData = [headerRow, ...dataRows];

  const numRows = tableData.length;
  const numCols = tableData[0].length;

  // シートサイズをデータに合わせて拡張
  if (helperSheet.getMaxRows() < numRows) {
    helperSheet.insertRowsAfter(helperSheet.getMaxRows(), numRows - helperSheet.getMaxRows());
  }
  if (helperSheet.getMaxColumns() < numCols) {
    helperSheet.insertColumnsAfter(helperSheet.getMaxColumns(), numCols - helperSheet.getMaxColumns());
  }

  helperSheet.getRange(1, 1, numRows, numCols).setValues(tableData);

  return { numRows, numCols };
}

/**
 * 比較シートに絶対日時を横軸とした折れ線グラフを追加する。
 * ヘルパーテーブルのデータを参照する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} compSheet
 * @param {{ startCol: number, numRows: number, numCols: number }} helperInfo
 * @param {number} mainTableRows  メインテーブルの行数（グラフ位置の計算基点）
 */
function buildComparisonChart_(compSheet, helperSheet, helperInfo, mainTableRows) {
  const { WIDTH, HEIGHT } = CONFIG.CHART;
  const { numRows, numCols } = helperInfo;

  const chart = compSheet.newChart()
    .asLineChart()
    .addRange(helperSheet.getRange(1, 1, numRows, numCols))
    .setNumHeaders(1)
    .setPosition(mainTableRows + 2, 1, 0, 0)
    .setOption('title', '全動画 再生数推移')
    .setOption('width', WIDTH)
    .setOption('height', HEIGHT)
    .setOption('interpolateNulls', true)
    .setOption('pointSize', 2)
    .setOption('lineWidth', 2)
    .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
    .setOption('chartArea', { left: '6%', top: '10%', width: '65%', height: '75%' })
    .setOption('hAxis', {
      slantedText: true,
      slantedTextAngle: 30,
      textStyle: { fontSize: 9 },
      gridlines: { count: -1 },
    })
    .setOption('vAxis', {
      format: '#,###',
      gridlines: { color: '#b0b0b0' },
      minorGridlines: { count: 4, color: '#e8e8e8' },
    })
    .build();

  compSheet.insertChart(chart);
  console.log(`グラフを生成しました（${WIDTH}×${HEIGHT}px）`);
}

// ==========================================
// 8. 投稿日起点グラフ
// ==========================================
/**
 * 全動画の横軸を「投稿日からの経過日数」に正規化した折れ線グラフを生成する。
 * ヘルパーテーブルを専用シートに書き込みグラフのデータ源とする。
 * @param {GoogleAppsScript.Spreadsheet.Sheet}   compSheet       グラフ配置先
 * @param {GoogleAppsScript.Spreadsheet.Sheet}   helperSheet     データ書き込み先ヘルパーシート
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} videoSheets
 * @param {object} publishDateMap  { sheetName: Date }
 * @param {object} elapsedMaps    { sheetName: Map<elapsedHours, viewCount> }
 * @param {number} mainTableRows   メインテーブルの行数（グラフ位置の計算基点）
 */
function buildElapsedDaysChart_(compSheet, helperSheet, videoSheets, publishDateMap, elapsedMaps, mainTableRows) {
  const validSheets = videoSheets.filter(sh => publishDateMap[sh.getName()]);
  if (validSheets.length === 0) return;

  // aggregateVideoData_ で構築済みの経過日数マップを使用（シートの重複読み込みを省く）
  const videoMaps = validSheets.map(sh => elapsedMaps[sh.getName()] || new Map());

  // 全バケツの和集合（ソート済み）
  const allHoursSet = new Set();
  videoMaps.forEach(map => map.forEach((_, h) => allHoursSet.add(h)));
  const sortedHours = [...allHoursSet].sort((a, b) => a - b);

  // テーブル構築（ヘッダー行 + データ行）
  const titleRow = ['経過日数', ...validSheets.map(sh => sh.getRange('A1').getValue() || sh.getName())];
  const dataRows = sortedHours.map(h => [
    Math.round(h / 24 * 100) / 100, // 経過日数（小数第2位まで）
    ...videoMaps.map(map => map.has(h) ? map.get(h) : null),
  ]);
  const tableData = [titleRow, ...dataRows];

  const numRows  = tableData.length;
  const numCols  = tableData[0].length;

  // ヘルパーシートにデータを書き込む
  if (helperSheet.getMaxRows() < numRows) {
    helperSheet.insertRowsAfter(helperSheet.getMaxRows(), numRows - helperSheet.getMaxRows());
  }
  if (helperSheet.getMaxColumns() < numCols) {
    helperSheet.insertColumnsAfter(helperSheet.getMaxColumns(), numCols - helperSheet.getMaxColumns());
  }
  helperSheet.getRange(1, 1, numRows, numCols).setValues(tableData);

  // グラフ作成（1枚目グラフの下に配置）
  const { WIDTH, HEIGHT } = CONFIG.CHART;
  const chartRow = mainTableRows + 2 + Math.ceil(HEIGHT / 21) + 5;

  const chart = compSheet.newChart()
    .asLineChart()
    .addRange(helperSheet.getRange(1, 1, numRows, numCols))
    .setNumHeaders(1)
    .setPosition(chartRow, 1, 0, 0)
    .setOption('title', '全動画 再生数推移（投稿日起点）')
    .setOption('width', WIDTH)
    .setOption('height', HEIGHT)
    .setOption('interpolateNulls', true)
    .setOption('pointSize', 2)
    .setOption('lineWidth', 2)
    .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
    .setOption('chartArea', { left: '6%', top: '10%', width: '65%', height: '75%' })
    .setOption('hAxis', {
      title: '経過日数',
      slantedText: true,
      slantedTextAngle: 30,
      textStyle: { fontSize: 9 },
    })
    .setOption('vAxis', {
      format: '#,###',
      gridlines: { color: '#b0b0b0' },
      minorGridlines: { count: 4, color: '#e8e8e8' },
    })
    .build();

  compSheet.insertChart(chart);
  console.log('経過日数グラフを生成しました');
}

/**
 * 経過ミリ秒を SAMPLING.RULES の間隔でバケツ丸めした経過時間（h）を返す。
 * @param {number} elapsedMs
 * @returns {number}
 */
function getRoundedElapsedHours_(elapsedMs) {
  const elapsedDays   = elapsedMs / MS_PER_DAY;
  const rule          = CONFIG.SAMPLING.RULES.find(r => elapsedDays <= r.maxDays);
  const intervalHours = rule && rule.keepEveryHours ? rule.keepEveryHours : 1;
  return Math.round((elapsedMs / MS_PER_HOUR) / intervalHours) * intervalHours;
}

// ==========================================
// 9. チャンネル内順位グラフ
// ==========================================
/**
 * チャンネルごとに「チャンネル内順位の推移」折れ線グラフを比較シートに追加する。
 * 各動画シートの C 列（順位）を時系列で集約し、チャンネルごとに 1 グラフ生成する。
 * Y 軸は昇順逆転（direction:-1）により、順位 1（最上位）がグラフ上部に表示される。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet} compSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} rankHelperSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} videoSheets
 * @param {number} mainTableRows  テーブル行数（グラフ位置の計算基点）
 */
function buildChannelRankCharts_(ss, compSheet, rankHelperSheet, videoSheets, mainTableRows) {
  const { WIDTH, HEIGHT } = CONFIG.CHART;

  // 動画シートをチャンネルごとにグループ化（B2 = channelId、C2 = channelTitle）
  const channelGroups = {}; // { channelId: { title, sheets: [] } }
  videoSheets.forEach(sh => {
    const channelId    = sh.getRange('B2').getValue();
    const channelTitle = sh.getRange('C2').getValue() || channelId || 'Unknown';
    if (!channelId) return;
    if (!channelGroups[channelId]) channelGroups[channelId] = { title: channelTitle, sheets: [] };
    channelGroups[channelId].sheets.push(sh);
  });

  if (Object.keys(channelGroups).length === 0) {
    console.log('チャンネル情報がないためランクグラフをスキップ（次回 main() 実行後に記録されます）');
    return;
  }

  // 絶対時刻グラフ・経過日数グラフの下に配置
  const baseChartRow = mainTableRows + 2 + 2 * (Math.ceil(HEIGHT / 21) + 5);
  let chartOffset = 0;

  Object.entries(channelGroups)
    .sort(([a], [b]) => a.localeCompare(b)) // チャンネル ID でソートして順序を安定させる
    .forEach(([channelId, { title: channelTitle, sheets }]) => {
      // 各動画シートの C 列（順位）を時系列で収集
      const allTimestamps = new Set();
      const rankDataMap   = {}; // { timestamp: { sheetName: rank } }

      sheets.forEach(sh => {
        const name    = sh.getName();
        const lastRow = sh.getLastRow();
        if (lastRow < 4) return;

        sh.getRange(4, 1, lastRow - 3, 3).getValues().forEach(row => {
          if (!(row[0] instanceof Date)) return;
          const rank = row[2];
          if (rank === '' || rank == null) return; // 順位未記録行はスキップ
          const ts = formatTimestamp_(row[0]);
          allTimestamps.add(ts);
          setNestedValue_(rankDataMap, ts, name, Number(rank));
        });
      });

      if (allTimestamps.size === 0) {
        console.log(`チャンネル ${channelTitle}: 順位データなしのためグラフをスキップ`);
        return;
      }

      const sheetNames  = sheets.map(sh => sh.getName());
      const sheetTitles = sheets.map(sh => sh.getRange('A1').getValue() || sh.getName());
      const ascTs       = [...allTimestamps].sort(); // 昇順（グラフの左→右）

      const tableData = [
        ['日時', ...sheetTitles],
        ...ascTs.map(ts => [
          new Date(ts),
          ...sheetNames.map(name => rankDataMap[ts]?.[name] ?? null),
        ]),
      ];

      const numRows = tableData.length;
      const numCols = tableData[0].length;

      if (rankHelperSheet.getMaxRows() < numRows) {
        rankHelperSheet.insertRowsAfter(rankHelperSheet.getMaxRows(), numRows - rankHelperSheet.getMaxRows());
      }
      if (rankHelperSheet.getMaxColumns() < numCols) {
        rankHelperSheet.insertColumnsAfter(rankHelperSheet.getMaxColumns(), numCols - rankHelperSheet.getMaxColumns());
      }
      rankHelperSheet.getRange(1, 1, numRows, numCols).setValues(tableData);

      const chart = compSheet.newChart()
        .asLineChart()
        .addRange(rankHelperSheet.getRange(1, 1, numRows, numCols))
        .setNumHeaders(1)
        .setPosition(baseChartRow + chartOffset, 1, 0, 0)
        .setOption('title', `チャンネル内順位推移 — ${channelTitle}（値が小さいほど上位）`)
        .setOption('width', WIDTH)
        .setOption('height', HEIGHT)
        .setOption('interpolateNulls', true)
        .setOption('pointSize', 2)
        .setOption('lineWidth', 2)
        .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
        .setOption('chartArea', { left: '6%', top: '10%', width: '65%', height: '75%' })
        .setOption('hAxis', {
          slantedText: true,
          slantedTextAngle: 30,
          textStyle: { fontSize: 9 },
        })
        .setOption('vAxis', {
          format: '#,##0',
          direction: -1, // 順位 1（最上位）をグラフ上部に表示
          gridlines: { color: '#b0b0b0' },
          minorGridlines: { count: 4, color: '#e8e8e8' },
        })
        .build();

      compSheet.insertChart(chart);
      console.log(`チャンネル内順位グラフを生成しました（${channelTitle}、${sheets.length} 動画）`);
      chartOffset += Math.ceil(HEIGHT / 21) + 5;
    });
}

// ==========================================
// 10. ユーティリティ
// ==========================================
/**
 * Date を "yyyy/MM/dd HH:mm"（JST）にフォーマットする。
 * @param {Date} date
 * @returns {string}
 */
function formatTimestamp_(date) {
  return Utilities.formatDate(date, 'JST', 'yyyy/MM/dd HH:mm');
}

/**
 * タイムアウトしやすい Spreadsheet 操作をリトライ付きで実行する。
 * @param {Function} fn      実行する関数
 * @param {number}   maxRetries  最大リトライ回数（デフォルト 3）
 * @returns {*} fn の戻り値
 */
function retryOnTimeout_(fn, maxRetries = 3) {
  for (let i = 0; i <= maxRetries; i++) {
    try {
      return fn();
    } catch (e) {
      if (i === maxRetries || !e.message.includes('timed out')) throw e;
      console.warn(`タイムアウト発生、リトライ ${i + 1}/${maxRetries}...`);
      SpreadsheetApp.flush();
      Utilities.sleep(3000 * (i + 1));
    }
  }
}

/**
 * allData（降順ソート済み）から指定ウィンドウの増加量を計算する。
 * @param {any[][]} allData          [[Date, viewCount], ...] 降順
 * @param {number}  currentViewCount 最新再生数
 * @param {number}  nowMs            現在時刻（ms）
 * @param {number}  fromMs           ウィンドウ開始（nowMs から遡る ms）
 * @param {number}  toMs             ウィンドウ終了（0 = 現在値を使用）
 * @returns {{ value: number, isEstimate: boolean } | null}
 */
function calcIncrease_(allData, currentViewCount, nowMs, fromMs, toMs) {
  const windowMs = fromMs - toMs;

  function findNearest(targetMs) {
    const row = allData.find(r => r[0] instanceof Date && r[0].getTime() <= targetMs);
    return row ? { v: row[1], t: row[0].getTime() } : null;
  }

  const endPt   = toMs === 0
    ? { v: currentViewCount, t: nowMs }
    : findNearest(nowMs - toMs);
  const startPt = findNearest(nowMs - fromMs);

  if (!endPt || !startPt) return null;

  const actualSpanMs = endPt.t - startPt.t;
  if (actualSpanMs <= 0) return null;

  const value = Math.round((endPt.v - startPt.v) / actualSpanMs * windowMs);

  const startErr   = Math.abs(startPt.t - (nowMs - fromMs));
  const endErr     = toMs === 0 ? 0 : Math.abs(endPt.t - (nowMs - toMs));
  const isEstimate = !(startErr <= fromMs && endErr <= fromMs);

  return { value, isEstimate };
}

/**
 * calcIncrease_ の結果を "320 (0.6%)" 形式の文字列にフォーマットする。
 * @param {{ value: number, isEstimate: boolean } | null} result
 * @param {number} currentViewCount
 * @returns {string}
 */
function formatIncreaseWithRate_(result, currentViewCount) {
  if (!result) return '---';
  const { value, isEstimate } = result;
  const prefix = isEstimate ? '~' : '';
  const base   = currentViewCount - value;
  if (base > 0) {
    const rate = (value / base * 100).toFixed(1);
    return `${prefix}${value} (${rate}%)`;
  }
  return `${prefix}${value}`;
}

/**
 * ネストされたオブジェクトに安全に値をセットする。
 * @param {object} obj
 * @param {string} key1
 * @param {string} key2
 * @param {any}    value
 */
function setNestedValue_(obj, key1, key2, value) {
  if (!obj[key1]) obj[key1] = {};
  obj[key1][key2] = value;
}

// ==========================================
// 11. 増加量サマリー
// ==========================================
/**
 * 各動画シートの D1:E6 に直近の再生数増加量サマリーを書き込む。
 * データは降順ソート済みを前提とする（sortVideoSheetDescending_ 後に呼ぶこと）。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} currentViewCount  最新の再生数
 * @param {Date}   now               計測日時
 */
function updateGrowthSummary_(sheet, currentViewCount, now, allData) {
  if (!allData || allData.length === 0) return;

  const nowMs = now.getTime();

  // calcIncrease_ のラッパー: 既存の文字列形式（数値 or '~数値' or '---'）で返す
  function calcIncrease(fromMs, toMs) {
    const result = calcIncrease_(allData, currentViewCount, nowMs, fromMs, toMs);
    if (!result) return '---';
    return result.isEstimate ? `~${result.value}` : result.value;
  }

  // 各期間ごとのウィンドウを生成: [[fromMs, toMs], ...]
  const makeWindows = (unit, count) =>
    Array.from({ length: count }, (_, i) => [(i + 1) * unit, i * unit]);

  const N = CONFIG.SUMMARY_WINDOWS;
  const PERIODS = [
    { label: '1時間', windows: makeWindows(MS_PER_HOUR,     N) },
    { label: '1日',   windows: makeWindows(MS_PER_DAY,      N) },
    { label: '1週間', windows: makeWindows(7  * MS_PER_DAY, N) },
    { label: '1ヶ月', windows: makeWindows(30  * MS_PER_DAY, N) },
    { label: '1年',   windows: makeWindows(365 * MS_PER_DAY, N) },
  ];

  const headers = ['期間', '直近', ...Array.from({ length: N - 1 }, (_, i) => `${i + 1}期前`)];
  const tableData = [
    headers,
    ...PERIODS.map(({ label, windows }) => {
      const vals = windows.map(([from, to]) => calcIncrease(from, to));
      while (vals.length < N) vals.push('---');
      return [label, ...vals];
    }),
  ];

  const numRows  = tableData.length;
  const numCols  = headers.length;
  const startRow = 4; // 固定行(1〜3)の下から開始
  sheet.getRange(startRow, 4, numRows, numCols).setValues(tableData);

  // 書式設定（テーブルサイズ変更に対応するためバージョン管理）
  const fmtKey = `summary_fmt_v2_${sheet.getName()}`;
  if (!PropertiesService.getScriptProperties().getProperty(fmtKey)) {
    sheet.getRange(startRow, 4, 1, numCols).setBackground('#eeeeee').setFontWeight('bold');
    sheet.getRange(startRow + 1, 5, numRows - 1, numCols - 1).setNumberFormat('#,##0');
    PropertiesService.getScriptProperties().setProperty(fmtKey, 'true');
  }
}

// ==========================================
// 12. 初期成長曲線の補完
// ==========================================
/**
 * 投稿日(0再生)と最初の計測値の間を、べき乗則曲線で補完する。
 *
 * alpha の推定には最初の 1〜3 計測点を log-log 空間で重み付き最小二乗法により求める。
 * 重み [1, 2, 4] は後の計測点ほど大きく設定し、最新の勾配に合わせる。
 * 計測点が 1 点のみの場合は平方根曲線（alpha = 0.5）にフォールバックする。
 *
 * SAMPLING.RULES と同じ頻度の補完点を挿入するため、
 * 補完後に間引きの対象外となり自然に保持される。
 *
 * ScriptProperties でシートごとに実行済みフラグを管理し、一度だけ実行する。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Date} publishedAt
 */
function fillInitialGrowthCurve_(sheet, publishedAt) {
  const propKey = `curve_filled_${sheet.getName()}`;
  if (PropertiesService.getScriptProperties().getProperty(propKey)) return;

  const dataCount = sheet.getLastRow() - 3;
  if (dataCount < 2) return; // 起点行 + 最低1計測が必要

  const t0 = publishedAt.getTime();

  // 起点行(row4)直後の最大3計測点を読み込む（row5〜row7）
  const nSamples = Math.min(dataCount - 1, 3);
  const rawRows  = sheet.getRange(5, 1, nSamples, 2).getValues();

  const points = rawRows
    .filter(r => r[0] instanceof Date && r[1] > 0)
    .map(r => ({ d: r[0].getTime() - t0, v: r[1] }))
    .filter(p => p.d > 0);

  if (points.length === 0) return;

  // log-log 空間での重み付き最小二乗法で alpha を推定
  // Y = log(v) = log(C) + alpha * log(d) = log(C) + alpha * X
  let alpha = 0.5; // フォールバック: 平方根曲線
  if (points.length >= 2) {
    const ws = points.map((_, i) => Math.pow(2, i)); // [1, 2, 4]: 後の点ほど重く
    const xs = points.map(p => Math.log(p.d));
    const ys = points.map(p => Math.log(p.v));

    const W    = ws.reduce((s, w)    => s + w, 0);
    const xBar = ws.reduce((s, w, i) => s + w * xs[i], 0) / W;
    const yBar = ws.reduce((s, w, i) => s + w * ys[i], 0) / W;
    const num  = ws.reduce((s, w, i) => s + w * (xs[i] - xBar) * (ys[i] - yBar), 0);
    const den  = ws.reduce((s, w, i) => s + w * Math.pow(xs[i] - xBar, 2), 0);

    if (den > 0) alpha = Math.max(0.1, Math.min(1.5, num / den));
  }

  const { d: d1, v: v1 } = points[0];
  const C = v1 / Math.pow(d1, alpha);

  // SAMPLING.RULES と同じ頻度で補完点を生成する
  const newRows = [];
  let curMs = t0;
  while (true) {
    const ageDays = (curMs - t0) / MS_PER_DAY;
    const rule    = CONFIG.SAMPLING.RULES.find(r => ageDays <= r.maxDays);
    const stepMs  = (rule && rule.keepEveryHours ? rule.keepEveryHours : 1) * MS_PER_HOUR;
    curMs += stepMs;
    if (curMs >= t0 + d1) break;
    const v = Math.round(C * Math.pow(curMs - t0, alpha));
    if (v <= 0) continue;
    newRows.push([new Date(curMs), v, '']); // 補間点は順位未計測
  }
  if (newRows.length === 0) return;

  sheet.insertRowsAfter(4, newRows.length);
  sheet.getRange(5, 1, newRows.length, 3).setValues(newRows);
  PropertiesService.getScriptProperties().setProperty(propKey, 'true');
  console.log(`成長曲線を補完 [${sheet.getName()}]: ${newRows.length} 点追加 (alpha=${alpha.toFixed(2)})`);
}

// ==========================================
// 13. シート並び替え
// ==========================================
/**
 * 動画シートを投稿日時の昇順（古い順が左、新しい順が右）に並び替える。
 * PRESERVE_SHEET_NAMES のシートは先頭（左側）に固定する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function sortVideoSheetsByPublishDate_(ss) {
  const sheets = ss.getSheets();
  const preserved = sheets.filter(s => CONFIG.PRESERVE_SHEET_NAMES.includes(s.getName()));
  const videoSheets = sheets.filter(s => !CONFIG.PRESERVE_SHEET_NAMES.includes(s.getName()));

  videoSheets.sort((a, b) => {
    const dateA = a.getRange('A2').getValue();
    const dateB = b.getRange('A2').getValue();
    if (!(dateA instanceof Date)) return 1;
    if (!(dateB instanceof Date)) return -1;
    return dateA.getTime() - dateB.getTime(); // 昇順: 古い順が左
  });

  [...preserved, ...videoSheets].forEach((sheet, index) => {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(index + 1);
  });
}

// ==========================================
// 14. 管理用ユーティリティ（手動実行）
// ==========================================
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
