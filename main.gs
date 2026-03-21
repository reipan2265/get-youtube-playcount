// ==========================================
// 設定エリア（ここだけ編集してOK）
// ==========================================
const CONFIG = {
  // 対象プレイリストの ID（不要な場合は空文字）
  PLAYLIST_ID: 'PLriG7RRWaKk-YG8N7y4Fr8C15NJqnkLYG',

  // プレイリスト外で個別追加したい動画 ID
  EXTRA_VIDEO_IDS: ['Z_BpyttvaKI', 'WGrgo8-8XwY'],

  // 全動画比較シートのシート名
  COMP_SHEET_NAME: '再生数比較',

  // 削除・リセット対象から除外するシート名
  PRESERVE_SHEET_NAMES: ['再生数比較', 'シート1', '_abs_helper', '_elapsed_helper'],

  // 比較グラフのサイズ（ピクセル）
  CHART: {
    WIDTH:  2210,
    HEIGHT:  850,
  },

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

  const videoIds = collectVideoIds_();
  console.log(`対象: ${videoIds.length} 本`);

  videoIds.forEach((id, index) => {
    processVideo_(ss, id, index, videoIds.length, now);
    SpreadsheetApp.flush();
  });

  console.log('比較シートを更新します...');
  updateComparisonSheet_(ss);
  console.log('完了。');
}

// ==========================================
// 2. 動画ごとの処理
// ==========================================
/**
 * 1本の動画を処理する（取得 → 記録 → 補完 → 間引き → グラフ更新）。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} id     YouTube 動画 ID
 * @param {number} index  現在のインデックス（ログ用）
 * @param {number} total  総件数（ログ用）
 * @param {Date}   now    実行日時
 */
function processVideo_(ss, id, index, total, now) {
  try {
    const video = fetchVideoData_(id);
    if (!video) {
      console.warn(`[${index + 1}/${total}] 動画が見つかりません (id: ${id})`);
      return;
    }

    const { fullTitle, viewCount, publishedAt } = parseVideoData_(video);
    const sheetName = buildSheetName_(fullTitle, id);
    const sheet     = getOrCreateVideoSheet_(ss, sheetName, fullTitle, publishedAt);

    sheet.appendRow([now, viewCount]);
    console.log(`[${index + 1}/${total}] ${sheetName}: ${viewCount.toLocaleString()} 回`);

    fillInitialGrowthCurve_(sheet, publishedAt);
    runSampling_(sheet, publishedAt);
    updateIndividualChart_(sheet);
    sortVideoSheetDescending_(sheet);
    // ソート後のデータを1回読んで渡す（updateGrowthSummary_ 内の重複読み込みを省く）
    const lastRow_ = sheet.getLastRow();
    const allData_ = lastRow_ >= 4 ? sheet.getRange(4, 1, lastRow_ - 3, 2).getValues() : [];
    updateGrowthSummary_(sheet, viewCount, now, allData_);

  } catch (e) {
    console.error(`[${index + 1}/${total}] エラー (id: ${id}): ${e.message}\n${e.stack}`);
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
    fullTitle  : video.snippet.title,
    viewCount  : Number(video.statistics.viewCount),
    publishedAt: new Date(video.snippet.publishedAt),
  };
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
 * EXTRA_VIDEO_IDS とプレイリストを合わせて重複排除した動画 ID リストを返す。
 * @returns {string[]}
 */
function collectVideoIds_() {
  const playlistIds = fetchPlaylistVideoIds_();
  console.log(`プレイリストから ${playlistIds.length} 本を検出`);
  return [...new Set([...CONFIG.EXTRA_VIDEO_IDS, ...playlistIds])];
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
function getOrCreateVideoSheet_(ss, sheetName, fullTitle, publishedAt) {
  let sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    // 既存シートで投稿日が未記入の場合のみ補完
    if (!sheet.getRange('A2').getValue()) {
      sheet.getRange('A2').setValue(publishedAt);
    }
    return sheet;
  }

  console.log(`シートを作成: ${sheetName}`);
  sheet = ss.insertSheet(sheetName);

  sheet.getRange('A1').setValue(fullTitle).setFontWeight('bold');
  sheet.getRange('A2').setValue(publishedAt);
  sheet.getRange('A3:B3').setValues([['日時', '再生数']]).setBackground('#eeeeee');
  sheet.setFrozenRows(3);
  sheet.appendRow([publishedAt, 0]); // 投稿日時点の起点レコード（再生数 = 0）

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
  sheet.getRange(4, 1, lastRow - 3, 2).sort({ column: 1, ascending: false });
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

  const values     = sheet.getRange(4, 1, dataCount, 2).getValues();
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
    sheet.getRange(4, 1, sheet.getLastRow() - 3, 2).clearContent();
    sheet.getRange(4, 1, keepRows.length, 2).setValues(keepRows);
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
 * 全動画の再生数推移を1シートに集約し、比較グラフを2種類生成する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function updateComparisonSheet_(ss) {
  SpreadsheetApp.flush();
  Utilities.sleep(2000);

  const compSheet   = ss.getSheetByName(CONFIG.COMP_SHEET_NAME) || ss.insertSheet(CONFIG.COMP_SHEET_NAME, 0);
  const videoSheets = ss.getSheets().filter(s => !CONFIG.PRESERVE_SHEET_NAMES.includes(s.getName()));

  const { dataMap, publishDateMap, elapsedMaps, sortedTimestamps } = aggregateVideoData_(videoSheets);
  if (sortedTimestamps.length === 0) {
    console.warn('比較シートに書き込むデータがありません');
    return;
  }

  const { tableValues, sortedNames } = buildComparisonTable_(videoSheets, dataMap, publishDateMap, sortedTimestamps);

  renderComparisonSheet_(compSheet, tableValues);

  const absHelperSheet = getOrCreateHelperSheet_(ss, '_abs_helper');
  const absHelper = buildAbsoluteTimeHelperTable_(absHelperSheet, sortedNames, dataMap, sortedTimestamps);
  buildComparisonChart_(compSheet, absHelperSheet, absHelper, tableValues.length);

  const elapsedHelperSheet = getOrCreateHelperSheet_(ss, '_elapsed_helper');
  buildElapsedDaysChart_(compSheet, elapsedHelperSheet, videoSheets, publishDateMap, elapsedMaps, tableValues.length);
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
  const now = new Date();

  const videoRows = videoSheets
    .map(sh => {
      const name    = sh.getName();
      const pubDate = publishDateMap[name];
      if (!pubDate) return null;

      const values  = sortedTimestamps.map(ts => dataMap[ts]?.[name] ?? null);
      // sortedTimestamps は降順なので先頭が最新値（reverseは不要）
      const lastVal = values.find(v => v !== null) ?? 0;

      // 投稿後30日以内の最後のデータ点を基準に1日平均を算出
      // 30日未満の動画はその経過日数で割る
      const FIRST_MONTH_DAYS = 30;
      const targetMs         = pubDate.getTime() + FIRST_MONTH_DAYS * MS_PER_DAY;
      const daysSincePublish = (now.getTime() - pubDate.getTime()) / MS_PER_DAY;

      let baseViews, baseDays;
      if (daysSincePublish >= FIRST_MONTH_DAYS) {
        // 30日時点以前で最も新しいデータ点を探す（降順リストの先頭から見て targetMs 以下の最初の値）
        const ts30 = sortedTimestamps.find(ts =>
          dataMap[ts]?.[name] != null && new Date(ts).getTime() <= targetMs
        );
        baseViews = ts30 ? dataMap[ts30][name] : lastVal;
        baseDays  = FIRST_MONTH_DAYS;
      } else {
        baseViews = lastVal;
        baseDays  = Math.max(1, daysSincePublish);
      }
      const dailyAvg = Math.round(baseViews / baseDays);

      const escapedName  = name.replace(/'/g, "''");
      const titleFormula = `='${escapedName}'!$A$1`;

      return { row: [titleFormula, dailyAvg], lastVal, name };
    })
    .filter(Boolean)
    .sort((a, b) => b.lastVal - a.lastVal);

  const headerRow = ['動画名', '1日平均再生数'];
  return {
    tableValues: [headerRow, ...videoRows.map(r => r.row)],
    sortedNames: videoRows.map(r => r.name),
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
  compSheet.getRange(1, 2, rows, 1).setBackground('#fff2cc').setFontWeight('bold'); // B列（1日平均）を強調
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
function buildAbsoluteTimeHelperTable_(helperSheet, sortedNames, dataMap, sortedTimestamps) {
  // グラフ用に昇順にする
  const ascTimestamps = [...sortedTimestamps].reverse();

  const headerRow = ['日時', ...sortedNames];
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
// 9. ユーティリティ
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
// 10. 増加量サマリー
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

  // targetMs 以前で最も近いデータ点を {v, t} で返す
  function findNearest(targetMs) {
    const row = allData.find(r => r[0] instanceof Date && r[0].getTime() <= targetMs);
    return row ? { v: row[1], t: row[0].getTime() } : null;
  }

  // fromMs前〜toMs前の増加量を返す。toMs=0 は現在値を使用。
  // 正確なデータがない場合は隣接点から線形補間し ~ プレフィックスで推定値を返す。
  // 両端が同じデータ点になる場合（推定不能）は --- を返す。
  function calcIncrease(fromMs, toMs) {
    const windowMs = fromMs - toMs;

    const endPt   = toMs === 0
      ? { v: currentViewCount, t: nowMs }
      : findNearest(nowMs - toMs);
    const startPt = findNearest(nowMs - fromMs);

    if (!endPt || !startPt) return '---';

    const actualSpanMs = endPt.t - startPt.t;
    if (actualSpanMs <= 0) return '---'; // 同一点なので推定不能

    const value = Math.round((endPt.v - startPt.v) / actualSpanMs * windowMs);

    // 両端のデータ点が期間幅(fromMs)以内の誤差なら正確値、それ以外は推定値
    const startErr = Math.abs(startPt.t - (nowMs - fromMs));
    const endErr   = toMs === 0 ? 0 : Math.abs(endPt.t - (nowMs - toMs));
    return startErr <= fromMs && endErr <= fromMs ? value : `~${value}`;
  }

  // windows: [[fromMs, toMs], ...]  各3ウィンドウ
  const PERIODS = [
    { label: '1時間', windows: [
      [          MS_PER_HOUR,           0 ],
      [  2 * MS_PER_HOUR,   MS_PER_HOUR ],
      [  3 * MS_PER_HOUR, 2 * MS_PER_HOUR ],
    ]},
    { label: '1日',   windows: [
      [      MS_PER_DAY,           0 ],
      [  2 * MS_PER_DAY,   MS_PER_DAY ],
      [  3 * MS_PER_DAY, 2 * MS_PER_DAY ],
    ]},
    { label: '1週間', windows: [
      [  7 * MS_PER_DAY,           0 ],
      [ 14 * MS_PER_DAY,  7 * MS_PER_DAY ],
      [ 21 * MS_PER_DAY, 14 * MS_PER_DAY ],
    ]},
    { label: '1ヶ月', windows: [
      [ 30 * MS_PER_DAY,           0 ],
      [ 60 * MS_PER_DAY, 30 * MS_PER_DAY ],
      [ 90 * MS_PER_DAY, 60 * MS_PER_DAY ],
    ]},
  ];

  const maxCols  = 3; // 最大ウィンドウ数
  const headers  = ['期間', '直近', '前の期間', '更に前の期間'];
  const tableData = [
    headers,
    ...PERIODS.map(({ label, windows }) => {
      const vals = windows.map(([from, to]) => calcIncrease(from, to));
      while (vals.length < maxCols) vals.push('---');
      return [label, ...vals];
    }),
  ];

  const numRows  = tableData.length;
  const numCols  = headers.length;
  const startRow = 4; // 固定行(1〜3)の下から開始
  sheet.getRange(startRow, 4, numRows, numCols).setValues(tableData);

  // 書式設定は初回のみ（以降はセルに保持される）
  const fmtKey = `summary_fmt_${sheet.getName()}`;
  if (!PropertiesService.getScriptProperties().getProperty(fmtKey)) {
    sheet.getRange(startRow, 4, 1, numCols).setBackground('#eeeeee').setFontWeight('bold');
    sheet.getRange(startRow + 1, 5, numRows - 1, numCols - 1).setNumberFormat('#,##0');
    PropertiesService.getScriptProperties().setProperty(fmtKey, 'true');
  }
}

// ==========================================
// 11. 初期成長曲線の補完
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
    newRows.push([new Date(curMs), v]);
  }
  if (newRows.length === 0) return;

  sheet.insertRowsAfter(4, newRows.length);
  sheet.getRange(5, 1, newRows.length, 2).setValues(newRows);
  PropertiesService.getScriptProperties().setProperty(propKey, 'true');
  console.log(`成長曲線を補完 [${sheet.getName()}]: ${newRows.length} 点追加 (alpha=${alpha.toFixed(2)})`);
}

// ==========================================
// 12. 管理用ユーティリティ（手動実行）
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
