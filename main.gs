// ==========================================
// 設定エリア（ここだけ編集してOK）
// ==========================================
const CONFIG = {
  PLAYLIST_ID: 'PLriG7RRWaKk-YG8N7y4Fr8C15NJqnkLYG',
  EXTRA_VIDEO_IDS: ['Z_BpyttvaKI', 'WGrgo8-8XwY'],
  COMP_SHEET_NAME: '再生数比較',
  PRESERVE_SHEET_NAMES: ['再生数比較', 'シート1'],
  CHART: {
    WIDTH: 2210,
    HEIGHT: 850,
  },
  SAMPLING: {
    MIN_ROWS_TO_SAMPLE: 10,
    // 経過日数ごとの保持間隔（時間）: null = 毎時, 6 = 6時間ごと, 等
    RULES: [
      { maxDays: 30,  keepEveryHours: null },   // 30日以内: 全件保持
      { maxDays: 90,  keepEveryHours: 6    },   // ~90日: 6時間ごと
      { maxDays: 180, keepEveryHours: 12   },   // ~180日: 12時間ごと
      { maxDays: 365, keepEveryHours: 24   },   // ~365日: 1日ごと
      { maxDays: Infinity, keepEveryHours: 168 }, // 365日超: 週1
    ],
  },
};

// ==========================================
// 定数
// ==========================================
const MS_PER_DAY  = 24 * 60 * 60 * 1000;
const MS_PER_HOUR = 60 * 60 * 1000;

// ==========================================
// 1. メイン処理
// ==========================================
/**
 * トリガーから呼び出すエントリーポイント。
 * 全動画の再生数を取得してシートに記録し、比較シートを更新する。
 */
function main() {
  console.log('🚀 再生数取得プロセスを開始します...');

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const now   = new Date();

  const videoIds = collectVideoIds_();
  console.log(`🎬 処理対象: ${videoIds.length} 本`);

  videoIds.forEach((id, index) => processVideo_(ss, id, index, videoIds.length, now));

  console.log('📊 比較シートを更新します...');
  updateComparisonSheet_(ss);
  console.log('✨ 全プロセス完了。');
}

// ==========================================
// 2. 動画ごとの処理
// ==========================================
/**
 * 1本の動画を処理する（取得 → シート書き込み → 間引き → グラフ更新）。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} id  YouTube動画ID
 * @param {number} index  現在のインデックス（ログ用）
 * @param {number} total  総件数（ログ用）
 * @param {Date}   now  実行日時
 */
function processVideo_(ss, id, index, total, now) {
  try {
    const video = fetchVideoData_(id);
    if (!video) {
      console.warn(`⚠️ [${index + 1}/${total}] ID:${id} が見つかりません。`);
      return;
    }

    const { fullTitle, viewCount, publishedAt } = parseVideoData_(video);
    const sheetName = buildSheetName_(fullTitle, id);
    const sheet     = getOrCreateVideoSheet_(ss, sheetName, fullTitle, publishedAt);

    sheet.appendRow([now, viewCount]);
    console.log(`✅ [${index + 1}/${total}] ${sheetName}: ${viewCount.toLocaleString()} 回`);

    fillInitialGrowthCurve_(sheet, publishedAt);
    runSampling_(sheet, publishedAt);
    updateIndividualChart_(sheet);

  } catch (e) {
    console.error(`❌ エラー (ID:${id}): ${e.message}\n${e.stack}`);
  }
}

// ==========================================
// 3. YouTube API ラッパー
// ==========================================
/**
 * YouTube Data API から動画情報を取得する。
 * @param {string} id
 * @returns {object|null} API レスポンスの items[0]、または null
 */
function fetchVideoData_(id) {
  const res = YouTube.Videos.list('snippet,statistics', { id });
  return res.items?.[0] ?? null;
}

/**
 * APIレスポンスから必要フィールドを抽出して返す。
 * @param {object} video
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
 * プレイリスト内の全動画IDを取得する（ページネーション対応）。
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
    console.warn(`⚠️ プレイリスト取得失敗: ${e.message}`);
  }

  return ids;
}

/**
 * EXTRA_VIDEO_IDS とプレイリストを合わせて重複排除した動画IDリストを返す。
 * @returns {string[]}
 */
function collectVideoIds_() {
  const playlistIds = fetchPlaylistVideoIds_();
  console.log(`📂 プレイリストから ${playlistIds.length} 本を検出。`);
  return [...new Set([...CONFIG.EXTRA_VIDEO_IDS, ...playlistIds])];
}

// ==========================================
// 4. シート操作
// ==========================================
/**
 * 動画タイトルからシート名を生成する（Sheets の制約: 31文字以内、特殊文字禁止）。
 * @param {string} fullTitle
 * @param {string} fallback  タイトルが空の場合に使う動画ID
 * @returns {string}
 */
function buildSheetName_(fullTitle, fallback) {
  const INVALID_CHARS = /[\\\/\?\*\:\[\]]/g;
  const MAX_LEN = 31;

  let name = fullTitle.replace(/【.*?】/g, '').trim().replace(INVALID_CHARS, '_');

  if (!name) return fallback;

  if (name.length > MAX_LEN) {
    const half = Math.floor((MAX_LEN - 1) / 2); // 省略記号1文字分を引く
    name = name.substring(0, half) + '…' + name.substring(name.length - half);
  }

  return name;
}

/**
 * 動画シートを取得、なければ新規作成してヘッダーを書き込む。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @param {string} fullTitle
 * @param {Date}   publishedAt
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateVideoSheet_(ss, sheetName, fullTitle, publishedAt) {
  let sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    // 既存シートでも投稿日が未記入なら補完
    if (!sheet.getRange('A2').getValue()) {
      sheet.getRange('A2').setValue(publishedAt).setFontColor('white');
    }
    return sheet;
  }

  console.log(`📝 新規シート作成: ${sheetName}`);
  sheet = ss.insertSheet(sheetName);

  // A1: タイトル（表示用）
  sheet.getRange('A1').setValue(fullTitle).setFontWeight('bold');
  // A2: 投稿日（グラフ起点計算用・白文字で非表示風に）
  sheet.getRange('A2').setValue(publishedAt).setFontColor('white');
  // A3:B3: ヘッダー行
  sheet.getRange('A3:B3').setValues([['日時', '再生数']]).setBackground('#eeeeee');
  sheet.setFrozenRows(3);

  // 投稿日時点での0件レコードをグラフ起点として挿入
  sheet.appendRow([publishedAt, 0]);

  return sheet;
}

// ==========================================
// 5. サンプリング（データの間引き）
// ==========================================
/**
 * 古いデータを間引いてシートの行数を抑制する。
 * CONFIG.SAMPLING.RULES に従い、経過日数ごとに保持間隔を変える。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Date} publishedAt
 */
function runSampling_(sheet, publishedAt) {
  const lastRow   = sheet.getLastRow();
  const dataCount = lastRow - 3; // ヘッダー3行を除いたデータ行数

  if (dataCount < CONFIG.SAMPLING.MIN_ROWS_TO_SAMPLE) return;

  const values     = sheet.getRange(4, 1, dataCount, 2).getValues();
  const seenBucket = new Set(); // バケツIDの重複チェック用

  const keepRows = values.filter((row, index) => {
    // 先頭・末尾は必ず保持
    if (index === 0 || index === values.length - 1) return true;
    if (!(row[0] instanceof Date)) return false;

    const ageDays = (row[0].getTime() - publishedAt.getTime()) / MS_PER_DAY;
    const rule    = CONFIG.SAMPLING.RULES.find(r => ageDays <= r.maxDays);

    // keepEveryHours が null なら全件保持
    if (!rule || rule.keepEveryHours === null) return true;

    // バケツ方式: 経過時間をkeepEveryHoursで割った整数が同じ行は最初の1件だけ保持
    // 剰余方式と違い、記録タイミングがズレていても均等に間引ける
    const ageHours  = (row[0].getTime() - publishedAt.getTime()) / MS_PER_HOUR;
    const bucketKey = `${rule.keepEveryHours}_${Math.floor(ageHours / rule.keepEveryHours)}`;

    if (seenBucket.has(bucketKey)) return false;
    seenBucket.add(bucketKey);
    return true;
  });

  // 間引き後に最低保持数を下回るなら今回はスキップ
  if (keepRows.length < CONFIG.SAMPLING.MIN_ROWS_TO_SAMPLE) {
    console.log(`⏭️ 間引きスキップ [${sheet.getName()}]: 間引き後 ${keepRows.length} 行 < 最低 ${CONFIG.SAMPLING.MIN_ROWS_TO_SAMPLE} 行`);
    return;
  }

  if (keepRows.length < values.length) {
    console.log(`✂️ 間引き [${sheet.getName()}]: ${values.length} → ${keepRows.length} 行`);
    sheet.getRange(4, 1, lastRow - 3, 2).clearContent();
    sheet.getRange(4, 1, keepRows.length, 2).setValues(keepRows);
  }
}

// ==========================================
// 6. 個別動画グラフの更新
// ==========================================
/**
 * 各動画シートに再生数推移の折れ線グラフを作成・更新する。
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
    .setPosition(1, 4, 0, 0)
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
// 7. 比較シートの作成・更新
// ==========================================
/**
 * 全動画の再生数推移を1シートに集約し、比較グラフを生成する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function updateComparisonSheet_(ss) {
  SpreadsheetApp.flush();
  Utilities.sleep(2000); // flush が反映されるまで待機

  const compSheet   = ss.getSheetByName(CONFIG.COMP_SHEET_NAME) || ss.insertSheet(CONFIG.COMP_SHEET_NAME, 0);
  const videoSheets = ss.getSheets().filter(s => !CONFIG.PRESERVE_SHEET_NAMES.includes(s.getName()));

  const { dataMap, publishDateMap, sortedTimestamps } = aggregateVideoData_(videoSheets);
  if (sortedTimestamps.length === 0) {
    console.warn('⚠️ 比較シートに書き込むデータがありません。');
    return;
  }

  const tableValues = buildComparisonTable_(videoSheets, dataMap, publishDateMap, sortedTimestamps);

  renderComparisonSheet_(compSheet, tableValues);
  buildComparisonChart_(compSheet, tableValues.length, tableValues[0].length);
}

/**
 * 全動画シートのデータを集約して返す。
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} videoSheets
 * @returns {{ dataMap: object, publishDateMap: object, sortedTimestamps: string[] }}
 */
function aggregateVideoData_(videoSheets) {
  const dataMap         = {}; // { timestamp: { sheetName: viewCount } }
  const publishDateMap  = {}; // { sheetName: Date }
  const allTimestamps   = new Set();

  videoSheets.forEach(sh => {
    const name       = sh.getName();
    const pubDateVal = sh.getRange('A2').getValue();
    if (!pubDateVal) return;

    const pubDate = new Date(pubDateVal);
    publishDateMap[name] = pubDate;

    // 投稿時点(0回)を起点として追加
    const pubTs = formatTimestamp_(pubDate);
    allTimestamps.add(pubTs);
    setNestedValue_(dataMap, pubTs, name, 0);

    const lastRow = sh.getLastRow();
    if (lastRow < 4) return;

    sh.getRange(4, 1, lastRow - 3, 2).getValues().forEach(row => {
      if (!(row[0] instanceof Date)) return;
      const ts = formatTimestamp_(row[0]);
      allTimestamps.add(ts);
      setNestedValue_(dataMap, ts, name, row[1]);
    });
  });

  return {
    dataMap,
    publishDateMap,
    sortedTimestamps: [...allTimestamps].sort(),
  };
}

/**
 * 比較シート用のテーブルデータ（2次元配列）を構築する。
 * @returns {any[][]}
 */
function buildComparisonTable_(videoSheets, dataMap, publishDateMap, sortedTimestamps) {
  const now = new Date();

  const videoRows = videoSheets
    .map(sh => {
      const name        = sh.getName();
      const pubDate     = publishDateMap[name];
      if (!pubDate) return null;

      // 各タイムスタンプの値（null = データなし）
      const values = sortedTimestamps.map(ts => dataMap[ts]?.[name] ?? null);

      // 最新再生数（nullを除いた最後の値）
      const lastVal = [...values].reverse().find(v => v !== null) ?? 0;

      // 1日平均
      const diffDays  = Math.max(1, (now.getTime() - pubDate.getTime()) / MS_PER_DAY);
      const dailyAvg  = Math.round(lastVal / diffDays);

      // A1を参照式で動画タイトルを取得（シート名のシングルクォートをエスケープ）
      const escapedName = name.replace(/'/g, "''");
      const titleFormula = `='${escapedName}'!$A$1`;

      return { row: [titleFormula, dailyAvg, ...values], lastVal };
    })
    .filter(Boolean)
    .sort((a, b) => b.lastVal - a.lastVal);

  const headerRow = ['動画名', '1日平均再生数', ...sortedTimestamps.map(ts => new Date(ts))];
  return [headerRow, ...videoRows.map(r => r.row)];
}

/**
 * 比較シートにテーブルを書き込む。
 * データが 0 と空白しかない列（= まだ誰も再生数を持たないタイムスタンプ列）は非表示にする。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} compSheet
 * @param {any[][]} tableValues
 */
function renderComparisonSheet_(compSheet, tableValues) {
  // 既存グラフを削除してシートをクリア
  compSheet.getCharts().forEach(c => compSheet.removeChart(c));
  compSheet.clear();

  const rows = tableValues.length;
  const cols = tableValues[0].length;

  // 列が足りなければ追加
  if (compSheet.getMaxColumns() < cols) {
    compSheet.insertColumnsAfter(compSheet.getMaxColumns(), cols - compSheet.getMaxColumns());
  }

  // いったん全列を表示状態に戻す（前回の非表示をリセット）
  compSheet.showColumns(1, cols);

  compSheet.getRange(1, 1, rows, cols).setValues(tableValues);
  // ヘッダー行の日時列を書式設定
  compSheet.getRange(1, 3, 1, cols - 2).setNumberFormat('yyyy/MM/dd');
  // B列（1日平均）を強調
  compSheet.getRange(1, 2, rows, 1).setBackground('#fff2cc').setFontWeight('bold');
  // ヘッダー行を固定
  compSheet.setFrozenRows(1);

  // データ列（C列以降）のうち、全行が 0 または空白の列を非表示にする
  hideEmptyDataColumns_(compSheet, tableValues);
}

/**
 * C列以降（タイムスタンプ列）で、データ行がすべて 0 または null/空白の列を非表示にする。
 * A列（動画名）・B列（1日平均）は対象外。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} compSheet
 * @param {any[][]} tableValues  書き込み済みのテーブルデータ（2次元配列）
 */
function hideEmptyDataColumns_(compSheet, tableValues) {
  const rows = tableValues.length;
  const cols = tableValues[0].length;

  // データ行のみ（ヘッダー行=index 0 を除く）でチェック
  const dataRows = tableValues.slice(1);

  let hiddenCount = 0;

  for (let colIndex = 2; colIndex < cols; colIndex++) { // 0-indexed: 2 = C列
    const hasValue = dataRows.some(row => {
      const v = row[colIndex];
      return v !== null && v !== '' && v !== 0;
    });

    if (!hasValue) {
      // hideColumns は 1-indexed
      compSheet.hideColumns(colIndex + 1);
      hiddenCount++;
    }
  }

  if (hiddenCount > 0) {
    console.log(`👻 空列を非表示: ${hiddenCount} 列`);
  }
}

/**
 * 比較シートに折れ線グラフを追加する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} compSheet
 * @param {number} rows  テーブルの行数
 * @param {number} cols  テーブルの列数
 */
function buildComparisonChart_(compSheet, rows, cols) {
  const { WIDTH, HEIGHT } = CONFIG.CHART;

  // A列（動画名）と C列以降（タイムスタンプ別データ）を使用。B列（平均）は除外
  const chart = compSheet.newChart()
    .asLineChart()
    .addRange(compSheet.getRange(1, 1, rows, 1))   // 動画名
    .addRange(compSheet.getRange(1, 3, rows, cols - 2)) // 再生数データ
    .setTransposeRowsAndColumns(true)
    .setNumHeaders(1)
    .setPosition(rows + 2, 1, 0, 0)
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
  console.log(`📊 グラフを生成しました（${WIDTH}×${HEIGHT}px）`);
}

// ==========================================
// 8. ユーティリティ
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
// 9. 初期成長曲線の補完
// ==========================================
/**
 * 投稿日(0再生)と最初の計測値の間を、べき乗則曲線で補完する。
 * 2回目の計測が揃った時点で1度だけ実行し、ScriptProperties で管理する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Date} publishedAt
 */
function fillInitialGrowthCurve_(sheet, publishedAt) {
  const propKey = `curve_filled_${sheet.getName()}`;
  if (PropertiesService.getScriptProperties().getProperty(propKey)) return;

  const lastRow   = sheet.getLastRow();
  const dataCount = lastRow - 3;
  if (dataCount < 3) return; // [pub,0], [t1,v1], [t2,v2] の3行が必要

  const data = sheet.getRange(5, 1, 2, 2).getValues();
  const t0 = publishedAt;
  const [t1, v1] = [data[0][0], data[0][1]];
  const [t2, v2] = [data[1][0], data[1][1]];

  if (!(t1 instanceof Date) || !(t2 instanceof Date)) return;
  if (v1 <= 0 || v2 <= v1) return;

  const d1 = t1.getTime() - t0.getTime();
  const d2 = t2.getTime() - t0.getTime();
  if (d1 <= 0 || d2 <= d1) return;

  // べき乗則 v = C * d^alpha でフィット
  const alpha = Math.log(v2 / v1) / Math.log(d2 / d1);
  if (alpha <= 0 || alpha > 2) return; // 想定外の値はスキップ
  const C = v1 / Math.pow(d1, alpha);

  // [pub,0] と [t1,v1] の間に N 点を線形時間間隔で挿入
  const N = 8;
  const newRows = [];
  for (let i = 1; i < N; i++) {
    const d = d1 * (i / N);
    const v = Math.round(C * Math.pow(d, alpha));
    if (v <= 0) continue;
    newRows.push([new Date(t0.getTime() + d), v]);
  }
  if (newRows.length === 0) return;

  sheet.insertRowsAfter(4, newRows.length);
  sheet.getRange(5, 1, newRows.length, 2).setValues(newRows);

  PropertiesService.getScriptProperties().setProperty(propKey, 'true');
  console.log(`📈 初期成長曲線を補完 [${sheet.getName()}]: ${newRows.length} 点追加`);
}

// ==========================================
// 10. 管理用ユーティリティ（手動実行）
// ==========================================
/**
 * 動画シートをすべて削除してリセットする（PRESERVE_SHEET_NAMES は保持）。
 * ⚠️ データが失われるため慎重に使用すること。
 */
function resetSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets()
    .filter(sh => !CONFIG.PRESERVE_SHEET_NAMES.includes(sh.getName()))
    .forEach(sh => ss.deleteSheet(sh));
  console.log('🧹 動画シートをリセットしました。');
}

/**
 * 比較シートのみを再生成する（動画データシートは変更しない）。
 * グラフやレイアウトを修正したい場合に使用する。
 */
function rebuildComparisonSheet() {
  console.log('🔄 比較シートを再構築します...');
  updateComparisonSheet_(SpreadsheetApp.getActiveSpreadsheet());
  console.log('✅ 完了。');
}
