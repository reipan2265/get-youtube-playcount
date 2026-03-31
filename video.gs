// ==========================================
// 動画ごとの処理・シート操作・サンプリング・個別グラフ
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
    // 順位列が追加される前に作成されたシートはヘッダーを補完
    if (!sheet.getRange('C3').getValue()) {
      sheet.getRange('C3').setValue('順位').setBackground('#eeeeee');
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

/**
 * 動画シートをチャンネル ID ごとにグループ化して返す。
 * B2 = channelId、C2 = channelTitle を使用する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} videoSheets
 * @returns {{ channelId: string, channelTitle: string, sheets: GoogleAppsScript.Spreadsheet.Sheet[] }[]}
 */
function groupSheetsByChannel_(videoSheets) {
  const groups = {};
  videoSheets.forEach(sh => {
    const channelId    = sh.getRange('B2').getValue() || '_unknown';
    const channelTitle = sh.getRange('C2').getValue() || channelId;
    if (!groups[channelId]) groups[channelId] = { channelTitle, sheets: [] };
    groups[channelId].sheets.push(sh);
  });
  return Object.entries(groups)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([channelId, { channelTitle, sheets }]) => ({ channelId, channelTitle, sheets }));
}

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

/**
 * 各動画シートの D4:H9 に直近の再生数増加量サマリーを書き込む。
 * データは降順ソート済みを前提とする（sortVideoSheetDescending_ 後に呼ぶこと）。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} currentViewCount  最新の再生数
 * @param {Date}   now               計測日時
 * @param {any[][]} allData          [[Date, viewCount], ...] 降順
 */
function updateGrowthSummary_(sheet, currentViewCount, now, allData) {
  if (!allData || allData.length === 0) return;

  const nowMs = now.getTime();

  function calcIncrease(fromMs, toMs) {
    const result = calcIncrease_(allData, currentViewCount, nowMs, fromMs, toMs);
    if (!result) return '---';
    return result.isEstimate ? `~${result.value}` : result.value;
  }

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
  const startRow = 4;
  sheet.getRange(startRow, 4, numRows, numCols).setValues(tableData);

  const fmtKey = `summary_fmt_v2_${sheet.getName()}`;
  if (!PropertiesService.getScriptProperties().getProperty(fmtKey)) {
    sheet.getRange(startRow, 4, 1, numCols).setBackground('#eeeeee').setFontWeight('bold');
    sheet.getRange(startRow + 1, 5, numRows - 1, numCols - 1).setNumberFormat('#,##0');
    PropertiesService.getScriptProperties().setProperty(fmtKey, 'true');
  }
}

/**
 * 投稿日(0再生)と最初の計測値の間を、べき乗則曲線で補完する。
 *
 * alpha の推定には最初の 1〜3 計測点を log-log 空間で重み付き最小二乗法により求める。
 * 重み [1, 2, 4] は後の計測点ほど大きく設定し、最新の勾配に合わせる。
 * 計測点が 1 点のみの場合は平方根曲線（alpha = 0.5）にフォールバックする。
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

/**
 * 「チャンネル内順位」シートに今回の順位計算結果を記録する。
 * 圧縮（runSampling_）に依存しない独立したシートで順位履歴を保持する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object<string, number>}  rankMap       { videoId: rank }
 * @param {Object<string, {title: string, channelTitle: string}>} metaMap  保存済みメタ情報
 * @param {Object<string, number>}  viewCountMap  { videoId: viewCount }（チャンネル取得時の再生数）
 * @param {Date} now
 */
function updateRankHistorySheet_(ss, rankMap, metaMap, viewCountMap, now) {
  if (Object.keys(rankMap).length === 0) return;

  let sheet = ss.getSheetByName(CONFIG.RANK_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.RANK_SHEET_NAME);
    sheet.getRange('A1:E1')
      .setValues([['日時', '動画タイトル', 'チャンネル', '再生数', '順位']])
      .setBackground('#eeeeee')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 400);
    sheet.setColumnWidth(3, 200);
    console.log(`シートを作成: ${CONFIG.RANK_SHEET_NAME}`);
  }

  const rows = Object.entries(rankMap)
    .filter(([, rank]) => rank != null)
    .sort(([, a], [, b]) => a - b) // 順位昇順
    .map(([id, rank]) => {
      const meta  = metaMap[id] || {};
      const title = meta.title        || id;
      const ch    = meta.channelTitle || '';
      const views = viewCountMap[id]  ?? '';
      return [now, title, ch, views, rank];
    });

  if (rows.length === 0) return;

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
  sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).setNumberFormat('#,##0');
  console.log(`チャンネル内順位シートに ${rows.length} 件記録`);
}
