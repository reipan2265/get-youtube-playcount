// ==========================================
// 比較シートの生成・グラフ
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
    !CONFIG.PRESERVE_SHEET_NAMES.includes(s.getName()) && !s.getName().startsWith('_') && !excludeSheets.has(s.getName())
  );

  const { dataMap, publishDateMap, sortedTimestamps } = aggregateVideoData_(videoSheets);
  if (sortedTimestamps.length === 0) return;

  const { tableValues } = buildComparisonTable_(videoSheets, dataMap, publishDateMap, sortedTimestamps);

  const rows = tableValues.length;
  const cols = tableValues[0].length;

  compSheet.getRange(1, 1, rows, cols).setValues(tableValues);
  compSheet.getRange(2, 2, rows - 1, 1).setNumberFormat('#,###').setFontWeight('bold');
  compSheet.getRange(2, 3, rows - 1, 1).setNumberFormat('0');
  SpreadsheetApp.flush();
  console.log('比較シートのテーブルを更新しました。');
}

/**
 * 全動画の再生数推移を1シートに集約し、比較グラフを生成する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Set<string>} [excludeSheets]
 */
function updateComparisonSheet_(ss, excludeSheets) {
  SpreadsheetApp.flush();
  Utilities.sleep(2000);

  if (!excludeSheets) excludeSheets = resolveWatchOnlySheetNames_(ss);

  const compSheet   = ss.getSheetByName(CONFIG.COMP_SHEET_NAME) || ss.insertSheet(CONFIG.COMP_SHEET_NAME, 0);
  const videoSheets = ss.getSheets().filter(s =>
    !CONFIG.PRESERVE_SHEET_NAMES.includes(s.getName()) && !s.getName().startsWith('_') && !excludeSheets.has(s.getName())
  );

  const { dataMap, publishDateMap, elapsedMaps, sortedTimestamps } = aggregateVideoData_(videoSheets);
  if (sortedTimestamps.length === 0) {
    console.warn('比較シートに書き込むデータがありません');
    return;
  }

  const { tableValues } = buildComparisonTable_(videoSheets, dataMap, publishDateMap, sortedTimestamps);
  renderComparisonSheet_(compSheet, tableValues);

  const { HEIGHT } = CONFIG.CHART;
  const channelGroups = groupSheetsByChannel_(videoSheets);
  let nextChartRow = tableValues.length + 2;

  channelGroups.forEach(({ channelId, channelTitle, sheets }, idx) => {
    const ch = aggregateVideoData_(sheets);
    if (ch.sortedTimestamps.length === 0) return;

    const chVideoRows = sheets.map(sh => {
      const name    = sh.getName();
      const lastVal = ch.sortedTimestamps.find(ts => ch.dataMap[ts]?.[name] != null);
      return { name, title: sh.getRange('A1').getValue() || name, lastVal: lastVal ? ch.dataMap[lastVal][name] : 0 };
    }).sort((a, b) => b.lastVal - a.lastVal);

    const chNames  = chVideoRows.map(r => r.name);
    const chTitles = chVideoRows.map(r => r.title);

    // ① 再生数推移（絶対時刻）
    const absHelper = getOrCreateHelperSheet_(ss, `_abs_${idx}`);
    const absInfo   = buildAbsoluteTimeHelperTable_(absHelper, chNames, chTitles, ch.dataMap, ch.sortedTimestamps);
    buildComparisonChart_(compSheet, absHelper, absInfo, nextChartRow, `再生数推移 — ${channelTitle}`);
    nextChartRow += Math.ceil(HEIGHT / 21) + 5;

    // ② 再生数推移（経過日数）
    const elapsedHelper = getOrCreateHelperSheet_(ss, `_elapsed_${idx}`);
    buildElapsedDaysChart_(compSheet, elapsedHelper, sheets, ch.publishDateMap, ch.elapsedMaps, nextChartRow, channelTitle);
    nextChartRow += Math.ceil(HEIGHT / 21) + 5;

    // ③ チャンネル内順位推移
    const rankHelper = getOrCreateHelperSheet_(ss, `_rank_${idx}`);
    buildChannelRankChart_(compSheet, rankHelper, sheets, nextChartRow, channelTitle);
    nextChartRow += Math.ceil(HEIGHT / 21) + 5;
  });
}

/**
 * 全動画シートのデータを集約して返す。
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} videoSheets
 * @returns {{ dataMap: object, publishDateMap: object, elapsedMaps: object, sortedTimestamps: string[] }}
 */
function aggregateVideoData_(videoSheets) {
  const dataMap        = {}; // { timestamp: { sheetName: viewCount } }
  const publishDateMap = {}; // { sheetName: Date }
  const elapsedMaps    = {}; // { sheetName: Map<elapsedHours, viewCount> }
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
    elapsed.set(0, 0);
    elapsedMaps[name] = elapsed;

    const lastRow = sh.getLastRow();
    if (lastRow < 4) return;

    sh.getRange(4, 1, lastRow - 3, 2).getValues().forEach(row => {
      if (!(row[0] instanceof Date)) return;
      const ts = formatTimestamp_(row[0]);
      allTimestamps.add(ts);
      setNestedValue_(dataMap, ts, name, row[1]);

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
    sortedTimestamps: [...allTimestamps].sort().reverse(),
  };
}

/**
 * 比較シート用のテーブルデータを構築する。
 * 動画行は最新再生数の降順でソートする。
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} videoSheets
 * @param {object}   dataMap
 * @param {object}   publishDateMap
 * @param {string[]} sortedTimestamps
 * @returns {{ tableValues: any[][], sortedNames: string[], sortedTitles: string[] }}
 */
function buildComparisonTable_(videoSheets, dataMap, publishDateMap, sortedTimestamps) {
  const now   = new Date();
  const nowMs = now.getTime();

  const videoRows = videoSheets
    .map(sh => {
      const name    = sh.getName();
      const pubDate = publishDateMap[name];
      if (!pubDate) return null;

      const allData = sortedTimestamps
        .filter(ts => dataMap[ts]?.[name] != null)
        .map(ts => [new Date(ts), dataMap[ts][name]]);

      const lastVal          = allData.length > 0 ? allData[0][1] : 0;
      const daysSincePublish = Math.floor((nowMs - pubDate.getTime()) / MS_PER_DAY);

      const fmt = (fromMs) =>
        formatIncreaseWithRate_(calcIncrease_(allData, lastVal, nowMs, fromMs, 0), lastVal);

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
        fullTitle: sh.getRange('A1').getValue() || name,
      };
    })
    .filter(Boolean)
    .sort((a, b) => b.lastVal - a.lastVal);

  const headerRow = ['動画名', '最新再生数', '経過日数', '24h増加', '7日増加', '30日増加'];
  return {
    tableValues:  [headerRow, ...videoRows.map(r => r.row)],
    sortedNames:  videoRows.map(r => r.name),
    sortedTitles: videoRows.map(r => r.fullTitle),
  };
}

/**
 * 比較シートにテーブルを書き込む。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} compSheet
 * @param {any[][]} tableValues
 */
function renderComparisonSheet_(compSheet, tableValues) {
  compSheet.getCharts().forEach(c => compSheet.removeChart(c));
  compSheet.clear();

  const rows = tableValues.length;
  const cols = tableValues[0].length;

  const maxCols = compSheet.getMaxColumns();
  if (maxCols > cols) {
    compSheet.deleteColumns(cols + 1, maxCols - cols);
  }

  compSheet.getRange(1, 1, rows, cols).setValues(tableValues);
  compSheet.getRange(1, 1, 1, cols).setBackground('#eeeeee').setFontWeight('bold');
  compSheet.getRange(2, 2, rows - 1, 1).setNumberFormat('#,###').setFontWeight('bold');
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
 * ヘルパーシートに絶対日時テーブルを書き込む。グラフのデータ源として使用する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} helperSheet
 * @param {string[]} sortedNames
 * @param {string[]} sortedTitles
 * @param {object}   dataMap
 * @param {string[]} sortedTimestamps
 * @returns {{ numRows: number, numCols: number }}
 */
function buildAbsoluteTimeHelperTable_(helperSheet, sortedNames, sortedTitles, dataMap, sortedTimestamps) {
  const ascTimestamps = [...sortedTimestamps].reverse();

  const headerRow = ['日時', ...sortedTitles];
  const dataRows  = ascTimestamps.map(ts => [
    new Date(ts),
    ...sortedNames.map(name => dataMap[ts]?.[name] ?? null),
  ]);
  const tableData = [headerRow, ...dataRows];

  const numRows = tableData.length;
  const numCols = tableData[0].length;

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
 * @param {GoogleAppsScript.Spreadsheet.Sheet} compSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} helperSheet
 * @param {{ numRows: number, numCols: number }} helperInfo
 * @param {number} startRow
 * @param {string} title
 */
function buildComparisonChart_(compSheet, helperSheet, helperInfo, startRow, title) {
  const { WIDTH, HEIGHT } = CONFIG.CHART;
  const { numRows, numCols } = helperInfo;

  const chart = compSheet.newChart()
    .asLineChart()
    .addRange(helperSheet.getRange(1, 1, numRows, numCols))
    .setNumHeaders(1)
    .setPosition(startRow, 1, 0, 0)
    .setOption('title', title || '再生数推移')
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

/**
 * 全動画の横軸を「投稿日からの経過日数」に正規化した折れ線グラフを生成する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet}   compSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet}   helperSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} videoSheets
 * @param {object} publishDateMap
 * @param {object} elapsedMaps
 * @param {number} startRow
 * @param {string} channelTitle
 */
function buildElapsedDaysChart_(compSheet, helperSheet, videoSheets, publishDateMap, elapsedMaps, startRow, channelTitle) {
  const validSheets = videoSheets.filter(sh => publishDateMap[sh.getName()]);
  if (validSheets.length === 0) return;

  const videoMaps = validSheets.map(sh => elapsedMaps[sh.getName()] || new Map());

  const allHoursSet = new Set();
  videoMaps.forEach(map => map.forEach((_, h) => allHoursSet.add(h)));
  const sortedHours = [...allHoursSet].sort((a, b) => a - b);

  const titleRow = ['経過日数', ...validSheets.map(sh => sh.getRange('A1').getValue() || sh.getName())];
  const dataRows = sortedHours.map(h => [
    Math.round(h / 24 * 100) / 100,
    ...videoMaps.map(map => map.has(h) ? map.get(h) : null),
  ]);
  const tableData = [titleRow, ...dataRows];

  const numRows = tableData.length;
  const numCols = tableData[0].length;

  if (helperSheet.getMaxRows() < numRows) {
    helperSheet.insertRowsAfter(helperSheet.getMaxRows(), numRows - helperSheet.getMaxRows());
  }
  if (helperSheet.getMaxColumns() < numCols) {
    helperSheet.insertColumnsAfter(helperSheet.getMaxColumns(), numCols - helperSheet.getMaxColumns());
  }
  helperSheet.getRange(1, 1, numRows, numCols).setValues(tableData);

  const { WIDTH, HEIGHT } = CONFIG.CHART;
  const chart = compSheet.newChart()
    .asLineChart()
    .addRange(helperSheet.getRange(1, 1, numRows, numCols))
    .setNumHeaders(1)
    .setPosition(startRow, 1, 0, 0)
    .setOption('title', `再生数推移（投稿日起点）— ${channelTitle || ''}`)
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
 * チャンネル内順位の推移グラフを生成する。
 * Y 軸は逆転（direction:-1）し、順位 1 がグラフ上部に表示される。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} compSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} rankHelperSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} channelSheets
 * @param {number} startRow
 * @param {string} channelTitle
 */
function buildChannelRankChart_(compSheet, rankHelperSheet, channelSheets, startRow, channelTitle) {
  const { WIDTH, HEIGHT } = CONFIG.CHART;
  const allTimestamps = new Set();
  const rankDataMap   = {};

  channelSheets.forEach(sh => {
    const name    = sh.getName();
    const lastRow = sh.getLastRow();
    if (lastRow < 4) return;
    sh.getRange(4, 1, lastRow - 3, 3).getValues().forEach(row => {
      if (!(row[0] instanceof Date) || row[2] === '' || row[2] == null) return;
      const ts = formatTimestamp_(row[0]);
      allTimestamps.add(ts);
      setNestedValue_(rankDataMap, ts, name, Number(row[2]));
    });
  });

  if (allTimestamps.size === 0) {
    console.log(`チャンネル ${channelTitle}: 順位データなしのためグラフをスキップ`);
    return;
  }

  const sheetNames  = channelSheets.map(sh => sh.getName());
  const sheetTitles = channelSheets.map(sh => sh.getRange('A1').getValue() || sh.getName());
  const ascTs       = [...allTimestamps].sort();

  const tableData = [
    ['日時', ...sheetTitles],
    ...ascTs.map(ts => [new Date(ts), ...sheetNames.map(n => rankDataMap[ts]?.[n] ?? null)]),
  ];

  const numRows = tableData.length;
  const numCols = tableData[0].length;
  if (rankHelperSheet.getMaxRows() < numRows) rankHelperSheet.insertRowsAfter(rankHelperSheet.getMaxRows(), numRows - rankHelperSheet.getMaxRows());
  if (rankHelperSheet.getMaxColumns() < numCols) rankHelperSheet.insertColumnsAfter(rankHelperSheet.getMaxColumns(), numCols - rankHelperSheet.getMaxColumns());
  rankHelperSheet.getRange(1, 1, numRows, numCols).setValues(tableData);

  compSheet.insertChart(
    compSheet.newChart()
      .asLineChart()
      .addRange(rankHelperSheet.getRange(1, 1, numRows, numCols))
      .setNumHeaders(1)
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', `チャンネル内順位推移 — ${channelTitle}（値が小さいほど上位）`)
      .setOption('width', WIDTH)
      .setOption('height', HEIGHT)
      .setOption('interpolateNulls', true)
      .setOption('pointSize', 2)
      .setOption('lineWidth', 2)
      .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
      .setOption('chartArea', { left: '6%', top: '10%', width: '65%', height: '75%' })
      .setOption('hAxis', { slantedText: true, slantedTextAngle: 30, textStyle: { fontSize: 9 } })
      .setOption('vAxis', { format: '#,##0', direction: -1, gridlines: { color: '#b0b0b0' }, minorGridlines: { count: 4, color: '#e8e8e8' } })
      .build()
  );
  console.log(`チャンネル内順位グラフを生成しました（${channelTitle}、${channelSheets.length} 動画）`);
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
