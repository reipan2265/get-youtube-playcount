// ==========================================
// ユーティリティ
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
 * @param {Function} fn
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
