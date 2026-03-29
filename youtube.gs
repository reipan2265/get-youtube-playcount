// ==========================================
// YouTube API ラッパー
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
 * チャンネル内ランクを今回更新すべきか判定する。
 * 前回更新から 12 時間未満の場合は false を返してスキップする。
 * @returns {boolean}
 */
function shouldUpdateRank_() {
  const INTERVAL_MS = 12 * 60 * 60 * 1000;
  const props = PropertiesService.getScriptProperties();
  const last  = Number(props.getProperty('last_rank_update') || '0');
  if (Date.now() - last >= INTERVAL_MS) {
    return true;
  }
  console.log('チャンネル内ランク: 前回更新から 12 時間未満のためスキップ');
  return false;
}

/**
 * チャンネル内ランクの更新タイマーをリセットする。
 * 次回の main() 実行時に強制的に順位を再計算させたい場合に手動で実行する。
 */
function resetRankTimer() {
  PropertiesService.getScriptProperties().deleteProperty('last_rank_update');
  console.log('ランク更新タイマーをリセットしました。次回の main() 実行時に順位を再計算します。');
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
