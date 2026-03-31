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
 * 動画メタ情報（channelId・タイトル等）を Script Properties に保存する。
 * main() 実行後に updateChannelRanks() が API 呼び出し不要でチャンネルを特定できるようにする。
 * @param {Object<string, object>} videoDataMap  { videoId: YouTubeVideoItem }
 */
function saveVideoMetadataToProps_(videoDataMap) {
  const meta = {};
  Object.entries(videoDataMap).forEach(([id, item]) => {
    meta[id] = {
      channelId:    item.snippet.channelId,
      title:        item.snippet.title,
      channelTitle: item.snippet.channelTitle,
    };
  });
  PropertiesService.getScriptProperties().setProperty('video_metadata', JSON.stringify(meta));
}

/**
 * Script Properties から動画メタ情報を読み込む。
 * @returns {Object<string, {channelId: string, title: string, channelTitle: string}>}
 */
function loadVideoMetadataFromProps_() {
  const raw = PropertiesService.getScriptProperties().getProperty('video_metadata');
  return raw ? JSON.parse(raw) : {};
}

/**
 * 計算済み rankMap を Script Properties に保存する。
 * main() が次回実行時に読み込み、動画シートの順位列に書き込む。
 * @param {Object<string, number>} rankMap  { videoId: rank }
 */
function saveRankMapToProps_(rankMap) {
  PropertiesService.getScriptProperties().setProperty('rank_map', JSON.stringify(rankMap));
}

/**
 * Script Properties から最後に計算した rankMap を読み込む。
 * @returns {Object<string, number>}
 */
function loadRankMapFromProps_() {
  const raw = PropertiesService.getScriptProperties().getProperty('rank_map');
  return raw ? JSON.parse(raw) : {};
}

/**
 * 保存済みメタ情報と追跡動画 ID リストから、チャンネルID → 動画IDリスト のマップを返す。
 * API 呼び出し不要。
 * @param {Object<string, {channelId: string}>} metaMap
 * @param {string[]} videoIds
 * @returns {Object<string, string[]>}  { channelId: [videoId, ...] }
 */
function buildChannelGroups_(metaMap, videoIds) {
  const groups = {};
  videoIds.forEach(id => {
    const cid = metaMap[id]?.channelId;
    if (!cid) return;
    if (!groups[cid]) groups[cid] = [];
    groups[cid].push(id);
  });
  return groups;
}

/**
 * チャンネルグループごとに全動画の再生数を1回だけ取得し、順位と再生数を返す。
 * @param {Object<string, string[]>} channelGroups  { channelId: [trackedVideoId, ...] }
 * @returns {{ rankMap: Object<string, number>, viewCountMap: Object<string, number> }}
 */
function computeRanksByChannelGroups_(channelGroups) {
  const rankMap      = {};
  const viewCountMap = {};

  Object.entries(channelGroups).forEach(([channelId, trackedIds]) => {
    console.log(`チャンネル ${channelId} の全動画を取得中...`);
    const allIds = fetchChannelVideoIds_(channelId);
    console.log(`チャンネル全動画: ${allIds.length} 本`);

    const viewCounts = fetchViewCountsOnly_(allIds);

    const sorted = allIds
      .filter(id => viewCounts[id] != null)
      .sort((a, b) => viewCounts[b] - viewCounts[a]);

    trackedIds.forEach(id => {
      const idx          = sorted.indexOf(id);
      rankMap[id]        = idx >= 0 ? idx + 1 : null;
      viewCountMap[id]   = viewCounts[id] ?? null;
    });

    console.log(`ランク算出完了: ${trackedIds.map(id => `${id}=${rankMap[id]}`).join(', ')}`);
  });

  return { rankMap, viewCountMap };
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
