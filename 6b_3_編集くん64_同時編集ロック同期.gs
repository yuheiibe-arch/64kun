/**
 * ============================================================================
 * 6b_3_編集くん64_同時編集ロック同期.gs（6b系 3分割の3）
 * ・アクティブ作業者のロック取得（getActivePlayerLocks64_）
 * ・SpreadsheetApp を一切開かない超軽量ハートビート通信（heartbeatHenshukunPlayer）
 * ============================================================================
 */

function getActivePlayerLocks64_() {
  const cache = CacheService.getScriptCache();
  const raw = cache.get(HENSHUKUN_CONFIG.ACTIVE_USERS_CACHE_KEY);
  if (!raw) return {};
  try {
    const obj = JSON.parse(raw);
    const now = Date.now();
    const valid = {};
    Object.keys(obj).forEach(function(sessionId) {
      if (now - obj[sessionId].ts < HENSHUKUN_CONFIG.ACTIVE_USER_TTL_SEC * 1000) {
        valid[sessionId] = obj[sessionId];
      }
    });
    return valid;
  } catch (e) {
    return {};
  }
}

/**
 * 15秒おきのハートビートでは SpreadsheetApp を一切開かず、Cache / Properties のみ参照してシート競合タイムアウトを防ぐ
 */
function heartbeatHenshukunPlayer(sessionId, staffName) {
  const cache = CacheService.getScriptCache();
  const locks = getActivePlayerLocks64_();

  if (staffName) {
    for (const [sid, info] of Object.entries(locks)) {
      if (sid !== sessionId && info.staff === staffName) {
        return { ok: false, conflictStaff: staffName, activeLocks: locks };
      }
    }
    locks[sessionId] = { staff: staffName, ts: Date.now() };
  } else {
    delete locks[sessionId];
  }

  cache.put(HENSHUKUN_CONFIG.ACTIVE_USERS_CACHE_KEY, JSON.stringify(locks), 120);

  const props = PropertiesService.getScriptProperties();
  const docId = props.getProperty('SINGLE_MTG_DOC_ID');
  const latestDraft = (docId && typeof loadLatestDraftFastNoSheet64_ === 'function')
    ? loadLatestDraftFastNoSheet64_(docId)
    : null;

  return {
    ok: true,
    activeLocks: locks,
    latestDraft: latestDraft
  };
}