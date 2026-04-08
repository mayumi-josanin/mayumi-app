// ===== GAS設定 =====
// ↓ GASウェブアプリURLをここに貼り付け ↓
const GAS_URL = 'https://script.google.com/macros/s/AKfycbzf3iBSe2IFIeJJgaGxd4_MeFVErRnKdS2Y9C4xkPA1d6If5dgKhm-rjRAwqtYE6CotCA/exec';
const CURRENT_WEB_BUNDLE_VERSION = '2026.04.06.63';
const APP_RUNTIME_CONFIG_STORAGE_KEY = 'mayumi_app_runtime_config';
const DEFAULT_APP_RUNTIME_CONFIG = Object.freeze({
  latestAppVersion: '1.1.1',
  minimumSupportedVersion: '0.0.0',
  iosStoreUrl: '',
  updateTitle: 'アップデートが必要です',
  updateMessage: 'このアプリを引き続き利用するには、最新版へアップデートしてください。',
  webBundleVersion: CURRENT_WEB_BUNDLE_VERSION
});
let currentAppRuntimeConfig = { ...DEFAULT_APP_RUNTIME_CONFIG };
let currentInstalledAppInfo = { version: '', build: '', bundleId: '', isNative: false, source: 'unknown' };
let currentRequiredUpdateUrl = '';

let USER_MENUS = [];
let _profile = null;
try { _profile = JSON.parse(localStorage.getItem('mayumi_profile') || 'null'); } catch (e) { }
let CUSTOMER_NAME = _profile && _profile.name ? _profile.name : '';
let stampCount = 0;
try { stampCount = parseInt(localStorage.getItem('mayumi_stamp') || '0') || 0; } catch (e) { }
let stampCardNum = 1;
try { stampCardNum = parseInt(localStorage.getItem('mayumi_stamp_card') || '1') || 1; } catch (e) { }
let STAMP_REWARD_CONFIG = [];
let CURRENT_MONTHLY_REWARD = null;
const STAMP_HISTORY_STORAGE_KEY = 'mayumi_stamp_history';
const LAST_STAMP_AT_STORAGE_KEY = 'mayumi_last_stamp_at';
let EARNED_REWARDS = [];
try { EARNED_REWARDS = JSON.parse(localStorage.getItem('mayumi_earned_rewards') || '[]'); } catch (e) { }
let STAMP_HISTORY = [];
try { STAMP_HISTORY = JSON.parse(localStorage.getItem(STAMP_HISTORY_STORAGE_KEY) || '[]'); } catch (e) { }
let cart = [], orders = [];
let isOrderSubmitting = false;
let isCancelSubmitting = false;
let isReceiptSubmitting = false;
let selectedPay = null, currentProdIdx = null, modalQty = 1;
let cancelOrderId = null;
let receiptSubmittingOrderId = null;
let blogItems = [];
let allBlogCategories = [];
let pushNotices = [];
let isDataLoaded = false;
let supportFaqItems = [];
let supportChatHistory = [];
let isSupportChatSending = false;
try { supportChatHistory = JSON.parse(localStorage.getItem('mayumi_support_chat_history') || '[]'); } catch (e) { }
let lastSupportTopic = '';
if (supportChatHistory.length) {
  try { lastSupportTopic = localStorage.getItem('mayumi_support_chat_topic') || ''; } catch (e) { }
}
let activeAppDialogResolver = null;
let lastSyncedRewardStatus = null;
let isRewardGachaDrawing = false;
let lastRewardGachaResult = null;
let calendarData = [];
let calendarLoaded = false;
let currentMonthDate = new Date();
let selectedDate = new Date();
const PASSCODE_SET_STORAGE_KEY = 'mayumi_passcode_set';
const LOCAL_PASSCODE_HASH_STORAGE_KEY = 'mayumi_local_passcode_hash';
const PASSCODE_LOGIN_ENABLED_STORAGE_KEY = 'mayumi_passcode_login_enabled';
const PASSCODE_SKIP_ONCE_SESSION_KEY = 'mayumi_passcode_skip_once';
const UPDATE_RESTORE_PAGE_SESSION_KEY = 'mayumi_update_restore_page';
const UPDATE_RELOAD_TARGET_SESSION_KEY = 'mayumi_update_reload_target';
const PASSCODE_RESUME_LOCK_DELAY_MS = 1500;
const SECURE_STORE_DB_NAME = 'mayumi_secure_store';
const SECURE_STORE_NAME = 'kv';
const SECURE_STORE_VERSION = 1;
const SECURE_PASSCODE_HASH_KEY = 'local_passcode_hash';
const DEVICE_SESSION_ID_STORAGE_KEY = 'mayumi_device_session_id';
const RETRY_QUEUE_STORAGE_KEY = 'mayumi_retry_queue';
const FAVORITES_STORAGE_KEY = 'mayumi_favorites';
const ITEM_SEEN_STORAGE_KEY = 'mayumi_seen_items';
const ITEM_SEEN_BASELINE_STORAGE_KEY = 'mayumi_seen_items_initialized';
const ACCESSIBILITY_STORAGE_KEY = 'mayumi_accessibility';
const UPDATE_BANNER_DISMISSED_AT_STORAGE_KEY = 'mayumi_update_banner_dismissed_at';
const APP_INSTALL_CONTEXT_STORAGE_KEY = 'mayumi_install_context';
const DEVICE_SESSIONS_CACHE_STORAGE_KEY = 'mayumi_device_sessions_cache';
const RETRY_QUEUE_MAX = 30;
const RETRYABLE_ACTIONS = {
  order: true,
  updateUser: true,
  syncUserRewardStatus: true,
  syncUserDeviceSession: true,
  removeUserDeviceSession: true,
  unsubscribePush: true
};
let isPasscodeAuthenticated = false;
let initAppStarted = false;
let appHiddenAt = 0;
let currentUserDevices = [];
let currentDeviceSessionId = '';
let secureStoreDbPromise = null;
let cachedLocalPasscodeHash = '';
let retryQueueBusy = false;
let appUpdateContext = {
  needsReload: false,
  waitingServiceWorker: false,
  message: '',
  webBundleVersion: '',
  reloadStarted: false,
  reloadTimerId: 0
};
let itemSeenState = readJsonStorage(ITEM_SEEN_STORAGE_KEY, {});
let favoriteEntries = readJsonStorage(FAVORITES_STORAGE_KEY, []);
let accessibilitySettings = Object.assign({
  textSize: 'standard',
  highContrast: false
}, readJsonStorage(ACCESSIBILITY_STORAGE_KEY, {}));
const REWARD_GACHA_PRIZE_POOL = [
  { key: 'A', rankLabel: 'A賞', rewardName: 'A賞プレゼント', capsuleColor: '#f5cb6c', accentColor: '#b0791b', message: '当日のおたのしみプレゼントをご用意しています。', weight: 10 },
  { key: 'B', rankLabel: 'B賞', rewardName: 'B賞プレゼント', capsuleColor: '#f3b7c9', accentColor: '#b86282', message: '当日のおたのしみプレゼントをご用意しています。', weight: 20 },
  { key: 'C', rankLabel: 'C賞', rewardName: 'C賞プレゼント', capsuleColor: '#b9d8a7', accentColor: '#628f58', message: '当日のおたのしみプレゼントをご用意しています。', weight: 30 },
  { key: 'D', rankLabel: 'D賞', rewardName: 'D賞プレゼント', capsuleColor: '#b9d9f3', accentColor: '#547fa2', message: '当日のおたのしみプレゼントをご用意しています。', weight: 40 }
];

function readJsonStorage(key, fallbackValue) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return fallbackValue;
    const parsed = JSON.parse(raw);
    return parsed == null ? fallbackValue : parsed;
  } catch (e) {
    return fallbackValue;
  }
}

function writeJsonStorage(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch (e) { }
}

function removeStoredValue(key) {
  try {
    localStorage.removeItem(key);
  } catch (e) { }
}

function readSessionValue(key) {
  try {
    return sessionStorage.getItem(key) || '';
  } catch (e) {
    return '';
  }
}

function writeSessionValue(key, value) {
  try {
    sessionStorage.setItem(key, String(value || ''));
  } catch (e) { }
}

function removeSessionValue(key) {
  try {
    sessionStorage.removeItem(key);
  } catch (e) { }
}

function openSecureStoreDb() {
  if (!('indexedDB' in window)) return Promise.resolve(null);
  if (secureStoreDbPromise) return secureStoreDbPromise;
  secureStoreDbPromise = new Promise(function (resolve) {
    try {
      const request = indexedDB.open(SECURE_STORE_DB_NAME, SECURE_STORE_VERSION);
      request.onupgradeneeded = function () {
        const db = request.result;
        if (!db.objectStoreNames.contains(SECURE_STORE_NAME)) {
          db.createObjectStore(SECURE_STORE_NAME);
        }
      };
      request.onsuccess = function () {
        resolve(request.result);
      };
      request.onerror = function () {
        resolve(null);
      };
    } catch (e) {
      resolve(null);
    }
  });
  return secureStoreDbPromise;
}

async function secureStoreGet(key) {
  const db = await openSecureStoreDb();
  if (!db) return '';
  return new Promise(function (resolve) {
    try {
      const tx = db.transaction(SECURE_STORE_NAME, 'readonly');
      const store = tx.objectStore(SECURE_STORE_NAME);
      const request = store.get(key);
      request.onsuccess = function () {
        resolve(String(request.result || ''));
      };
      request.onerror = function () {
        resolve('');
      };
    } catch (e) {
      resolve('');
    }
  });
}

async function secureStoreSet(key, value) {
  const db = await openSecureStoreDb();
  if (!db) return false;
  return new Promise(function (resolve) {
    try {
      const tx = db.transaction(SECURE_STORE_NAME, 'readwrite');
      tx.objectStore(SECURE_STORE_NAME).put(String(value || ''), key);
      tx.oncomplete = function () { resolve(true); };
      tx.onerror = function () { resolve(false); };
    } catch (e) {
      resolve(false);
    }
  });
}

async function initSecureLocalStore() {
  const storedHash = await secureStoreGet(SECURE_PASSCODE_HASH_KEY);
  if (storedHash) {
    cachedLocalPasscodeHash = storedHash;
    return cachedLocalPasscodeHash;
  }
  try {
    const legacyHash = String(localStorage.getItem(LOCAL_PASSCODE_HASH_STORAGE_KEY) || '');
    if (legacyHash) {
      cachedLocalPasscodeHash = legacyHash;
      await secureStoreSet(SECURE_PASSCODE_HASH_KEY, legacyHash);
      localStorage.removeItem(LOCAL_PASSCODE_HASH_STORAGE_KEY);
    }
  } catch (e) { }
  return cachedLocalPasscodeHash;
}

function getCurrentPlatformName() {
  try {
    if (window.Capacitor && typeof Capacitor.getPlatform === 'function') {
      return Capacitor.getPlatform();
    }
  } catch (e) { }
  if (window.Capacitor && Capacitor.Plugins && Capacitor.Plugins.PushNotifications) {
    return 'ios';
  }
  return 'web';
}

function isLikelyCapacitorRuntime() {
  const href = window.location && window.location.href ? String(window.location.href) : '';
  const ua = navigator && navigator.userAgent ? String(navigator.userAgent) : '';
  // capacitor:// は確実にネイティブアプリ。Capacitor文字列はUAに含まれる場合がある。
  // window.Capacitor が存在し、かつ window.Capacitor.isNative が true なら確実に実機。
  return href.startsWith('capacitor://') ||
    ua.includes('Capacitor') ||
    (window.Capacitor && window.Capacitor.isNative === true);
}

async function waitForCapacitorPushPlugin(timeoutMs) {
  if (!isLikelyCapacitorRuntime()) return null;
  const deadline = Date.now() + Math.max(0, timeoutMs || 0);
  do {
    if (window.Capacitor && Capacitor.Plugins && Capacitor.Plugins.PushNotifications) {
      return Capacitor.Plugins.PushNotifications;
    }
    if (Date.now() >= deadline) break;
    await new Promise(function (resolve) { setTimeout(resolve, 100); });
  } while (true);
  return null;
}

function isNativeAppRuntime() {
  return getCurrentPlatformName() !== 'web';
}

function getCameraPermissionDeviceLabel() {
  const ua = navigator && navigator.userAgent ? String(navigator.userAgent) : '';
  const platform = getCurrentPlatformName();
  if (platform === 'ios' || /iPhone|iPad|iPod/i.test(ua)) return 'iPhone';
  if (platform === 'android' || /Android/i.test(ua)) return 'Android';
  return '端末';
}

function getCameraPermissionTargetLabel() {
  if (isLikelyCapacitorRuntime()) {
    return 'アプリ';
  }
  const ua = navigator && navigator.userAgent ? String(navigator.userAgent) : '';
  if (/iPhone|iPad|iPod/i.test(ua)) return 'Safari';
  if (/Android/i.test(ua)) return 'Chrome';
  return 'ブラウザ';
}

async function getCameraPermissionState() {
  try {
    if (!navigator.permissions || typeof navigator.permissions.query !== 'function') {
      return 'unknown';
    }
    const status = await navigator.permissions.query({ name: 'camera' });
    if (status && typeof status.state === 'string') {
      return status.state;
    }
  } catch (e) {
    console.log('camera permission query unsupported:', e);
  }
  return 'unknown';
}

async function openCameraPermissionSettings() {
  try {
    const appPlugin = window.Capacitor && Capacitor.Plugins ? Capacitor.Plugins.App : null;
    if (appPlugin && typeof appPlugin.openSettings === 'function') {
      await appPlugin.openSettings();
      return true;
    }
  } catch (e) {
    console.error('openCameraPermissionSettings App.openSettings error:', e);
  }

  const ua = navigator && navigator.userAgent ? String(navigator.userAgent) : '';
  const canAttemptScheme = isLikelyCapacitorRuntime() || /iPhone|iPad|iPod|Android/i.test(ua);
  if (!canAttemptScheme) return false;

  try {
    window.location.href = 'app-settings:';
    return true;
  } catch (e) {
    console.error('openCameraPermissionSettings app-settings error:', e);
  }
  return false;
}

async function showCameraPermissionRecoveryDialog(permissionState) {
  const deviceLabel = getCameraPermissionDeviceLabel();
  const targetLabel = getCameraPermissionTargetLabel();
  const isDenied = permissionState === 'denied';

  if (!isDenied) {
    return showAppConfirm(
      'スタンプ取得にはカメラの許可が必要です。\nこのあと表示される確認画面で「許可」を選んでください。',
      {
        title: 'カメラの許可が必要です',
        confirmLabel: '許可を確認する',
        cancelLabel: '閉じる'
      }
    );
  }

  const shouldOpenSettings = await showAppConfirm(
    deviceLabel + 'でカメラの許可がオフになっています。\n「設定を開く」を押して、' + targetLabel + 'のカメラを許可してください。',
    {
      title: 'カメラの許可がオフです',
      confirmLabel: '設定を開く',
      cancelLabel: '閉じる'
    }
  );

  if (!shouldOpenSettings) return false;

  const opened = await openCameraPermissionSettings();
  if (opened) {
    await showAppAlert(
      '設定画面を開いています。\nカメラを許可したあと、もう一度「カメラを起動して読み取る」を押してください。',
      {
        title: '設定で許可してください',
        confirmLabel: 'OK'
      }
    );
    return false;
  }

  await showAppAlert(
    targetLabel + 'または' + deviceLabel + 'の設定でカメラを許可したあと、もう一度お試しください。',
    {
      title: 'カメラの許可が必要です',
      confirmLabel: 'OK'
    }
  );
  return false;
}

function getVersionParts(version) {
  const parts = String(version || '')
    .split(/[^0-9]+/)
    .filter(Boolean)
    .map(function (part) { return parseInt(part, 10) || 0; });
  return parts.length ? parts : [0];
}

function compareVersionStrings(left, right) {
  const a = getVersionParts(left);
  const b = getVersionParts(right);
  const maxLength = Math.max(a.length, b.length);
  for (let i = 0; i < maxLength; i++) {
    const leftPart = a[i] || 0;
    const rightPart = b[i] || 0;
    if (leftPart > rightPart) return 1;
    if (leftPart < rightPart) return -1;
  }
  return 0;
}

function normalizePasscodeInput(value) {
  return String(value == null ? '' : value).trim();
}

function normalizePhoneInput(value) {
  return String(value == null ? '' : value).replace(/\D/g, '');
}

function normalizeNameInput(value) {
  return String(value == null ? '' : value).trim();
}

function normalizeDateOnlyInput(value) {
  return String(value == null ? '' : value).trim();
}

function getUserActivityProfilePayload() {
  const profile = _profile || {};
  return {
    name: normalizeNameInput(profile.name || CUSTOMER_NAME || ''),
    phone: normalizePhoneInput(profile.phone || ''),
    birthday: normalizeDateOnlyInput(profile.birthday || ''),
    address: String(profile.address || '').trim()
  };
}

function normalizeTransferCodeInput(value) {
  return String(value == null ? '' : value).replace(/\D/g, '').slice(0, 8);
}

function isValidPasscodeValue(value) {
  return /^(?:\d{4}|\d{6})$/.test(normalizePasscodeInput(value));
}

function isValidTransferCodeValue(value) {
  return /^\d{8}$/.test(normalizeTransferCodeInput(value));
}

function getStoredLocalPasscodeHash() {
  if (cachedLocalPasscodeHash) return cachedLocalPasscodeHash;
  try {
    return String(localStorage.getItem(LOCAL_PASSCODE_HASH_STORAGE_KEY) || '');
  } catch (e) {
    return '';
  }
}

function getStoredPasscodeLoginPreference() {
  try {
    const value = String(localStorage.getItem(PASSCODE_LOGIN_ENABLED_STORAGE_KEY) || '').trim();
    return value === 'true' || value === 'false' ? value : '';
  } catch (e) {
    return '';
  }
}

function hasConfiguredLocalPasscode() {
  return !!getStoredLocalPasscodeHash();
}

function ensurePasscodeLoginPreference(defaultEnabled) {
  const saved = getStoredPasscodeLoginPreference();
  if (saved) return saved === 'true';
  const nextEnabled = defaultEnabled !== false;
  try {
    localStorage.setItem(PASSCODE_LOGIN_ENABLED_STORAGE_KEY, nextEnabled ? 'true' : 'false');
  } catch (e) { }
  return nextEnabled;
}

function isPasscodeLoginEnabled() {
  if (!_profile || !hasConfiguredLocalPasscode()) return false;
  return ensurePasscodeLoginPreference(true);
}

function setPasscodeLoginEnabled(enabled) {
  const nextEnabled = enabled !== false;
  try {
    localStorage.setItem(PASSCODE_LOGIN_ENABLED_STORAGE_KEY, nextEnabled ? 'true' : 'false');
  } catch (e) { }
  updatePasscodeLoginUI();
  return nextEnabled;
}

function needsRequiredPasscodeSetup() {
  return !!(_profile && !hasConfiguredLocalPasscode());
}

function fallbackHashString(text) {
  let hash = 2166136261;
  for (let i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return 'fallback:' + String(hash >>> 0).padStart(10, '0');
}

async function hashPasscodeValue(passcode) {
  const source = 'mayumi-lock:' + normalizePasscodeInput(passcode);
  if (window.crypto && window.crypto.subtle && typeof TextEncoder !== 'undefined') {
    try {
      const encoded = new TextEncoder().encode(source);
      const digest = await window.crypto.subtle.digest('SHA-256', encoded);
      return Array.from(new Uint8Array(digest)).map(function (byte) {
        return byte.toString(16).padStart(2, '0');
      }).join('');
    } catch (e) { }
  }
  return fallbackHashString(source);
}

async function storeLocalPasscode(passcode) {
  const hash = await hashPasscodeValue(passcode);
  cachedLocalPasscodeHash = hash;
  try {
    localStorage.setItem(PASSCODE_SET_STORAGE_KEY, 'true');
  } catch (e) { }
  await secureStoreSet(SECURE_PASSCODE_HASH_KEY, hash);
  removeStoredValue(LOCAL_PASSCODE_HASH_STORAGE_KEY);
  ensurePasscodeLoginPreference(true);
  return hash;
}

async function verifyLocalPasscode(passcode) {
  const storedHash = getStoredLocalPasscodeHash();
  if (!storedHash) return false;
  return storedHash === await hashPasscodeValue(passcode);
}

function markPasscodeUnlockSkippedOnce() {
  try {
    sessionStorage.setItem(PASSCODE_SKIP_ONCE_SESSION_KEY, 'true');
  } catch (e) { }
}

function consumePasscodeUnlockSkippedOnce() {
  try {
    if (sessionStorage.getItem(PASSCODE_SKIP_ONCE_SESSION_KEY) === 'true') {
      sessionStorage.removeItem(PASSCODE_SKIP_ONCE_SESSION_KEY);
      return true;
    }
  } catch (e) { }
  return false;
}

function getInstallModeLabel() {
  const isStandalone = window.matchMedia && window.matchMedia('(display-mode: standalone)').matches;
  const isIosStandalone = window.navigator && window.navigator.standalone === true;
  if (isLikelyCapacitorRuntime()) return 'アプリ';
  if (isStandalone || isIosStandalone) return 'ホーム画面';
  return 'ブラウザ';
}

function getCurrentDeviceSessionId() {
  if (currentDeviceSessionId) return currentDeviceSessionId;
  try {
    currentDeviceSessionId = String(localStorage.getItem(DEVICE_SESSION_ID_STORAGE_KEY) || '').trim();
  } catch (e) { }
  if (currentDeviceSessionId) return currentDeviceSessionId;
  currentDeviceSessionId = (window.crypto && typeof window.crypto.randomUUID === 'function')
    ? window.crypto.randomUUID()
    : ('device-' + Date.now() + '-' + Math.random().toString(36).slice(2, 10));
  try {
    localStorage.setItem(DEVICE_SESSION_ID_STORAGE_KEY, currentDeviceSessionId);
  } catch (e) { }
  return currentDeviceSessionId;
}

function buildCurrentDeviceLabel() {
  const ua = navigator && navigator.userAgent ? String(navigator.userAgent) : '';
  const deviceKind = /iPhone/i.test(ua) ? 'iPhone'
    : /iPad/i.test(ua) ? 'iPad'
      : /Android/i.test(ua) ? 'Android'
        : /Macintosh/i.test(ua) ? 'Mac'
          : /Windows/i.test(ua) ? 'Windows'
            : 'この端末';
  return deviceKind + ' / ' + getInstallModeLabel();
}

function buildCurrentDeviceSessionPayload() {
  return {
    deviceId: getCurrentDeviceSessionId(),
    label: buildCurrentDeviceLabel(),
    platform: getCurrentPlatformName(),
    appVersion: CURRENT_WEB_BUNDLE_VERSION,
    passcodeEnabled: !!(_profile && hasConfiguredLocalPasscode() && isPasscodeLoginEnabled()),
    pushEnabled: isPushEnabled()
  };
}

function cacheDeviceSessions(devices) {
  currentUserDevices = Array.isArray(devices) ? devices.slice() : [];
  writeJsonStorage(DEVICE_SESSIONS_CACHE_STORAGE_KEY, currentUserDevices);
}

function getCachedDeviceSessions() {
  if (Array.isArray(currentUserDevices) && currentUserDevices.length) return currentUserDevices.slice();
  currentUserDevices = readJsonStorage(DEVICE_SESSIONS_CACHE_STORAGE_KEY, []);
  return currentUserDevices.slice();
}

function getRetryQueue() {
  return readJsonStorage(RETRY_QUEUE_STORAGE_KEY, []).filter(function (entry) {
    return entry && entry.payload && entry.payload.type;
  });
}

function saveRetryQueue(queue) {
  writeJsonStorage(RETRY_QUEUE_STORAGE_KEY, (queue || []).slice(0, RETRY_QUEUE_MAX));
  renderRetryQueueStatus();
}

function getRetryPayloadKey(payload) {
  const action = String(payload && payload.type || '');
  if (!action) return '';
  if (action === 'order') return action + ':' + String(payload.orderId || '');
  if (action === 'updateUser') return action + ':' + String(payload.memberId || '');
  if (action === 'syncUserRewardStatus') return action + ':' + String(payload.memberId || '');
  if (action === 'syncUserDeviceSession' || action === 'removeUserDeviceSession') {
    return action + ':' + String(payload.memberId || '') + ':' + String(payload.deviceId || '');
  }
  if (action === 'unsubscribePush') return action + ':' + String(payload.memberId || '') + ':' + String(payload.playerId || payload.pushToken || '');
  return action + ':' + JSON.stringify(payload);
}

function enqueueRetryPayload(payload) {
  const key = getRetryPayloadKey(payload);
  if (!key) return;
  const queue = getRetryQueue();
  const existingIndex = queue.findIndex(function (entry) {
    return entry.key === key;
  });
  const nextEntry = {
    key: key,
    payload: payload,
    createdAt: Date.now(),
    retryCount: existingIndex >= 0 ? Number(queue[existingIndex].retryCount || 0) : 0
  };
  if (existingIndex >= 0) queue.splice(existingIndex, 1, nextEntry);
  else queue.unshift(nextEntry);
  saveRetryQueue(queue);
}

function isRetryableAction(action) {
  return !!RETRYABLE_ACTIONS[String(action || '').trim()];
}

async function flushRetryQueue() {
  if (retryQueueBusy || !navigator.onLine) return;
  const queue = getRetryQueue();
  if (!queue.length) {
    renderRetryQueueStatus();
    return;
  }
  retryQueueBusy = true;
  try {
    let remaining = queue.slice();
    for (let i = 0; i < queue.length; i++) {
      const entry = queue[i];
      const result = await postToGAS(entry.payload, {
        skipRetryQueue: true,
        silent: true
      });
      if (result && (result.status === 'ok' || result.duplicate === true)) {
        remaining = remaining.filter(function (queued) {
          return queued.key !== entry.key;
        });
        saveRetryQueue(remaining);
      } else {
        remaining = remaining.map(function (queued) {
          if (queued.key !== entry.key) return queued;
          queued.retryCount = Number(queued.retryCount || 0) + 1;
          return queued;
        });
        saveRetryQueue(remaining);
        break;
      }
    }
  } finally {
    retryQueueBusy = false;
    renderRetryQueueStatus();
  }
}

function getItemSeenMap() {
  return itemSeenState && typeof itemSeenState === 'object' ? itemSeenState : {};
}

function saveItemSeenMap(map) {
  itemSeenState = map || {};
  writeJsonStorage(ITEM_SEEN_STORAGE_KEY, itemSeenState);
}

function buildContentItemKey(kind, item) {
  if (item && item.sourceKey) return String(item.sourceKey);
  return [
    String(kind || '').trim(),
    String(item && (item.rowIdx || item.originalIndex || item.memberId || item.id || item.name || item.title) || '').trim(),
    String(item && (item.updatedAt || item.date || item.publishAt || item.time || '') || '').trim()
  ].join('::');
}

function getContentItemTimestamp(item) {
  return parseLooseDateToTimestamp(item && (item.updatedAt || item.date || item.publishAt || item.time || ''));
}

function isContentItemUnread(kind, item) {
  const key = buildContentItemKey(kind, item);
  if (!key) return false;
  const seenAt = Number(getItemSeenMap()[key] || 0);
  const updatedAt = getContentItemTimestamp(item);
  return updatedAt > 0 && seenAt < updatedAt;
}

function markContentItemSeen(kind, item) {
  const key = buildContentItemKey(kind, item);
  if (!key) return;
  const map = Object.assign({}, getItemSeenMap());
  map[key] = Math.max(Date.now(), getContentItemTimestamp(item) || 0);
  saveItemSeenMap(map);
}

function ensureSeenBaselineInitialized() {
  try {
    if (localStorage.getItem(ITEM_SEEN_BASELINE_STORAGE_KEY) === 'true') return false;
  } catch (e) { }
  const map = Object.assign({}, getItemSeenMap());
  (blogItems || []).forEach(function (item) { map[buildContentItemKey('blog', item)] = getContentItemTimestamp(item) || Date.now(); });
  (calendarData || []).forEach(function (item) { map[buildContentItemKey('calendar', item)] = getContentItemTimestamp(item) || Date.now(); });
  (PRODUCTS || []).forEach(function (item) { map[buildContentItemKey('product', item)] = getContentItemTimestamp(item) || Date.now(); });
  (USER_MENUS || []).forEach(function (item) { map[buildContentItemKey('menu', item)] = getContentItemTimestamp(item) || Date.now(); });
  saveItemSeenMap(map);
  try {
    localStorage.setItem(ITEM_SEEN_BASELINE_STORAGE_KEY, 'true');
  } catch (e) { }
  return true;
}

function getFavoriteEntries() {
  if (!Array.isArray(favoriteEntries)) favoriteEntries = [];
  return favoriteEntries;
}

function saveFavoriteEntries(entries) {
  favoriteEntries = Array.isArray(entries) ? entries.slice(0, 100) : [];
  writeJsonStorage(FAVORITES_STORAGE_KEY, favoriteEntries);
  renderFavoriteList();
}

function isFavoriteKey(key) {
  return getFavoriteEntries().some(function (entry) {
    return entry && entry.key === key;
  });
}

function toggleFavoriteEntry(entry) {
  if (!entry || !entry.key) return false;
  const current = getFavoriteEntries();
  const next = current.filter(function (item) {
    return item.key !== entry.key;
  });
  const existed = next.length !== current.length;
  if (!existed) {
    next.unshift(Object.assign({
      savedAt: Date.now()
    }, entry));
  }
  saveFavoriteEntries(next);
  return !existed;
}

function getTextSizeScale() {
  if (accessibilitySettings.textSize === 'large') return 'large';
  if (accessibilitySettings.textSize === 'xlarge') return 'xlarge';
  return 'standard';
}

function applyAccessibilitySettings() {
  document.body.classList.toggle('high-contrast', accessibilitySettings.highContrast === true);
  document.body.dataset.textScale = getTextSizeScale();
  writeJsonStorage(ACCESSIBILITY_STORAGE_KEY, accessibilitySettings);
  renderAccessibilitySettings();
}

function cycleTextSizeSetting() {
  const order = ['standard', 'large', 'xlarge'];
  const current = order.indexOf(getTextSizeScale());
  accessibilitySettings.textSize = order[(current + 1) % order.length];
  applyAccessibilitySettings();
}

function toggleHighContrastSetting() {
  accessibilitySettings.highContrast = accessibilitySettings.highContrast !== true;
  applyAccessibilitySettings();
}

function renderAccessibilitySettings() {
  const sizeLabel = document.getElementById('textSizeStatus');
  const contrastLabel = document.getElementById('highContrastStatus');
  if (sizeLabel) {
    sizeLabel.textContent = accessibilitySettings.textSize === 'xlarge' ? '大きめ' : accessibilitySettings.textSize === 'large' ? '少し大きめ' : '標準';
  }
  if (contrastLabel) {
    contrastLabel.textContent = accessibilitySettings.highContrast ? 'オン' : 'オフ';
  }
}

function showAppUpdateBanner(message, webBundleVersion) {
  appUpdateContext.needsReload = true;
  appUpdateContext.message = message || '最新版があります。再読み込みして最新の内容を反映します。';
  appUpdateContext.webBundleVersion = webBundleVersion || '';
  setTimeout(function () {
    applyPendingAppUpdate().catch(function (e) {
      console.error('applyPendingAppUpdate auto error:', e);
    });
  }, 50);
}

function hideAppUpdateBanner(persistDismiss) {
  if (persistDismiss) {
    try {
      localStorage.setItem(UPDATE_BANNER_DISMISSED_AT_STORAGE_KEY, String(Date.now()));
    } catch (e) { }
  }
}

function normalizeAppPageName(value) {
  const raw = String(value || '').trim().toLowerCase();
  const allowed = {
    home: 'home',
    'menu-list': 'menu-list',
    menu: 'menu-list',
    shop: 'shop',
    blog: 'blog',
    news: 'blog',
    cart: 'cart',
    calendar: 'calendar',
    mypage: 'mypage',
    profile: 'mypage',
    notices: 'notices',
    notice: 'notices'
  };
  return allowed[raw] || '';
}

function getActivePageName() {
  const activePage = document.querySelector('.page.active');
  if (!activePage || !activePage.id) return 'home';
  return normalizeAppPageName(activePage.id.replace('page-', '')) || 'home';
}

function activatePageSilently(name) {
  const pageName = normalizeAppPageName(name) || 'home';
  document.querySelectorAll('.page').forEach(function (page) {
    page.classList.remove('active');
  });
  document.querySelectorAll('.nav-btn').forEach(function (button) {
    button.classList.remove('active');
  });
  const targetPage = document.getElementById('page-' + pageName);
  if (targetPage) targetPage.classList.add('active');
  const targetNav = document.getElementById('nav-' + pageName);
  if (targetNav) targetNav.classList.add('active');
}

function getPendingUpdateRestorePage() {
  return normalizeAppPageName(readSessionValue(UPDATE_RESTORE_PAGE_SESSION_KEY));
}

function setPendingUpdateRestorePage(pageName) {
  const normalized = normalizeAppPageName(pageName);
  if (!normalized) return '';
  writeSessionValue(UPDATE_RESTORE_PAGE_SESSION_KEY, normalized);
  return normalized;
}

function clearPendingUpdateRestorePage() {
  removeSessionValue(UPDATE_RESTORE_PAGE_SESSION_KEY);
}

function buildAppUpdateReloadUrl(targetPage) {
  const url = new URL(window.location.href);
  const normalizedPage = normalizeAppPageName(targetPage);
  url.searchParams.set('upd', Date.now());
  if (normalizedPage) {
    url.searchParams.set('restorePage', normalizedPage);
  } else {
    url.searchParams.delete('restorePage');
  }
  return url.toString();
}

function setPendingAppReloadTarget(url, pageName) {
  if (url) writeSessionValue(UPDATE_RELOAD_TARGET_SESSION_KEY, url);
  setPendingUpdateRestorePage(pageName);
}

function getPendingAppReloadTarget() {
  return readSessionValue(UPDATE_RELOAD_TARGET_SESSION_KEY);
}

function clearPendingAppReloadTarget() {
  removeSessionValue(UPDATE_RELOAD_TARGET_SESSION_KEY);
}

function getPreferredStartupPage() {
  try {
    const params = new URLSearchParams(window.location.search);
    const openPage = normalizeOpenPageTarget(params.get('open'));
    if (openPage) return openPage;
    const restorePage = normalizeAppPageName(params.get('restorePage'));
    if (restorePage) return restorePage;
  } catch (e) { }
  return getPendingUpdateRestorePage() || 'home';
}

async function applyPendingAppUpdate() {
  if (appUpdateContext.reloadStarted) return;
  appUpdateContext.reloadStarted = true;
  appUpdateContext.needsReload = true;
  const restorePage = setPendingUpdateRestorePage(getActivePageName()) || 'home';
  const targetUrl = buildAppUpdateReloadUrl(restorePage);
  setPendingAppReloadTarget(targetUrl, restorePage);
  if ('serviceWorker' in navigator) {
    try {
      const registrations = await navigator.serviceWorker.getRegistrations();
      let waiting = null;
      registrations.forEach(function (registration) {
        if (registration.waiting) waiting = registration.waiting;
      });
      if (waiting) {
        appUpdateContext.waitingServiceWorker = true;
        waiting.postMessage({ type: 'SKIP_WAITING' });
        if (appUpdateContext.reloadTimerId) {
          clearTimeout(appUpdateContext.reloadTimerId);
        }
        appUpdateContext.reloadTimerId = setTimeout(function () {
          window.location.href = targetUrl;
        }, 1200);
        return;
      }
    } catch (e) {
      console.error('applyPendingAppUpdate waiting worker error:', e);
    }
  }
  window.location.href = targetUrl;
}

function renderRetryQueueStatus() {
  const count = getRetryQueue().length;
  const el = document.getElementById('retryQueueStatus');
  if (!el) return;
  if (!count) {
    el.innerHTML = '<b>同期状況:</b> すべて最新です';
    return;
  }
  el.innerHTML = '<b>同期状況:</b> ' + count + '件の送信待ちがあります。通信が戻ると自動で再送します。';
}

function renderCurrentDeviceGuidance() {
  const el = document.getElementById('deviceGuidanceContent');
  if (!el) return;
  const platform = getCurrentPlatformName();
  const installMode = getInstallModeLabel();
  let lines = [];
  if (platform === 'ios') {
    lines = [
      'iPhone版をご利用中です。カメラや通知が反応しない時は、iPhoneの「設定」からまゆみ助産院アプリの権限をご確認ください。',
      '再インストールや機種変更の前に、マイページで引き継ぎコードを発行しておくと復元しやすくなります。'
    ];
  } else if (platform === 'android') {
    lines = [
      'Android版をご利用中です。カメラや通知が反応しない時は、Androidの「設定」からアプリの権限と電池最適化をご確認ください。',
      'ホーム画面から開いている場合でも、まず以前登録した方の復元を使うと会員IDの重複を防げます。'
    ];
  } else if (installMode === 'ホーム画面') {
    lines = [
      'ホーム画面追加版をご利用中です。端末設定でSafariまたはChromeのカメラ・通知権限が必要です。',
      '再インストールや端末変更の際は、新規登録ではなく復元を選んでください。'
    ];
  } else {
    lines = [
      'ブラウザ版をご利用中です。カメラや通知はSafariまたはChromeの権限設定に影響されます。',
      'よく使う場合はホーム画面に追加すると、次回以降すぐ開けます。'
    ];
  }
  el.innerHTML = '<div class="device-guidance-chip">' + escapeHtml(buildCurrentDeviceLabel()) + '</div><ul class="device-guidance-list">' + lines.map(function (line) {
    return '<li>' + escapeHtml(line) + '</li>';
  }).join('') + '</ul>';
}

function triggerConfetti() {
  const container = document.createElement('div');
  container.className = 'confetti-container';
  document.body.appendChild(container);

  const colors = ['#a2b4a1', '#e8dace', '#d4a373', '#ccd5ae', '#faedcd', '#fefae0', '#e9edc9'];
  const count = 80;

  for (let i = 0; i < count; i++) {
    const piece = document.createElement('div');
    piece.className = 'confetti-piece';

    // ランダムな配置
    const left = Math.random() * 100 + '%';
    const delay = Math.random() * 3 + 's';
    const color = colors[Math.floor(Math.random() * colors.length)];
    const size = (Math.random() * 6 + 4) + 'px';
    const rotation = Math.random() * 360 + 'deg';

    piece.style.left = left;
    piece.style.animationDelay = delay;
    piece.style.backgroundColor = color;
    piece.style.width = size;
    piece.style.height = size;
    piece.style.transform = `rotate(${rotation})`;

    container.appendChild(piece);
  }

  // 演出が終わったら要素を削除
  setTimeout(function () {
    if (container.parentNode) {
      container.parentNode.removeChild(container);
    }
  }, 6000);
}

function updateStampModalPresentation(isMilestone) {
  const title = document.getElementById('stampModalTitle');
  const wrap = document.getElementById('stampPopWrap');
  const badge = document.getElementById('stampAchievementBadge');
  const icon = document.getElementById('stampPopIcon');
  const note = document.getElementById('stampCelebrateNote');
  const actionBtn = document.getElementById('stampMilestoneActionBtn');
  if (!title || !wrap || !badge || !icon || !note) return;

  const rewardDrawn = hasCurrentCardReward();

  title.textContent = isMilestone ? 'スタンプ10個達成です！' : 'スタンプを取得しました！';
  wrap.classList.toggle('milestone', !!isMilestone);
  badge.style.display = isMilestone ? 'inline-flex' : 'none';
  note.style.display = isMilestone ? 'block' : 'none';
  note.innerHTML = isMilestone
    ? (rewardDrawn
      ? 'おめでとうございます。<br>このカードのガチャ結果は保存済みです。次のスタンプカードを始められます。'
      : 'おめでとうございます。<br>特典ガチャを回して、ごほうびを受け取ってください。')
    : 'おめでとうございます。<br>マイページの「スタンプ・特典履歴」で特典をご確認いただけます。';
  icon.textContent = isMilestone ? '🎉' : '🌿';
  icon.classList.toggle('milestone', !!isMilestone);
  if (actionBtn) {
    actionBtn.style.display = isMilestone ? 'block' : 'none';
    actionBtn.textContent = rewardDrawn ? '🌸 新しいスタンプカードを取得' : '🎯 特典ガチャを回す';
  }
}

function sanitizeAppRuntimeConfig(raw) {
  const input = raw && typeof raw === 'object' ? raw : {};
  const config = {
    latestAppVersion: String(input.latestAppVersion || DEFAULT_APP_RUNTIME_CONFIG.latestAppVersion).trim() || DEFAULT_APP_RUNTIME_CONFIG.latestAppVersion,
    minimumSupportedVersion: String(input.minimumSupportedVersion || DEFAULT_APP_RUNTIME_CONFIG.minimumSupportedVersion).trim() || DEFAULT_APP_RUNTIME_CONFIG.minimumSupportedVersion,
    iosStoreUrl: String(input.iosStoreUrl || DEFAULT_APP_RUNTIME_CONFIG.iosStoreUrl).trim(),
    updateTitle: String(input.updateTitle || DEFAULT_APP_RUNTIME_CONFIG.updateTitle).trim() || DEFAULT_APP_RUNTIME_CONFIG.updateTitle,
    updateMessage: String(input.updateMessage || DEFAULT_APP_RUNTIME_CONFIG.updateMessage).trim() || DEFAULT_APP_RUNTIME_CONFIG.updateMessage,
    webBundleVersion: String(input.webBundleVersion || DEFAULT_APP_RUNTIME_CONFIG.webBundleVersion).trim() || DEFAULT_APP_RUNTIME_CONFIG.webBundleVersion
  };
  if (compareVersionStrings(config.latestAppVersion, config.minimumSupportedVersion) < 0) {
    config.latestAppVersion = config.minimumSupportedVersion;
  }
  if (config.iosStoreUrl && !/^https?:\/\//i.test(config.iosStoreUrl)) {
    config.iosStoreUrl = '';
  }
  return config;
}

function readCachedAppRuntimeConfig() {
  try {
    const raw = localStorage.getItem(APP_RUNTIME_CONFIG_STORAGE_KEY);
    if (!raw) return null;
    return sanitizeAppRuntimeConfig(JSON.parse(raw));
  } catch (e) {
    return null;
  }
}

function cacheAppRuntimeConfig(config) {
  try {
    localStorage.setItem(APP_RUNTIME_CONFIG_STORAGE_KEY, JSON.stringify(sanitizeAppRuntimeConfig(config)));
  } catch (e) { }
}

function withTimeout(promise, timeoutMs) {
  return Promise.race([
    promise,
    new Promise(function (resolve) {
      setTimeout(function () { resolve(null); }, timeoutMs);
    })
  ]);
}

async function fetchAppRuntimeConfig() {
  const cached = readCachedAppRuntimeConfig();
  try {
    const res = await withTimeout(getFromGAS('getAppRuntimeConfig'), 4000);
    if (res && res.status === 'ok' && res.config) {
      currentAppRuntimeConfig = sanitizeAppRuntimeConfig(res.config);
      cacheAppRuntimeConfig(currentAppRuntimeConfig);
      return currentAppRuntimeConfig;
    }
  } catch (e) {
    console.error('fetchAppRuntimeConfig error:', e);
  }
  currentAppRuntimeConfig = cached || { ...DEFAULT_APP_RUNTIME_CONFIG };
  return currentAppRuntimeConfig;
}

async function getInstalledAppInfo() {
  const isNative = isNativeAppRuntime();
  const base = {
    version: isNative ? '0.0.0' : DEFAULT_APP_RUNTIME_CONFIG.latestAppVersion,
    build: '',
    bundleId: '',
    isNative: isNative,
    source: isNative ? 'fallback' : 'web'
  };
  if (!isNative) {
    currentInstalledAppInfo = base;
    return base;
  }
  try {
    if (window.Capacitor && Capacitor.Plugins && Capacitor.Plugins.NativeAppInfo &&
      typeof Capacitor.Plugins.NativeAppInfo.getInfo === 'function') {
      const info = await Capacitor.Plugins.NativeAppInfo.getInfo();
      currentInstalledAppInfo = {
        version: String(info && info.version || '').trim() || '0.0.0',
        build: String(info && info.build || '').trim(),
        bundleId: String(info && info.bundleId || '').trim(),
        isNative: true,
        source: 'native-plugin'
      };
      return currentInstalledAppInfo;
    }
  } catch (e) {
    console.error('NativeAppInfo getInfo failed:', e);
  }
  currentInstalledAppInfo = base;
  return base;
}

function fillRequiredUpdateModal(config, appInfo) {
  const title = document.getElementById('requiredUpdateTitle');
  const message = document.getElementById('requiredUpdateMessage');
  const meta = document.getElementById('requiredUpdateMeta');
  const primary = document.getElementById('requiredUpdatePrimaryBtn');
  const secondary = document.getElementById('requiredUpdateSecondaryBtn');
  const latestLabel = config.latestAppVersion ? '\n最新版: ' + config.latestAppVersion : '';
  title.textContent = config.updateTitle || DEFAULT_APP_RUNTIME_CONFIG.updateTitle;
  message.textContent = config.updateMessage || DEFAULT_APP_RUNTIME_CONFIG.updateMessage;
  meta.textContent =
    '現在のアプリ版: ' + (appInfo.version || '不明') +
    '\n必要なアプリ版: ' + (config.minimumSupportedVersion || '未設定') +
    latestLabel;
  currentRequiredUpdateUrl = config.iosStoreUrl || '';
  if (currentRequiredUpdateUrl) {
    primary.textContent = 'アップデートする';
    secondary.style.display = 'none';
  } else {
    primary.textContent = '閉じる';
    secondary.style.display = 'none';
  }
}

function openRequiredUpdateModal() {
  const modal = document.getElementById('requiredUpdateModal');
  if (modal) modal.classList.add('open');
}

function closeRequiredUpdateModal() {
  const modal = document.getElementById('requiredUpdateModal');
  if (modal) modal.classList.remove('open');
}

function openRequiredUpdateLink() {
  if (!currentRequiredUpdateUrl) {
    closeRequiredUpdateModal();
    return;
  }
  window.location.href = currentRequiredUpdateUrl;
}

async function ensureSupportedAppVersion() {
  const config = await fetchAppRuntimeConfig();
  const appInfo = await getInstalledAppInfo();
  currentAppRuntimeConfig = config;
  currentInstalledAppInfo = appInfo;

  // 1. ネイティブアプリの強制アップデート判定
  const needsNativeUpdate = appInfo.isNative &&
    compareVersionStrings(appInfo.version || '0.0.0', config.minimumSupportedVersion || '0.0.0') < 0;

  // 2. Webプログラム（HTML/JS）の更新判定
  // GAS側の指定バージョンのほうが新しい時だけ更新を要求する
  const needsWebUpdate = compareVersionStrings(config.webBundleVersion || '0.0.0', CURRENT_WEB_BUNDLE_VERSION) > 0;

  if (needsNativeUpdate && config.iosStoreUrl) {
    fillRequiredUpdateModal(config, appInfo);
    openRequiredUpdateModal();
    return { blocked: true, needsWebUpdate: needsWebUpdate, config: config, appInfo: appInfo };
  }

  closeRequiredUpdateModal();
  return { blocked: false, needsWebUpdate: needsWebUpdate, config: config, appInfo: appInfo };
}

// GASからデータ取得（doGet対応）
// action: 'getNews' | 'getProducts'
async function getFromGAS(action, params) {
  if (!GAS_URL || GAS_URL === 'YOUR_GAS_URL_HERE') return null;
  try {
    let url = GAS_URL + '?action=' + action + '&t=' + Date.now();
    if (params) url += '&data=' + encodeURIComponent(JSON.stringify(params));
    const res = await fetch(url);
    return await res.json();
  } catch (e) { return null; }
}

// GASへデータ送信（GET/POST自動切り替え）
// payload: { type:'order|updateUser|uploadImage|...', ... }
async function postToGAS(payload, options) {
  if (!GAS_URL || GAS_URL === 'YOUR_GAS_URL_HERE') return;
  const opts = options || {};
  try {
    const action = payload.type;
    let res;

    // 認証系や更新系は必ずPOSTで送る
    if (
      action === 'updateUser' ||
      action === 'uploadImage' ||
      action === 'askSupportChat' ||
      action === 'syncUserRewardStatus' ||
      action === 'syncUserDeviceSession' ||
      action === 'removeUserDeviceSession' ||
      action === 'unsubscribePush' ||
      action === 'drawRewardGacha' ||
      action === 'recoverAccount' ||
      action === 'resetForgottenPasscode' ||
      action === 'issueTransferCode'
    ) {
      res = await fetch(GAS_URL, {
        method: 'POST',
        body: JSON.stringify(payload),
        headers: { 'Content-Type': 'text/plain' }, // CORS回避のため text/plain
        redirect: 'follow'
      });
    } else {
      // それ以外（注文など）は従来通りGETで送る（POST不達問題の回避）
      const data = encodeURIComponent(JSON.stringify(payload));
      const url = GAS_URL + '?action=' + action + '&data=' + data + '&t=' + Date.now();
      res = await fetch(url);
    }

    const json = await res.json();
    console.log('postToGAS response:', json);
    if (opts.skipRetryQueue !== true && (!json || json.status === 'error') && isRetryableAction(action)) {
      enqueueRetryPayload(payload);
      return {
        status: 'queued',
        queued: true,
        message: '通信が不安定なため、送信待ちにしました。'
      };
    }
    if (json && (json.status === 'ok' || json.duplicate === true) && opts.skipRetryQueue !== true) {
      flushRetryQueue().catch(function (error) {
        console.error('flushRetryQueue after success error:', error);
      });
    }
    return json;
  } catch (e) {
    console.log('postToGAS error:', e);
    if (opts.skipRetryQueue !== true && isRetryableAction(payload && payload.type)) {
      enqueueRetryPayload(payload);
      if (!opts.silent) {
        showToast('通信が不安定なため、送信待ちにしました');
      }
      return {
        status: 'queued',
        queued: true,
        message: '通信が不安定なため、送信待ちにしました。'
      };
    }
    return null;
  }
}

function getTodayStampKey() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function normalizeStampDateKey(value) {
  if (!value) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(value))) {
    return String(value);
  }
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return '';
  const year = parsed.getFullYear();
  const month = String(parsed.getMonth() + 1).padStart(2, '0');
  const day = String(parsed.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function normalizeRewardDateTime(value) {
  if (!value) return '';
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return '';
  return parsed.toISOString();
}

function normalizeSingleReward(reward, index) {
  const earnedDate = normalizeRewardDateTime(reward && reward.earnedDate) || new Date().toISOString();
  let expiryDate = normalizeRewardDateTime(reward && reward.expiryDate);
  if (!expiryDate) {
    const expiry = new Date(earnedDate);
    expiry.setMonth(expiry.getMonth() + 1);
    expiryDate = expiry.toISOString();
  }
  return {
    id: reward && reward.id ? String(reward.id) : `reward-${index + 1}`,
    cardNum: Math.max(1, Number(reward && reward.cardNum) || 1),
    rewardName: reward && reward.rewardName ? String(reward.rewardName) : '特典プレゼント',
    earnedDate: earnedDate,
    expiryDate: expiryDate,
    used: reward && reward.used === true,
    usedAt: normalizeRewardDateTime(reward && reward.usedAt)
  };
}

function normalizeRewardList(rewards) {
  if (!Array.isArray(rewards)) return [];
  return rewards.map(function (reward, index) {
    return normalizeSingleReward(reward, index);
  });
}

function normalizeSingleStampHistoryEntry(entry, index) {
  const cardNum = Math.max(1, Number(entry && entry.cardNum) || 1);
  const stampNumber = Math.max(1, Math.min(10, Number(entry && (entry.stampNumber !== undefined ? entry.stampNumber : entry.stampCount)) || 1));
  const acquiredDate = normalizeRewardDateTime(entry && (entry.acquiredDate || entry.earnedDate || entry.date)) || new Date().toISOString();
  return {
    id: entry && entry.id ? String(entry.id) : `stamp-${cardNum}-${stampNumber}-${index + 1}`,
    cardNum: cardNum,
    stampNumber: stampNumber,
    acquiredDate: acquiredDate,
    dateKey: normalizeStampDateKey(entry && (entry.dateKey || entry.acquiredDate || entry.earnedDate || entry.date))
  };
}

function normalizeStampHistoryList(history) {
  if (!Array.isArray(history)) return [];
  const merged = {};
  history.forEach(function (entry, index) {
    const normalized = normalizeSingleStampHistoryEntry(entry, index);
    const key = `${normalized.cardNum}:${normalized.stampNumber}`;
    if (!merged[key] || new Date(normalized.acquiredDate).getTime() >= new Date(merged[key].acquiredDate).getTime()) {
      merged[key] = normalized;
    }
  });
  return Object.keys(merged).map(function (key) {
    return merged[key];
  }).sort(function (a, b) {
    const diff = new Date(b.acquiredDate).getTime() - new Date(a.acquiredDate).getTime();
    if (diff !== 0) return diff;
    if (Number(b.cardNum) !== Number(a.cardNum)) return Number(b.cardNum) - Number(a.cardNum);
    return Number(b.stampNumber) - Number(a.stampNumber);
  });
}

function selectLatestRewardDateTime(values) {
  return (Array.isArray(values) ? values : []).reduce(function (latest, value) {
    const normalized = normalizeRewardDateTime(value);
    if (!normalized) return latest;
    if (!latest) return normalized;
    return new Date(normalized).getTime() >= new Date(latest).getTime() ? normalized : latest;
  }, '');
}

function getLatestStampHistoryDateTime(history) {
  const normalizedHistory = normalizeStampHistoryList(history);
  return normalizedHistory.length ? normalizedHistory[0].acquiredDate : '';
}

function deriveLastStampAt(statusLike, normalizedHistory) {
  const source = statusLike || {};
  const history = normalizedHistory || normalizeStampHistoryList(source.stampHistory);
  return selectLatestRewardDateTime([
    source.lastStampAt,
    history.length ? history[0].acquiredDate : '',
    source.lastStampDate
  ]);
}

function mergeRewardLists(leftRewards, rightRewards) {
  const merged = {};
  normalizeRewardList([].concat(leftRewards || [], rightRewards || [])).forEach(function (reward) {
    const key = [String(reward.id || ''), String(reward.cardNum || ''), String(reward.rewardName || ''), String(reward.earnedDate || '')].join('|');
    if (!merged[key]) {
      merged[key] = reward;
      return;
    }
    if (reward.used && !merged[key].used) {
      merged[key] = reward;
    }
  });
  return Object.keys(merged).map(function (key) {
    return merged[key];
  }).sort(function (a, b) {
    return new Date(b.earnedDate).getTime() - new Date(a.earnedDate).getTime();
  });
}

function mergeRewardStatuses(primaryStatus, secondaryStatus) {
  const primary = getComparableRewardStatus(primaryStatus);
  const secondary = getComparableRewardStatus(secondaryStatus);
  const mergedStampHistory = normalizeStampHistoryList([].concat(primary.stampHistory || [], secondary.stampHistory || []));
  const mergedLastStampAt = selectLatestRewardDateTime([
    deriveLastStampAt(primary, primary.stampHistory),
    deriveLastStampAt(secondary, secondary.stampHistory),
    mergedStampHistory.length ? mergedStampHistory[0].acquiredDate : ''
  ]);
  return {
    stampCount: Math.max(Number(primary.stampCount || 0), Number(secondary.stampCount || 0)),
    stampCardNum: Math.max(Number(primary.stampCardNum || 1), Number(secondary.stampCardNum || 1)),
    rewards: mergeRewardLists(primary.rewards, secondary.rewards),
    stampHistory: mergedStampHistory,
    lastStampDate: normalizeStampDateKey(primary.lastStampDate || secondary.lastStampDate || mergedLastStampAt),
    lastStampAt: mergedLastStampAt,
    stampAchievedDate: [primary.stampAchievedDate, secondary.stampAchievedDate].sort().pop() || ''
  };
}

function recordStampHistoryEntry(cardNum, stampNumber, acquiredDate) {
  STAMP_HISTORY = normalizeStampHistoryList([{
    id: `stamp-${cardNum}-${stampNumber}`,
    cardNum: Math.max(1, Number(cardNum) || 1),
    stampNumber: Math.max(1, Math.min(10, Number(stampNumber) || 1)),
    acquiredDate: normalizeRewardDateTime(acquiredDate) || new Date().toISOString()
  }].concat(STAMP_HISTORY));
  try {
    localStorage.setItem(STAMP_HISTORY_STORAGE_KEY, JSON.stringify(STAMP_HISTORY));
  } catch (e) { }
  return STAMP_HISTORY;
}

function getRewardGachaPrizeMeta(rewardName) {
  const normalized = String(rewardName || '').trim();
  const matched = REWARD_GACHA_PRIZE_POOL.find(function (prize) {
    return prize.rewardName === normalized ||
      prize.rankLabel === normalized ||
      (normalized && normalized.indexOf(prize.rankLabel) === 0);
  });
  if (matched) {
    return Object.assign({}, matched);
  }
  return {
    key: 'SPECIAL',
    rankLabel: 'ごほうび獲得',
    rewardName: normalized || '特典プレゼント',
    capsuleColor: '#d9c5a2',
    accentColor: '#8d6c46',
    message: '受付でその時の特典をお受け取りください。'
  };
}

function getCurrentCardReward(cardNum) {
  const targetCardNum = Math.max(1, Number(cardNum || stampCardNum) || 1);
  return normalizeRewardList(EARNED_REWARDS).find(function (reward) {
    return Math.max(1, Number(reward && reward.cardNum) || 1) === targetCardNum;
  }) || null;
}

function hasCurrentCardReward(cardNum) {
  return !!getCurrentCardReward(cardNum);
}

function drawRewardGachaWeightedPrize() {
  const totalWeight = REWARD_GACHA_PRIZE_POOL.reduce(function (sum, prize) {
    return sum + Math.max(1, Number(prize.weight) || 0);
  }, 0);
  let roll = Math.random() * totalWeight;
  for (let i = 0; i < REWARD_GACHA_PRIZE_POOL.length; i++) {
    roll -= Math.max(1, Number(REWARD_GACHA_PRIZE_POOL[i].weight) || 0);
    if (roll < 0) {
      return Object.assign({}, REWARD_GACHA_PRIZE_POOL[i]);
    }
  }
  return Object.assign({}, REWARD_GACHA_PRIZE_POOL[REWARD_GACHA_PRIZE_POOL.length - 1]);
}

function buildLocalGachaReward(prizeMeta) {
  const achievedAt = normalizeRewardDateTime(localStorage.getItem('mayumi_stamp_10_date')) || new Date().toISOString();
  const expiryDate = new Date(achievedAt);
  expiryDate.setMonth(expiryDate.getMonth() + 1);
  return normalizeSingleReward({
    id: 'reward-' + Date.now() + '-' + Math.floor(Math.random() * 1000),
    cardNum: stampCardNum,
    rewardName: prizeMeta.rewardName,
    earnedDate: achievedAt,
    expiryDate: expiryDate.toISOString(),
    used: false
  }, 0);
}

function renderRewardGachaCapsules(activePrizeKey) {
  const container = document.getElementById('rewardGachaCapsules');
  const slotWindow = document.getElementById('rewardGachaSlotWindow');
  if (!container || !slotWindow) return;

  const capsuleLayout = [
    { left: '18px', top: '18px', prizeIdx: 0 },
    { left: '86px', top: '26px', prizeIdx: 1 },
    { left: '148px', top: '54px', prizeIdx: 2 },
    { left: '34px', top: '90px', prizeIdx: 3 },
    { left: '112px', top: '106px', prizeIdx: 1 },
    { left: '170px', top: '122px', prizeIdx: 2 }
  ];

  container.innerHTML = capsuleLayout.map(function (capsule) {
    const prize = REWARD_GACHA_PRIZE_POOL[capsule.prizeIdx % REWARD_GACHA_PRIZE_POOL.length];
    const isActive = activePrizeKey && prize.key === activePrizeKey;
    const boxShadow = isActive
      ? '0 0 0 3px rgba(255,255,255,0.78), 0 12px 20px rgba(176, 121, 27, 0.28)'
      : 'inset -5px -10px 0 rgba(0, 0, 0, 0.08), inset 6px 8px 0 rgba(255, 255, 255, 0.28), 0 8px 14px rgba(0, 0, 0, 0.1)';
    return `
      <div class="capsule" style="left:${capsule.left}; top:${capsule.top}; background:${prize.capsuleColor}; box-shadow:${boxShadow};">
        <span class="capsule-content" aria-hidden="true">🎁</span>
      </div>
    `;
  }).join('');

  const slotPrize = activePrizeKey
    ? (REWARD_GACHA_PRIZE_POOL.find(function (prize) { return prize.key === activePrizeKey; }) || REWARD_GACHA_PRIZE_POOL[0])
    : null;
  slotWindow.textContent = slotPrize ? slotPrize.rankLabel : '???';
  slotWindow.style.color = slotPrize ? slotPrize.accentColor : '#8f7a61';
  slotWindow.style.background = slotPrize
    ? `linear-gradient(180deg, ${slotPrize.capsuleColor}, #fff7ea)`
    : 'linear-gradient(180deg, #fff8f0, #f7efe5)';
}

function showRewardGachaResult(prizeMeta, options) {
  const result = document.getElementById('rewardGachaResult');
  const badge = document.getElementById('rewardGachaResultBadge');
  const rank = document.getElementById('rewardGachaResultRank');
  const name = document.getElementById('rewardGachaResultName');
  const message = document.getElementById('rewardGachaResultMessage');
  const gift = document.getElementById('rewardGachaResultGift');
  const drawBtn = document.getElementById('rewardGachaDrawBtn');
  const nextBtn = document.getElementById('rewardGachaNextBtn');
  const handleBtn = document.getElementById('rewardGachaHandleBtn');
  const status = document.getElementById('rewardGachaStatus');
  const machine = document.getElementById('rewardGachaMachine');
  const prize = getRewardGachaPrizeMeta(prizeMeta && prizeMeta.rewardName ? prizeMeta.rewardName : prizeMeta);
  const alreadyDrawn = !!(options && options.alreadyDrawn);

  if (gift) {
    gift.textContent = '🎁';
  }

  if (result) result.style.display = 'block';
  if (badge) {
    badge.textContent = alreadyDrawn ? `${prize.rankLabel} 獲得済み` : `${prize.rankLabel} が当たりました`;
    badge.style.color = prize.accentColor;
  }
  // ランク表示（○賞）
  if (rank) {
    rank.textContent = prize.rankLabel;
    rank.style.color = prize.accentColor;
  }
  // 特典内容表示
  if (name) {
    // rewardNameから賞ラベル部分を除いて特典内容のみ表示
    const rewardContent = prize.rewardName.replace(prize.rankLabel, '').trim();
    name.textContent = rewardContent || prize.rewardName;
  }
  if (message) {
    message.innerHTML = `${escapeHtml(prize.message)}<br>受け取りの際は受付へ直接お問い合わせください。`;
  }
  if (drawBtn) drawBtn.style.display = 'none';
  if (nextBtn) nextBtn.style.display = 'block';
  if (handleBtn) handleBtn.disabled = true;
  if (status) {
    status.textContent = alreadyDrawn
      ? 'このカードのガチャ結果は保存済みです。次のスタンプカードへ進めます。'
      : 'ガチャ結果を保存しました。次のスタンプカードを始められます。';
  }
  if (machine) {
    machine.classList.remove('rolling');
    machine.classList.add('result-ready');
  }
  renderRewardGachaCapsules(prize.key);
  lastRewardGachaResult = prize;
}

function resetRewardGachaModal() {
  const machine = document.getElementById('rewardGachaMachine');
  const result = document.getElementById('rewardGachaResult');
  const drawBtn = document.getElementById('rewardGachaDrawBtn');
  const nextBtn = document.getElementById('rewardGachaNextBtn');
  const closeBtn = document.getElementById('rewardGachaCloseBtn');
  const handleBtn = document.getElementById('rewardGachaHandleBtn');
  const status = document.getElementById('rewardGachaStatus');
  const currentReward = getCurrentCardReward();

  if (machine) {
    machine.classList.remove('rolling', 'result-ready');
  }
  if (result) {
    result.style.display = 'none';
  }
  if (drawBtn) {
    drawBtn.style.display = currentReward ? 'none' : 'block';
    drawBtn.disabled = false;
  }
  if (nextBtn) {
    nextBtn.style.display = currentReward ? 'block' : 'none';
  }
  if (closeBtn) {
    closeBtn.disabled = false;
  }
  if (handleBtn) {
    handleBtn.disabled = !!currentReward;
  }
  if (status) {
    status.textContent = currentReward
      ? 'このカードのガチャ結果は保存済みです。次のスタンプカードへ進めます。'
      : 'ハンドルを回して、特典カプセルを受け取ってください。';
  }
  renderRewardGachaCapsules(currentReward ? getRewardGachaPrizeMeta(currentReward.rewardName).key : '');
  if (currentReward) {
    showRewardGachaResult(currentReward, { alreadyDrawn: true });
  } else {
    lastRewardGachaResult = null;
  }
}

function getLocalRewardStatus() {
  let lastStampDate = '';
  let lastStampAt = '';
  let stampAchievedDate = '';
  try {
    lastStampDate = normalizeStampDateKey(localStorage.getItem('mayumi_last_stamp_date') || '');
    lastStampAt = normalizeRewardDateTime(localStorage.getItem(LAST_STAMP_AT_STORAGE_KEY) || '');
    stampAchievedDate = normalizeRewardDateTime(localStorage.getItem('mayumi_stamp_10_date') || '');
  } catch (e) { }
  const normalizedStampHistory = normalizeStampHistoryList(STAMP_HISTORY);
  const normalizedLastStampAt = deriveLastStampAt({
    lastStampAt: lastStampAt,
    lastStampDate: lastStampDate,
    stampHistory: normalizedStampHistory
  }, normalizedStampHistory);
  return {
    stampCount: Math.max(0, Math.min(10, Number(stampCount) || 0)),
    stampCardNum: Math.max(1, Number(stampCardNum) || 1),
    rewards: normalizeRewardList(EARNED_REWARDS),
    stampHistory: normalizedStampHistory,
    lastStampDate: normalizeStampDateKey(lastStampDate || normalizedLastStampAt),
    lastStampAt: normalizedLastStampAt,
    stampAchievedDate: stampAchievedDate
  };
}

function hasMeaningfulRewardStatus(status) {
  if (!status) return false;
  return Number(status.stampCount || 0) > 0 ||
    Number(status.stampCardNum || 1) > 1 ||
    (Array.isArray(status.rewards) && status.rewards.length > 0) ||
    (Array.isArray(status.stampHistory) && status.stampHistory.length > 0) ||
    !!status.lastStampDate ||
    !!status.stampAchievedDate;
}

function getComparableRewardStatus(status) {
  const normalized = status || getLocalRewardStatus();
  const stampHistory = normalizeStampHistoryList(normalized.stampHistory);
  const lastStampAt = deriveLastStampAt(normalized, stampHistory);
  return {
    stampCount: Math.max(0, Math.min(10, Number(normalized.stampCount) || 0)),
    stampCardNum: Math.max(1, Number(normalized.stampCardNum) || 1),
    rewards: normalizeRewardList(normalized.rewards),
    stampHistory: stampHistory,
    lastStampDate: normalizeStampDateKey(normalized.lastStampDate || lastStampAt),
    lastStampAt: lastStampAt,
    stampAchievedDate: normalizeRewardDateTime(normalized.stampAchievedDate)
  };
}

function rewardStatusEquals(a, b) {
  return JSON.stringify(getComparableRewardStatus(a)) === JSON.stringify(getComparableRewardStatus(b));
}

function applyRewardStatusLocally(status) {
  const next = getComparableRewardStatus(status);
  stampCount = next.stampCount;
  stampCardNum = next.stampCardNum;
  EARNED_REWARDS = next.rewards;
  STAMP_HISTORY = next.stampHistory;
  try {
    localStorage.setItem('mayumi_stamp', String(stampCount));
    localStorage.setItem('mayumi_stamp_card', String(stampCardNum));
    localStorage.setItem('mayumi_earned_rewards', JSON.stringify(EARNED_REWARDS));
    localStorage.setItem(STAMP_HISTORY_STORAGE_KEY, JSON.stringify(STAMP_HISTORY));
    if (next.lastStampDate) localStorage.setItem('mayumi_last_stamp_date', next.lastStampDate);
    else localStorage.removeItem('mayumi_last_stamp_date');
    if (next.lastStampAt) localStorage.setItem(LAST_STAMP_AT_STORAGE_KEY, next.lastStampAt);
    else localStorage.removeItem(LAST_STAMP_AT_STORAGE_KEY);
    if (next.stampAchievedDate) localStorage.setItem('mayumi_stamp_10_date', next.stampAchievedDate);
    else localStorage.removeItem('mayumi_stamp_10_date');
  } catch (e) { }
  updateStampUI();
  renderEarnedRewards();
}

async function syncRewardStatus(force) {
  if (!_profile || !_profile.memberId) return null;
  const localStatus = getLocalRewardStatus();
  if (!force && lastSyncedRewardStatus && rewardStatusEquals(localStatus, lastSyncedRewardStatus)) {
    return lastSyncedRewardStatus;
  }
  const activityProfile = getUserActivityProfilePayload();
  const res = await postToGAS({
    type: 'syncUserRewardStatus',
    memberId: _profile.memberId,
    name: activityProfile.name,
    phone: activityProfile.phone,
    birthday: activityProfile.birthday,
    address: activityProfile.address,
    stampCount: localStatus.stampCount,
    stampCardNum: localStatus.stampCardNum,
    rewards: localStatus.rewards,
    stampHistory: localStatus.stampHistory,
    lastStampDate: localStatus.lastStampDate,
    lastStampAt: localStatus.lastStampAt,
    stampAchievedDate: localStatus.stampAchievedDate
  });
  if (res && res.status === 'ok' && res.rewardStatus) {
    lastSyncedRewardStatus = getComparableRewardStatus(res.rewardStatus);
    applyRewardStatusLocally(lastSyncedRewardStatus);
  }
  return res;
}

function renderStampGrid(id, count) {
  const grid = document.getElementById(id);
  if (!grid) return;
  grid.innerHTML = '';
  for (let i = 0; i < 10; i++) {
    const cell = document.createElement('div');
    cell.className = 'stamp-cell ' + (i < count ? 'filled' : 'empty');
    if (i < count) cell.textContent = '🌿';
    grid.appendChild(cell);
  }
}

function updateStampUI() {
  renderStampGrid('homeStampGrid', stampCount);

  const rem = Math.max(0, 10 - stampCount);
  const elNum = document.getElementById('homeStampNum');
  if (elNum) elNum.textContent = stampCount;
  const elCount = document.getElementById('homeStampCount');
  if (elCount) elCount.textContent = stampCount;

  const currentReward = getCurrentCardReward();
  let msg = rem > 0
    ? `あと${rem}回でプレゼント🎁`
    : (currentReward
      ? '達成済み🎉 ガチャ結果を受け取りました。新しいカードを始められます。'
      : '達成済み🎉 特典ガチャを回せます！');

  if (rem === 0) {
    const achievedAt = localStorage.getItem('mayumi_stamp_10_date');
    if (achievedAt) {
      const achievedDate = new Date(achievedAt);
      const expiryDate = new Date(achievedDate);
      expiryDate.setMonth(expiryDate.getMonth() + 1);
      const diffMs = expiryDate - new Date();
      const diffDays = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
      msg += diffDays > 0
        ? `<br><span style="font-size:12px; color:var(--danger); font-weight:bold;">受取期限: あと${diffDays}日</span>`
        : `<br><span style="font-size:12px; color:var(--danger); font-weight:bold;">受取期限: 本日まで</span>`;
    }
  }

  const elMsg = document.getElementById('homeStampMsg');
  if (elMsg) elMsg.innerHTML = msg;

  const cardLabel = document.getElementById('homeStampCardLabel');
  if (cardLabel) cardLabel.textContent = `${stampCardNum}枚目`;

  const newCardBtn = document.getElementById('stampNewCardBtn');
  if (newCardBtn) {
    newCardBtn.style.display = stampCount >= 10 ? 'block' : 'none';
    if (stampCount >= 10) {
      newCardBtn.textContent = currentReward ? '🌸 新しいスタンプカードを取得' : '🎯 特典ガチャを回す';
      newCardBtn.onclick = currentReward ? startNewCard : openRewardGachaModal;
    }
  }
}

function addStamp() {
  const today = getTodayStampKey();
  const acquiredAt = new Date().toISOString();
  let lastStampDate = '';
  try {
    lastStampDate = normalizeStampDateKey(localStorage.getItem('mayumi_last_stamp_date'));
  } catch (e) { }

  if (lastStampDate === today) {
    showToast('本日はすでにスタンプを取得済みです。来院スタンプは1日1回までです！');
    return;
  }
  if (stampCount >= 10) {
    showToast(hasCurrentCardReward() ? '達成済みです。新しいスタンプカードを始めてください。' : '達成済みです。特典ガチャを回してください。');
    return;
  }

  stampCount += 1;
  recordStampHistoryEntry(stampCardNum, stampCount, acquiredAt);
  try {
    localStorage.setItem('mayumi_stamp', String(stampCount));
    localStorage.setItem('mayumi_last_stamp_date', today);
    localStorage.setItem(LAST_STAMP_AT_STORAGE_KEY, acquiredAt);
    if (stampCount === 10) {
      localStorage.setItem('mayumi_stamp_10_date', acquiredAt);
    }
  } catch (e) { }

  updateStampUI();
  renderEarnedRewards();
  const reachedMilestone = stampCount === 10;
  if (reachedMilestone) {
    triggerConfetti();
  }
  updateStampModalPresentation(reachedMilestone);
  document.getElementById('stampPopMsg').textContent = reachedMilestone
    ? 'ご来院ありがとうございます！10個達成です。特典ガチャを回せます。'
    : 'ご来院ありがとうございます！スタンプを1つ追加しました！';
  document.getElementById('stampPopCount').textContent = `現在 ${stampCount} / 10`;

  openModal('stampModal');
  syncRewardStatus();
}

function handleStampMilestoneAction() {
  if (hasCurrentCardReward()) {
    startNewCard();
    return;
  }
  openRewardGachaModal();
}

function openRewardGachaModal() {
  if (stampCount < 10) {
    showToast('スタンプが10個たまると特典ガチャを回せます。');
    return;
  }
  closeModal('stampModal');
  resetRewardGachaModal();
  openModal('rewardGachaModal');
}

async function drawRewardGacha() {
  if (isRewardGachaDrawing) return;
  if (stampCount < 10) {
    showToast('スタンプが10個たまると特典ガチャを回せます。');
    return;
  }

  const currentReward = getCurrentCardReward();
  if (currentReward) {
    showRewardGachaResult(currentReward, { alreadyDrawn: true });
    return;
  }

  const drawBtn = document.getElementById('rewardGachaDrawBtn');
  const closeBtn = document.getElementById('rewardGachaCloseBtn');
  const handleBtn = document.getElementById('rewardGachaHandleBtn');
  const machine = document.getElementById('rewardGachaMachine');
  const status = document.getElementById('rewardGachaStatus');
  const dropCapsule = document.getElementById('rewardGachaDroppedCapsule');
  const startedAt = Date.now();

  isRewardGachaDrawing = true;
  if (drawBtn) drawBtn.disabled = true;
  if (closeBtn) closeBtn.disabled = true;
  if (handleBtn) handleBtn.disabled = true;
  if (machine) machine.classList.add('rolling');
  if (status) status.textContent = 'ガチャを回しています…';
  renderRewardGachaCapsules();

  try {
    let res = null;
    if (_profile && _profile.memberId) {
      res = await postToGAS({
        type: 'drawRewardGacha',
        memberId: _profile.memberId
      });
    }

    if (!res || res.status !== 'ok') {
      const fallbackPrize = drawRewardGachaWeightedPrize();
      const localReward = buildLocalGachaReward(fallbackPrize);
      const nextStatus = getComparableRewardStatus({
        stampCount: stampCount,
        stampCardNum: stampCardNum,
        rewards: [localReward].concat(normalizeRewardList(EARNED_REWARDS)),
        lastStampDate: normalizeStampDateKey(localStorage.getItem('mayumi_last_stamp_date') || ''),
        lastStampAt: normalizeRewardDateTime(localStorage.getItem(LAST_STAMP_AT_STORAGE_KEY) || ''),
        stampAchievedDate: normalizeRewardDateTime(localStorage.getItem('mayumi_stamp_10_date') || '')
      });
      applyRewardStatusLocally(nextStatus);
      lastSyncedRewardStatus = nextStatus;
      res = {
        status: 'ok',
        rewardStatus: nextStatus,
        drawnReward: Object.assign({}, fallbackPrize)
      };
    } else if (res.rewardStatus) {
      lastSyncedRewardStatus = getComparableRewardStatus(res.rewardStatus);
      applyRewardStatusLocally(lastSyncedRewardStatus);
    }

    // 演出時間の確保 (最小1.2秒)
    const elapsed = Date.now() - startedAt;
    if (elapsed < 1200) {
      await new Promise(function (resolve) { setTimeout(resolve, 1200 - elapsed); });
    }

    // 回転停止
    if (machine) machine.classList.remove('rolling');
    if (status) status.textContent = '特典が出てきます…';

    const drawnReward = (res && res.drawnReward) || getCurrentCardReward();
    const prizeColor = drawnReward ? getRewardGachaPrizeMeta(drawnReward.rewardName).capsuleColor : '#d9c5a2';

    // カプセル落下演出
    if (dropCapsule) {
      dropCapsule.style.background = prizeColor;
      dropCapsule.classList.add('falling');
      await new Promise(r => setTimeout(r, 600));
      dropCapsule.classList.add('bouncing');
      await new Promise(r => setTimeout(r, 400));
    }

    showRewardGachaResult(drawnReward, { alreadyDrawn: !!(res && res.alreadyDrawn) });
    triggerConfetti();
    if (closeBtn) closeBtn.disabled = false;
  } catch (err) {
    console.log('drawRewardGacha error:', err);
    showToast('特典ガチャの取得に失敗しました。');
    if (machine) machine.classList.remove('rolling');
    if (status) status.textContent = '通信に失敗しました。もう一度お試しください。';
    if (drawBtn) drawBtn.disabled = false;
    if (closeBtn) closeBtn.disabled = false;
    if (handleBtn) handleBtn.disabled = false;
  } finally {
    isRewardGachaDrawing = false;
  }
}

async function startNewCard() {
  if (stampCount < 10) {
    showToast('スタンプが10個たまると新しいカードを開始できます。');
    return;
  }
  if (!hasCurrentCardReward()) {
    showAppAlert('新しいスタンプカードを始める前に、特典ガチャを回してください。', {
      title: '特典ガチャのご案内',
      confirmLabel: '閉じる'
    });
    return;
  }

  stampCount = 0;
  stampCardNum += 1;
  try {
    localStorage.setItem('mayumi_stamp', String(stampCount));
    localStorage.setItem('mayumi_stamp_card', String(stampCardNum));
    localStorage.removeItem('mayumi_stamp_10_date');
  } catch (e) { }
  updateStampUI();
  triggerConfetti();
  closeModal('stampModal');
  closeModal('rewardGachaModal');
  showToast(`🌸 ${stampCardNum}枚目のスタンプカードを開始しました！`);
  await syncRewardStatus(true);
}

function renderEarnedRewards() {
  const container = document.getElementById('earnedRewardsList');
  if (!container) return;

  const stampHistoryItems = normalizeStampHistoryList(STAMP_HISTORY).map(function (entry) {
    return {
      type: 'stamp',
      occurredAt: entry.acquiredDate,
      entry: entry
    };
  });
  const rewardHistoryItems = normalizeRewardList(EARNED_REWARDS).map(function (reward) {
    return {
      type: 'reward',
      occurredAt: reward.earnedDate,
      entry: reward
    };
  });
  const historyItems = stampHistoryItems.concat(rewardHistoryItems).sort(function (a, b) {
    return new Date(b.occurredAt).getTime() - new Date(a.occurredAt).getTime();
  });

  if (!historyItems.length) {
    container.innerHTML = '<div style="text-align:center;font-size:13px;color:var(--text-light);padding:26px 0">スタンプ・特典の履歴はありません</div>';
    return;
  }

  const now = new Date();
  let html = '';
  historyItems.forEach(function (item) {
    if (item.type === 'stamp') {
      const stampEntry = item.entry;
      const acquiredStr = formatCustomerDateYmdHm(stampEntry.acquiredDate) || formatCustomerDateYmd(stampEntry.acquiredDate);
      const isMilestone = Number(stampEntry.stampNumber) === 10;
      html += `
        <div style="padding:16px; border-radius:16px; margin-bottom:12px; display:flex; flex-direction:column; box-shadow:0 4px 12px rgba(0,0,0,0.05); background:#f8fcf6; border:1px solid rgba(126, 154, 109, 0.28);">
          <div style="display:flex; justify-content:space-between; align-items:center; gap:10px; margin-bottom:10px;">
            <span style="font-weight:600; font-size:12px; color:#567244; background:rgba(181, 201, 168, 0.24); padding:4px 10px; border-radius:100px;">来院スタンプ ${stampEntry.cardNum}枚目カード</span>
            <span style="font-size:11px; padding:3px 8px; border-radius:12px; background:${isMilestone ? 'var(--primary)' : '#e6f2df'}; color:${isMilestone ? '#fff' : '#567244'};">${isMilestone ? '🎉 10個達成' : `🌿 ${stampEntry.stampNumber}個目`}</span>
          </div>
          <div style="font-size:18px; font-weight:bold; color:var(--text-dark); margin-bottom:8px;">${stampEntry.cardNum}枚目カードの${stampEntry.stampNumber}個目のスタンプを取得しました</div>
          <div style="font-size:13px; color:var(--sage-dark); font-weight:500; line-height:1.5; margin-bottom:12px; padding:10px; background:#fff; border-radius:8px;">
            ${isMilestone
          ? 'スタンプが10個たまりました。ホームから特典ガチャを回し、マイページの履歴で結果を確認できます。'
          : '来院スタンプを取得しました。スタンプが10個たまると特典ガチャに進めます。'}
          </div>
          <div style="font-size:12px; color:var(--text-light); line-height:1.5;">取得日時: ${acquiredStr}</div>
        </div>
      `;
      return;
    }

    const reward = item.entry;
    const expiry = new Date(reward.expiryDate);
    const isExpired = now > expiry;
    let statusHtml = '';
    let btnHtml = '';
    let cardStyle = 'background:#fff; border:1px solid var(--primary); opacity:1;';

    if (reward.used) {
      statusHtml = '<span style="font-size:11px; padding:3px 8px; border-radius:12px; background:#e0e0e0; color:#555;">使用済み</span>';
      cardStyle = 'background:#f9f9f9; border:1px solid #ddd; opacity:0.6;';
    } else if (isExpired) {
      statusHtml = '<span style="font-size:11px; padding:3px 8px; border-radius:12px; background:#ffebee; color:#d32f2f;">期限切れ</span>';
      cardStyle = 'background:#f9f9f9; border:1px solid #ddd; opacity:0.6;';
    } else {
      statusHtml = '<span style="font-size:11px; padding:3px 8px; border-radius:12px; background:var(--primary); color:#fff;">🎁 未使用</span>';
      btnHtml = `<button class="btn primary" style="padding:6px 12px; font-size:12px; margin-top:10px;" onclick='useReward(${JSON.stringify(reward.id)})'>使用する</button>`;
    }

    const earnedStr = formatCustomerDateYmd(reward.earnedDate);
    const expiryStr = formatCustomerDateYmd(reward.expiryDate);
    const rewardTitle = escapeHtml(String(reward.rewardName || '特典プレゼント'));
    let countdownHtml = '';
    if (!reward.used && !isExpired) {
      const diffMs = expiry - now;
      const diffDays = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
      countdownHtml = diffDays > 0
        ? `<div style="font-size:12px; font-weight:bold; color:var(--danger); margin-bottom:8px;">期限まで あと${diffDays}日</div>`
        : `<div style="font-size:12px; font-weight:bold; color:var(--danger); margin-bottom:8px;">期限まで 本日まで</div>`;
    }

    html += `
      <div style="padding:16px; border-radius:16px; margin-bottom:12px; display:flex; flex-direction:column; box-shadow: 0 4px 12px rgba(0,0,0,0.05); ${cardStyle}">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
          <span style="font-weight:600; font-size:12px; color:var(--text-main); background:rgba(181, 201, 168, 0.2); padding:4px 10px; border-radius:100px;">スタンプ ${reward.cardNum}枚目 特典</span>
          ${statusHtml}
        </div>
        <div style="font-size:18px; font-weight:bold; color:var(--text-dark); margin-bottom:10px;">${rewardTitle}</div>
        <div style="font-size:13px; color:var(--sage-dark); font-weight:500; line-height:1.5; margin-bottom:12px; padding:10px; background:var(--cream); border-radius:8px;">
          🎁 <strong>受け取りの際は受付へ直接お問い合わせください。</strong><br>
          ※ 特典は達成当日から使用できます。<br>
          ※ 受取期間は特典獲得から1ヶ月で、一度使用すると再度は使用できません。
        </div>
        ${countdownHtml}
        <div style="font-size:12px; color:var(--text-light); line-height:1.5; margin-bottom:12px;">獲得日: ${earnedStr} <br> 有効期限: ${expiryStr}まで</div>
        ${btnHtml}
      </div>
    `;
  });

  container.innerHTML = html;

  const badge = document.getElementById('badge-mypage');
  if (badge) {
    const hasExpiring = EARNED_REWARDS.some(function (reward) {
      if (reward.used) return false;
      const expiry = new Date(reward.expiryDate);
      if (now > expiry) return false;
      const diffMs = expiry - now;
      const diffDays = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
      return diffDays <= 7;
    });
    badge.style.display = hasExpiring ? 'block' : 'none';
  }
}

function useReward(id) {
  const reward = EARNED_REWARDS.find(function (entry) {
    return String(entry.id) === String(id);
  });
  if (!reward) return;

  if (reward.used) {
    showAppAlert('この特典はすでに使用済みのため、再度は使用できません。', {
      title: '特典のご利用について',
      confirmLabel: '閉じる'
    });
    return;
  }

  const expiry = new Date(reward.expiryDate);
  if (new Date() > expiry) {
    showAppAlert('この特典は受取期限を過ぎているため使用できません。', {
      title: '特典のご利用について',
      confirmLabel: '閉じる'
    });
    return;
  }

  showAppConfirm('本当に使用しますか？\nこの操作は取り消しできません。', {
    title: '特典を使用しますか？',
    confirmLabel: '使用する',
    cancelLabel: '戻る',
    confirmVariant: 'primary'
  }).then(function (confirmed) {
    if (!confirmed) return;

    reward.used = true;
    reward.usedAt = new Date().toISOString();
    try {
      localStorage.setItem('mayumi_earned_rewards', JSON.stringify(EARNED_REWARDS));
    } catch (e) { }
    renderEarnedRewards();
    syncRewardStatus(true);

    const now = new Date();
    const timeStr = now.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' });
    document.getElementById('rewardUseTimeMsg').textContent = timeStr;
    openModal('rewardUseCompleteModal');
  });
}

// ===== 商品データ =====

const PROD_IMAGES = {
  tea50: 'img/tea50.jpg',
  tea30: 'img/tea30.jpg',
  bath10: 'img/bath10.jpg',
  bath2: 'img/bath2.jpg',
  soap: 'img/soap.jpg',
  pad: 'img/pad.jpg',
};

const PRODUCTS = [
  { category: 'よもぎ茶', name: 'よもぎ茶（50パック）', price: 2100, icon: '🍵', bg: 'c1', imgKey: 'tea50', description: '無農薬栽培のよもぎをたっぷり50パック詰めました。毎日の健康維持に。' },
  { category: 'よもぎ茶', name: 'よもぎ茶（30パック）', price: 1575, icon: '🍵', bg: 'c1', imgKey: 'tea30', description: '持ち運びにも便利な30パック入り。良質なよもぎの香りが楽しめます。' },
  { category: 'よもぎ入浴剤', name: 'よもぎ入浴剤（10パック）', price: 1540, icon: '🛁', bg: 'c4', imgKey: 'bath10', description: '体の芯から温まるよもぎ入浴剤。10回分のお得なセットです。' },
  { category: 'よもぎ入浴剤', name: 'よもぎ入浴剤（2パック）', price: 350, icon: '🛁', bg: 'c4', imgKey: 'bath2', description: 'ちょっとしたお試しやプレゼントに最適な2パック入り入浴剤。' },
  { category: '石鹸', name: '石鹸（あこがれのきよら）', price: 540, icon: '🧼', bg: 'c2', imgKey: 'soap', description: '天然由来成分にこだわった、お肌に優しい洗顔・全身用石鹸です。' },
  { category: '冷え取りパット', name: '冷え取りパット（3枚セット）', price: 2400, icon: '🌸🌸🌸', bg: 'c3', imgKey: 'pad', description: '洗い替えに便利な3枚セット。オーガニックコットン使用で快適です。' },
  { category: '冷え取りパット', name: '冷え取りパット（2枚）', price: 1760, icon: '🌸🌸', bg: 'c3', imgKey: 'pad', description: '冷え対策の定番。肌触りの良いパット2枚セットです。' },
  { category: '冷え取りパット', name: '冷え取りパット（1枚）', price: 880, icon: '🌸', bg: 'c3', imgKey: 'pad', description: 'まずは試してみたい方に。高品質な国産よもぎ成分を配合。' },
];

const NOTICE_FEED_START_DATE = '2026-04-01';

const FALLBACK_BLOG = [
  { date: '2025.06.15', category: 'お知らせ', type: 'お知らせ', icon: '🌷', title: '夏の産後ヨガ体験会開催のご案内', body: '今年の夏も産後ヨガ体験会を開催いたします。初めての方も大歓迎です。お気軽にご参加ください。' },
  { date: '2025.06.10', category: 'ブログ', type: 'ブログ', icon: '🍵', title: '授乳期にうれしい！おすすめハーブティー', body: '授乳中のママにぴったりのよもぎ茶をご紹介します。ノンカフェインで体を温める効果があります。' },
  { date: '2025.06.05', category: '休診情報', type: 'お知らせ', icon: '📅', title: '6月22日（日）は臨時休診となります', body: '誠に恐れ入りますが、6月22日（日）は臨時休診とさせていただきます。ご不便をおかけして申し訳ございません。' },
  { date: '2025.05.28', category: 'ブログ', type: 'ブログ', icon: '🤱', title: '産後の骨盤ケア、いつから始める？', body: '産後の骨盤は出産の影響でゆるんだ状態。適切なタイミングと方法でケアを始めることが大切です。' },
];

/* SUPPORT_FAQ_FALLBACK_START */
const SUPPORT_FAQ_FALLBACK = [
  { category: 'アプリ全般', question: 'このアプリでできることを知りたい', keywords: 'アプリ,使い方,できること,何ができる,機能,全体', answer: 'このアプリでは、ホーム、ショップ、カレンダー、NEWS、マイページ、お知らせ一覧、スタンプQR読み取り、商品注文、注文履歴確認、特典確認、通知設定、起動時パスコード、データ引き継ぎ・復元、引き継ぎコード発行、使い方サポートが利用できます。予約確定や個別相談は公式LINEをご利用ください。', priority: 160 },
  { category: 'アプリ全般', question: 'アプリの画面構成を教えてください', keywords: '画面,構成,タブ,ナビ,メニュー,下部,上部', answer: '画面下部にはホーム、ショップ、カレンダー、NEWS、マイページがあります。画面上部からは、最新情報の更新、お知らせ一覧、カート、マイページショートカットを利用できます。', priority: 158 },
  { category: '会員登録', question: '初回起動時はどちらを選べばいいですか？', keywords: '初回,最初,はじめて,以前登録した方,はじめて登録する方,どちら', answer: 'アプリを開いた最初の選択画面で、以前に登録したことがある方は「以前登録した方はこちら ↺」、今回が初めての方は「はじめて登録する方はこちら」を選んでください。再インストール後、機種変更後、ブラウザ版からホーム画面追加した後も、以前登録したことがある方は復元側から進んでください。', priority: 156 },
  { category: 'プロフィール', question: 'プロフィールの登録方法を知りたい', keywords: 'プロフィール,登録,会員,名前,電話,住所,生年月日,初回登録', answer: '初回起動時は、まず「以前登録した方」と「はじめて登録する方」の選択画面が出ます。はじめて登録する方は、そのままプロフィール登録へ進み、お名前・電話番号・生年月日・住所を入力してください。保存すると会員IDが発行され、アプリを使い始められます。', priority: 154 },
  { category: 'プロフィール', question: 'プロフィールの変更方法を知りたい', keywords: 'プロフィール,変更,編集,名前,電話,住所,生年月日', answer: 'マイページを開き、「✏️ プロフィールを編集」を押してください。お名前、電話番号、生年月日、住所、アイコン画像、バナー画像を変更して保存できます。', priority: 152 },
  { category: 'プロフィール', question: '会員IDはどこで確認できますか？', keywords: '会員ID,会員番号,memberid,どこ,確認', answer: '会員IDはマイページ上部に表示されます。プロフィール登録または復元が完了すると発行されます。', priority: 150 },
  { category: 'ログイン', question: '起動時のパスコード設定について知りたい', keywords: 'パスコード,ログイン,起動時,4桁,6桁,設定', answer: '新しく登録する方も、すでに登録済みの方も、まずは4桁または6桁のパスコードを設定して使います。既存会員の方はアプリ起動時に設定画面が表示されます。', priority: 148 },
  { category: 'ログイン', question: 'ログイン時のパスコードを毎回入力したくないです', keywords: 'ログイン時のパスコード,毎回,入力したくない,オフ,省略', answer: 'パスコードを一度設定したあと、マイページの「ログイン時のパスコード」からオン・オフを切り替えられます。オフにすると、次回からアプリ起動時のパスコード入力を省略できます。', priority: 146 },
  { category: 'ログイン', question: 'パスコードの変更方法を知りたい', keywords: 'パスコード,変更,変える,ログイン,再設定', answer: 'マイページを開き、「🔐 パスコードを変更」を押してください。現在のパスコードを確認したあと、新しい4桁または6桁のパスコードへ変更できます。', priority: 144 },
  { category: 'ログイン', question: 'パスコードを忘れたときの再設定方法を知りたい', keywords: 'パスコード,忘れた,再設定,ログインできない', answer: 'ログイン画面、または「ログイン・引き継ぎ」画面にある「パスコードを忘れた場合の再設定はこちら」から再設定できます。登録したお名前・電話番号・生年月日を入力し、新しい4桁または6桁のパスコードを設定してください。', priority: 142 },
  { category: '引き継ぎ', question: 'データの引き継ぎ・復元方法を知りたい', keywords: '引き継ぎ,復元,機種変更,データ移行,ログイン,再インストール', answer: 'ログイン画面、または初回画面で「以前登録した方はこちら ↺」を選ぶと復元画面へ進めます。引き継ぎコードがある場合は、引き継ぎコードと新しいパスコードを入力してください。引き継ぎコードがない場合は、お名前に加えて電話番号・生年月日・現在のパスコードのうち1つ以上を入力すると復元できます。', priority: 140 },
  { category: '引き継ぎ', question: '引き継ぎコードの発行方法を知りたい', keywords: '引き継ぎコード,発行,機種変更,再インストール,コード', answer: 'マイページの「↺ 引き継ぎコードの発行」から発行できます。機種変更や再インストールの前に発行しておくと、新しい端末の「データの引き継ぎ・復元」で使えます。', priority: 138 },
  { category: '引き継ぎ', question: '引き継ぎコードの有効期限を知りたい', keywords: '引き継ぎコード,有効期限,いつまで,何日,使えない', answer: '引き継ぎコードは1回限りで、発行から1週間有効です。期限切れ、または一度使用したコードは使えません。必要な場合はマイページから新しいコードを発行してください。', priority: 136 },
  { category: '会員登録', question: '会員登録が重複しないようにする方法を知りたい', keywords: '重複,二重,会員登録,同じ名前,会員ID,ブラウザ,ホーム画面', answer: '以前登録したことがある方は、新規登録へ進まず、必ず「以前登録した方はこちら ↺」から復元してください。ブラウザで先に登録したあとホーム画面に追加した方、再インストールした方、機種変更した方も同じです。新規登録をすると、同じお名前でも別の会員IDが作られることがあります。', priority: 134 },
  { category: 'ホーム画面', question: 'ホーム画面への追加方法を知りたい', keywords: 'ホーム画面,追加,インストール,iphone,android,safari,chrome', answer: 'iPhone は Safari でサイトを開き、共有ボタンから「ホーム画面に追加」を選びます。Android は Chrome のメニューから「ホーム画面に追加」または「アプリをインストール」を選びます。すでに登録済みの方は、追加後に新規登録せず「以前登録した方はこちら ↺」から入ってください。', priority: 132 },
  { category: 'ホーム', question: 'ホーム画面の見方を知りたい', keywords: 'ホーム,トップ,home,見方', answer: 'ホーム画面では、スタンプQRの読み取り、現在のスタンプカード確認、おすすめ商品、最新のNEWS、メニュー一覧への移動、公式サイト・SNSへのリンクを確認できます。', priority: 130 },
  { category: '注文', question: '商品の注文方法を知りたい', keywords: '注文,買い方,ショップ,カート,購入', answer: '下部メニューの「ショップ」を開き、商品を選んで詳細を確認し、「カートに追加する」を押してください。画面上部の🛒からカートを開き、内容を確認して「ご注文を確定する」で注文できます。', priority: 128 },
  { category: '注文', question: 'カートの使い方を知りたい', keywords: 'カート,買い物かご,個数,削除,変更', answer: '商品を「カートに追加する」で入れたあと、画面上部の🛒を押すとカートを開けます。カートでは個数変更や削除ができます。商品が入っていない場合は空の画面が表示されます。', priority: 126 },
  { category: '注文', question: '支払い方法を知りたい', keywords: '支払い,決済,現金,支払方法', answer: '商品注文のお支払い方法は現在「現金払い」です。ご来院時に受付でお支払いください。', priority: 124 },
  { category: '注文', question: '商品の受け取り方法を知りたい', keywords: '受取,受け取り,配送,宅配,受領', answer: '商品は院内受け取りです。注文後、ご来院時にスタッフへお声がけください。配送には対応していません。', priority: 122 },
  { category: '注文', question: '注文履歴の見方を知りたい', keywords: '注文履歴,履歴,受取,受付中,キャンセル', answer: 'マイページの「📋 ご注文履歴」から確認できます。受付中の注文や受け取り前の注文が表示され、状況を確認できます。', priority: 120 },
  { category: '注文', question: '注文をキャンセルしたいです', keywords: '注文,キャンセル,取り消し,受付中', answer: 'マイページの「📋 ご注文履歴」を開き、受付中の注文に表示される「キャンセルする」を押してください。キャンセル済みの注文は履歴から表示されなくなります。', priority: 118 },
  { category: '注文', question: '受け取りましたボタンの使い方を知りたい', keywords: '受け取りました,受取完了,受取報告,注文履歴', answer: 'マイページの「📋 ご注文履歴」を開き、受け取り済みにしたい注文の「受け取りました」を押してください。更新後、その注文は履歴に残らなくなります。', priority: 116 },
  { category: 'メニュー', question: 'メニュー一覧の見方を知りたい', keywords: 'メニュー,施術,一覧,カテゴリ,見方', answer: 'ホーム画面の「メニュー一覧を見る🍴」を押すと一覧が開きます。必要に応じてカテゴリで絞り込みでき、各メニューをタップすると詳細や画像を確認できます。', priority: 114 },
  { category: '予約', question: '予約はアプリからできますか？', keywords: '予約,よやく,line,予約方法,相談', answer: '現在、アプリから予約確定はできません。予約や個別相談は公式LINEをご利用ください。ホーム画面またはマイページの「🔗 公式サイト・SNS」から公式LINEを開けます。', priority: 112 },
  { category: 'スタンプ', question: 'スタンプの集め方を知りたい', keywords: 'スタンプ,QR,QRコード,来院,カメラ', answer: 'ホーム画面の「📷 カメラを起動して読み取る」を押し、表示された案内でカメラを許可してから院内QRコードを読み取ってください。読み取りに成功するとスタンプが1つ追加されます。', priority: 110 },
  { category: 'スタンプ', question: 'スタンプは1日何回取得できますか？', keywords: 'スタンプ,1日,一日,何回,回数', answer: '来院スタンプは1日1回までです。同じ日に再度読み取ると、すでに取得済みの案内が表示されます。', priority: 108 },
  { category: 'トラブル', question: 'カメラが起動しないときはどうすればいいですか？', keywords: 'カメラ,起動しない,許可,権限,QR,読めない', answer: 'スタンプ取得時にカメラ許可の確認が出た場合は「許可」を選んでください。すでに拒否している場合は、表示される「設定を開く」から設定画面へ進み、iPhone や Android のカメラ許可をオンにしてから、もう一度「📷 カメラを起動して読み取る」を押してください。', priority: 106 },
  { category: 'スタンプ特典', question: 'スタンプが10個たまったらどうなりますか？', keywords: 'スタンプ,10個,達成,ガチャ,特典', answer: 'スタンプが10個たまると、ホーム画面から特典ガチャを回せます。結果はマイページの「🎁 スタンプ・特典履歴」で確認できます。ガチャ後はホーム画面の「🌸 新しいスタンプカードを取得」から次のカードを始められます。', priority: 104 },
  { category: 'スタンプ特典', question: '特典はどこで確認できますか？', keywords: '特典,どこ,確認,プレゼント,ガチャ', answer: '特典はマイページの「🎁 スタンプ・特典履歴」で確認できます。未使用の特典、使用済みの特典、受取期限を確認できます。', priority: 102 },
  { category: 'スタンプ特典', question: '特典の有効期限を知りたい', keywords: '特典,期限,有効期限,いつまで', answer: '特典の受取期限は、スタンプ10個を達成した日から1か月です。期限はマイページの「🎁 スタンプ・特典履歴」に表示されます。', priority: 100 },
  { category: '通知', question: '通知をオン・オフにしたい', keywords: '通知,オン,オフ,push,プッシュ,許可', answer: 'マイページの「🔔 通知設定」からオン・オフを切り替えられます。アプリ内でオンにしても届かない場合は、iPhone や Android 本体側の通知許可もご確認ください。', priority: 98 },
  { category: '通知', question: '通知が届かないときはどうすればいいですか？', keywords: '通知,届かない,push,プッシュ,こない', answer: 'まずマイページの「🔔 通知設定」がオンか確認してください。そのうえで、iPhone や Android 本体側の通知許可、通信状態、アプリの最新化をご確認ください。必要に応じて画面上部の🔄で最新情報を再取得してください。', priority: 96 },
  { category: '更新', question: '最新情報への更新方法を知りたい', keywords: '更新,最新,再読み込み,リロード,refresh,最新情報', answer: '画面上部の「🔄」ボタンを押すと、最新のNEWS、商品、カレンダー、メニュー、FAQ、注文履歴などを更新できます。通常の情報更新は再インストール不要です。', priority: 94 },
  { category: '更新', question: 'アップデートが必要と表示されたらどうすればいいですか？', keywords: 'アップデート,更新が必要,app store,最新版,バージョン', answer: '「アップデートが必要です」と表示された場合は、案内の「アップデートする」から最新版へ更新してください。画面上部の🔄は情報更新用で、必須アップデートの代わりにはなりません。', priority: 92 },
  { category: 'カレンダー', question: 'イベントカレンダーの見方を知りたい', keywords: 'カレンダー,イベント,予定,日程', answer: '下部メニューの「カレンダー」で予定を確認できます。左右の矢印で別の月に切り替えられ、日付を押すと詳細を確認できます。下部には今月のイベント一覧も表示されます。', priority: 90 },
  { category: 'カレンダー', question: 'カレンダーの記号の意味を知りたい', keywords: 'カレンダー,記号,意味,休,往,イ', answer: 'カレンダーでは、休＝休診日、往＝往診日、イ＝イベントを表しています。日付を押すと、その日の詳しい内容を確認できます。', priority: 88 },
  { category: 'NEWS', question: 'NEWSページの使い方を知りたい', keywords: 'NEWS,ニュース,お知らせ,記事,カテゴリ', answer: '下部メニューの「NEWS」を開くと記事一覧を確認できます。記事をタップすると詳細が開き、右上のカテゴリ選択で絞り込みもできます。', priority: 86 },
  { category: 'NEWS', question: 'お知らせ一覧の見方を知りたい', keywords: 'お知らせ一覧,通知一覧,拡声器,📢', answer: '画面上部の📢ボタンを押すと「お知らせ一覧」を開けます。ここでは NEWS、カレンダー、ショップ、ホームの更新情報を新しい順で確認できます。カテゴリの絞り込みもできます。', priority: 84 },
  { category: 'NEWS', question: 'NEWSのカテゴリ切り替え方法を知りたい', keywords: 'NEWS,カテゴリ,切り替え,絞り込み,全て', answer: 'NEWSページ右上のカテゴリ選択を押すと、カテゴリごとに絞り込みできます。「全て」を選ぶとすべての記事が表示されます。', priority: 82 },
  { category: 'NEWS', question: 'まゆみのつぶやきはどこで見られますか？', keywords: 'つぶやき,NEWS,カテゴリ,まゆみのつぶやき', answer: '「まゆみのつぶやき」は NEWS ページ右上のカテゴリ選択から「まゆみのつぶやき」を選ぶと表示されます。院長からの短いメッセージや大切なお知らせを確認できます。', priority: 80 },
  { category: 'NEWS', question: 'まゆみのブログとは何ですか？', keywords: 'まゆみのブログ,ブログ,外部ブログ', answer: '「まゆみのブログ」はマイページやホームの「🔗 公式サイト・SNS」から開ける外部ブログです。NEWS内の「まゆみのつぶやき」とは別の場所です。', priority: 78 },
  { category: 'リンク', question: '公式LINEやSNSの開き方を知りたい', keywords: 'LINE,ライン,instagram,facebook,ホームページ,公式サイト,SNS,問い合わせ', answer: 'ホーム画面またはマイページの「🔗 公式サイト・SNS」を開くと、公式ホームページ、Instagram、Facebook、公式LINE、まゆみのブログを選んで開けます。', priority: 76 },
  { category: '使い方サポート', question: '使い方チャットでは何を質問できますか？', keywords: 'チャット,サポート,ボット,相談,何が聞ける', answer: '使い方チャットでは、登録、復元、パスコード、注文、注文履歴、スタンプ、特典、通知、NEWS、お知らせ一覧、カレンダー、メニュー一覧、更新方法など、アプリの使い方について質問できます。診療相談や個別予約は公式LINEをご利用ください。', priority: 74 },
  { category: '引き継ぎ', question: '再インストールしたあとの入り方を知りたい', keywords: '再インストール,削除,アンインストール,復元,入り方', answer: 'アプリを入れ直したあとは、新規登録ではなく、初回画面またはログイン画面の「以前登録した方はこちら ↺」から復元してください。引き継ぎコードがある場合はコードで、ない場合はお名前と電話番号・生年月日・現在のパスコードのうち1つ以上で復元できます。', priority: 72 },
  { category: 'トラブル', question: '画面表示がおかしい・アプリが重いときはどうすればいいですか？', keywords: '表示されない,おかしい,崩れ,不具合,バグ,重い,遅い,フリーズ', answer: 'まず画面上部の🔄ボタンで最新情報を再取得してください。それでも改善しない場合は、アプリを一度閉じて再起動し、通信状態もご確認ください。再インストールが必要な場合は、先に引き継ぎコードを発行するか、復元方法を確認してから行ってください。', priority: 70 },
];
/* SUPPORT_FAQ_FALLBACK_END */

const SUPPORT_APP_GUIDE = [
  { category: 'アプリ全般', question: 'このアプリでできることを知りたい', keywords: 'アプリ,使い方,できること,何ができる,機能,全体', answer: 'このアプリでは、スタンプQRの読み取り、商品注文、カレンダー確認、NEWS確認、プロフィール変更、通知設定、スタンプ・特典履歴の確認ができます。予約や個別相談は公式LINEをご利用ください。', priority: 120 },
  { category: 'アプリ全般', question: 'アプリの画面構成を教えてください', keywords: '画面,構成,タブ,ナビ,メニュー,下部', answer: '画面下部にはホーム、ショップ、カレンダー、NEWS、マイページがあります。画面上部からはお知らせ一覧、カート、マイページ、更新ボタンも利用できます。', priority: 115 },
  { category: 'ホーム', question: 'ホーム画面の見方を知りたい', keywords: 'ホーム,トップ,home,見方', answer: 'ホーム画面では、スタンプカード、QR読み取り、おすすめ商品、最新のお知らせ、メニュー一覧、公式サイト・SNSへのリンクを確認できます。', priority: 100 },
  { category: 'プロフィール', question: 'プロフィール画像は変えられますか？', keywords: 'プロフィール画像,アイコン,写真,画像,変更', answer: 'はい。マイページの「✏️ プロフィールを編集」からアイコン画像を変更できます。変更後はマイページやホームの表示に反映されます。', priority: 98 },
  { category: 'プロフィール', question: 'プロフィールの変更方法を知りたい', keywords: 'プロフィール,変更,編集,名前,電話,住所', answer: 'マイページの「✏️ プロフィールを編集」を開き、必要な項目を変更して保存してください。', priority: 97 },
  { category: '通知', question: '通知をオフにしたい', keywords: '通知,オフ,push,プッシュ,解除', answer: 'マイページの「通知設定」でボタンをタップすると通知をオフにできます。端末側の通知設定も必要に応じてご確認ください。', priority: 96 },
  { category: '更新', question: '最新情報への更新方法を知りたい', keywords: '更新,最新,再読み込み,リロード,refresh', answer: '画面上部の「🔄」ボタンを押すと、最新のお知らせ・商品・カレンダー・FAQ・スタンプ・特典履歴などを更新できます。', priority: 95 },
  { category: '更新', question: 'アップデートが必要と表示されたらどうすればいいですか？', keywords: 'アップデート,更新が必要,app store,最新版', answer: '起動時に「アップデートが必要です」と表示された場合は、案内に従って最新版へ更新してください。軽微な情報更新は「🔄」ボタンで反映できます。', priority: 94 },
  { category: '予約', question: '予約はアプリからできますか？', keywords: '予約,よやく,line,予約方法', answer: 'このアプリから予約確定はできません。予約や個別相談は公式LINEからご連絡ください。メニュー一覧では内容確認のみできます。', priority: 93 },
  { category: 'スタンプ特典', question: 'スタンプが10個たまったらどうなりますか？', keywords: 'スタンプ,10個,達成,ガチャ,特典', answer: 'スタンプが10個たまると、ホーム画面から特典ガチャを回せます。ガチャ結果はマイページの「🎁 スタンプ・特典履歴」で確認できます。', priority: 92 },
  { category: 'NEWS', question: 'まゆみのブログとは何ですか？', keywords: 'まゆみのブログ,ブログ,外部ブログ', answer: '「まゆみのブログ」はマイページの「🔗 公式サイト・SNS」内にある外部ブログへのリンクです。院長の日々の想いや詳しい記事を読むことができます。ニュース内の「つぶやき」とは別物です。', priority: 88 },
  { category: 'NEWS', question: 'まゆみのつぶやきってどこで見られますか？', keywords: 'つぶやき,NEWS,カテゴリ,告知,メッセージ', answer: '「まゆみのつぶやき」はNEWSページの右上のカテゴリ選択で「まゆみのつぶやき」を選ぶと表示されます。アプリ内で手軽に読める院長からの短いメッセージや、大切なお知らせが配信されます。', priority: 86 }
];

const SUPPORT_APP_KEYWORDS = [
  'アプリ', '使い方', 'プロフィール', 'アイコン', '通知', '注文', '履歴', 'スタンプ', '特典',
  'ガチャ', '予約', 'メニュー', 'カレンダー', 'news', '更新', 'アップデート', '公式line',
  'ブログ', 'つぶやき', 'まゆみのブログ', 'まゆみのつぶやき', 'パスコード', 'ログイン',
  '引き継ぎ', '引き継ぎコード', '復元', '機種変更', '再インストール', 'ホーム画面',
  '会員id', '会員番号', 'お知らせ一覧', 'カート', '受け取り', '受取', 'カテゴリ'
];

function normalizeProductCategory(category) {
  return String(category || '')
    .trim()
    .replace(/[　\\s]+/g, '')
    .toLowerCase();
}

function normalizeProductEntry(product, index) {
  const item = product && typeof product === 'object' ? product : {};
  const iconImages = normalizeManagedImageList(item.imageUrls || item.iconImages || item.icon);
  const descriptionImageUrls = normalizeManagedImageList(item.descriptionImageUrls || item.descriptionImage);
  return {
    category: String(item.category || ''),
    name: String(item.name || ''),
    price: Number(item.price || 0),
    icon: iconImages[0] || String(item.icon || ''),
    imageUrls: iconImages,
    bg: String(item.bg || ['c1', 'c2', 'c3', 'c4'][index % 4] || 'c1'),
    imgKey: String(item.imgKey || ''),
    description: String(item.description || ''),
    descriptionImage: descriptionImageUrls[0] || String(item.descriptionImage || ''),
    descriptionImageUrls: descriptionImageUrls,
    updatedAt: String(item.updatedAt || ''),
    noticeListedAt: String(item.noticeListedAt || ''),
    noticeStatus: normalizeNoticeVisibilityStatus(item.noticeStatus)
  };
}

function getProductPricing(product, quantity) {
  const qty = Math.max(1, Number(quantity || 1));
  const regularUnitPrice = Number(product && product.price || 0);
  const regularTotal = regularUnitPrice * qty;

  if (isDashiProductName(product && product.name)) {
    const dashiPricing = calculateDashiPricing(qty);
    return {
      regularUnitPrice: regularUnitPrice,
      unitPrice: Number(dashiPricing.avgUnitPrice || regularUnitPrice || 0),
      originalTotal: regularTotal,
      total: Number(dashiPricing.totalRevenue || regularTotal || 0)
    };
  }

  return {
    regularUnitPrice: regularUnitPrice,
    unitPrice: regularUnitPrice,
    originalTotal: regularTotal,
    total: regularUnitPrice * qty
  };
}

function buildProductPriceMarkup(product, quantity, options) {
  const opts = options || {};
  const pricing = getProductPricing(product, quantity);
  const mode = opts.mode === 'total' ? 'total' : 'unit';
  const includeTax = opts.includeTax !== false;
  const effectiveAmount = mode === 'total' ? pricing.total : pricing.unitPrice;
  const taxLabel = includeTax ? '<small>（税込）</small>' : '';
  return `¥${effectiveAmount.toLocaleString()}${taxLabel}`;
}

async function loadProducts() {
  const data = await getFromGAS('getProducts');
  if (data && data.status === 'ok' && Array.isArray(data.products) && data.products.length) {
    PRODUCTS.splice(0, PRODUCTS.length);
    data.products.forEach(function (product, index) {
      PRODUCTS.push(normalizeProductEntry(product, index));
    });
  }

  const fixedGridIds = ['grid-recommended', 'grid-tea', 'grid-bath', 'grid-dashi', 'grid-soap-other', 'grid-pad'];
  fixedGridIds.forEach(function (id) {
    const node = document.getElementById(id);
    if (node) node.innerHTML = '';
  });
  document.querySelectorAll('.dynamic-cat').forEach(function (node) {
    node.remove();
  });

  const keywordMap = [
    { key: 'よもぎ茶', gridId: 'grid-tea' },
    { key: 'よもぎ入浴剤', gridId: 'grid-bath' },
    { key: '入浴剤', gridId: 'grid-bath' },
    { key: '天然だし調理粉', gridId: 'grid-dashi' },
    { key: '天然だし調味粉', gridId: 'grid-dashi' },
    { key: '天然だし', gridId: 'grid-dashi' },
    { key: 'だし', gridId: 'grid-dashi' },
    { key: '食品', gridId: 'grid-dashi' },
    { key: '石鹸', gridId: 'grid-soap-other' },
    { key: 'せっけん', gridId: 'grid-soap-other' },
    { key: 'ソープ', gridId: 'grid-soap-other' },
    { key: '冷え取りパット', gridId: 'grid-pad' },
    { key: '冷え取りパッド', gridId: 'grid-pad' },
    { key: '冷え取り', gridId: 'grid-pad' }
  ].map(function (item) {
    return {
      key: normalizeProductCategory(item.key),
      gridId: item.gridId
    };
  });

  const getOrCreateContainer = function (category) {
    const norm = normalizeProductCategory(category);
    const match = keywordMap.find(function (item) { return norm.indexOf(item.key) !== -1; });
    if (match) {
      return document.getElementById(match.gridId);
    }

    const dynamicId = 'grid-dynamic-' + encodeURIComponent(norm || String(category || 'other')).replace(/%/g, '');
    let container = document.getElementById(dynamicId);
    if (!container) {
      const shopSection = document.querySelector('#page-shop .main .section');
      if (!shopSection) return null;

      const label = document.createElement('div');
      label.className = 'cat-label dynamic-cat';
      label.style.marginTop = '24px';
      label.textContent = '📦 ' + (category || 'その他');

      container = document.createElement('div');
      container.id = dynamicId;
      container.className = 'shop-grid dynamic-cat';

      shopSection.appendChild(label);
      shopSection.appendChild(container);
    }
    return container;
  };

  const createProductCard = function (product, index) {
    const card = document.createElement('div');
    card.className = 'shop-card';
    card.setAttribute('onclick', `openProductModal(${index})`);
    const imageSource = (product.icon && (product.icon.startsWith('http') || product.icon.startsWith('data:')))
      ? getContentDisplayImageUrl(product.icon)
      : (PROD_IMAGES[product.imgKey] || `https://placehold.jp/24/c18151/ffffff/150x150.png?text=${encodeURIComponent(product.name)}`);
    const iconBadge = product.icon && !/^https?:\/\//i.test(product.icon) && !/^data:/i.test(product.icon)
      ? `<span class="shop-icon-badge">${escapeHtml(product.icon)}</span>`
      : '';
    const unreadBadge = buildUnreadBadgeHtml('product', product);
    const favoriteBadge = isFavoriteKey(buildContentItemKey('product', product)) ? '<span class="item-favorite-badge">★</span>' : '';
    card.innerHTML = `
      <div class="shop-img ${product.bg}">
        <img src="${imageSource}" alt="${product.name}">
        ${iconBadge}
        ${unreadBadge}
        ${favoriteBadge}
      </div>
      <div class="shop-info">
        <div class="shop-name">${product.name}</div>
        <div class="shop-price">${buildProductPriceMarkup(product, 1, { mode: 'unit', includeTax: true, showPeriod: true })}</div>
      </div>
    `;
    return card;
  };

  const recommended = document.getElementById('grid-recommended');
  if (recommended) {
    PRODUCTS.slice(0, 4).forEach(function (product, index) {
      recommended.appendChild(createProductCard(product, index));
    });
  }

  PRODUCTS.forEach(function (product, index) {
    const container = getOrCreateContainer(product.category);
    if (container) {
      container.appendChild(createProductCard(product, index));
    }
  });

  if (document.getElementById('page-notices').classList.contains('active')) {
    renderPushNotices();
  }
}

// ===== ブログ =====
async function loadBlog() {
  const data = await getFromGAS('getNews');
  if (data && data.status === 'ok' && data.news) {
    allBlogCategories = Array.isArray(data.categories) ? data.categories : [];
    blogItems = normalizeBlogItems(data.news, allBlogCategories);
  } else {
    allBlogCategories = [];
    blogItems = normalizeBlogItems(FALLBACK_BLOG.slice(), []);
  }

  updateBlogCategoryFilters();
  updateMenuCategoryFilter();
  renderBlogList('homeNewsList', 3);

  if (document.getElementById('page-blog').classList.contains('active')) {
    renderDividedBlogList();
  }
  if (document.getElementById('page-notices').classList.contains('active')) {
    renderPushNotices();
  }
  updateNavBadges();
}

async function loadPushNotices() {
  const data = await getFromGAS('getPushNotices');
  if (data && data.status === 'ok' && Array.isArray(data.notices)) {
    pushNotices = data.notices.map(function (notice) {
      return {
        date: notice && notice.date || 0,
        title: String(notice && notice.title || 'お知らせ'),
        body: String(notice && notice.body || '')
      };
    }).filter(function (notice) {
      return notice.title || notice.body;
    });
  } else {
    pushNotices = [];
  }

  if (document.getElementById('page-notices').classList.contains('active')) {
    renderPushNotices();
  }
  updateNavBadges();
}

async function loadCurrentMonthlyReward() {
  CURRENT_MONTHLY_REWARD = null;
  STAMP_REWARD_CONFIG = [];
  return null;
}

async function loadStampRewards() {
  await loadCurrentMonthlyReward();
  const localStatus = getLocalRewardStatus();
  if (_profile && _profile.memberId) {
    const remote = await getFromGAS('getUserRewardStatus', { memberId: _profile.memberId });
    if (remote && remote.status === 'ok' && remote.rewardStatus) {
      const remoteStatus = getComparableRewardStatus(remote.rewardStatus);
      if (hasMeaningfulRewardStatus(remoteStatus)) {
        const mergedStatus = mergeRewardStatuses(remoteStatus, localStatus);
        applyRewardStatusLocally(mergedStatus);
        lastSyncedRewardStatus = mergedStatus;
        if (!rewardStatusEquals(mergedStatus, remoteStatus)) {
          await syncRewardStatus(true);
        }
      } else if (hasMeaningfulRewardStatus(localStatus)) {
        await syncRewardStatus(true);
      }
    } else if (hasMeaningfulRewardStatus(localStatus)) {
      await syncRewardStatus(true);
    }
  } else {
    applyRewardStatusLocally(localStatus);
  }
  renderEarnedRewards();
}

async function loadSurveys() {
  return [];
}

function renderSurveys() {
  return;
}

function renderCalendarEventLists() {
  const dailyList = document.getElementById('calendar-events-list');
  const monthlyList = document.getElementById('calendar-monthly-events-list');
  if (!dailyList || !monthlyList) return;

  const selectedKey = formatCalendarDateKey(selectedDate);
  const selectedEvents = getCalendarEventsByDate(selectedKey);
  const monthYearStr = currentMonthDate.getFullYear() + '-' + String(currentMonthDate.getMonth() + 1).padStart(2, '0');
  const monthlyEvents = calendarData.filter(function (event) {
    return String(event.date || '').slice(0, 7) === monthYearStr &&
      !isCalendarHolidayEvent(event) &&
      !isCalendarVisitEvent(event);
  }).sort(compareCalendarEventsByDateAsc);

  const displayDate = formatCalendarDisplayDate(selectedKey);
  const dailyHeader = `<h3 style="margin:0; font-size:15px; color:#9d5b5b; padding-bottom:8px; border-bottom:1px dashed #f7d2d2; margin-bottom:12px;">${escapeHtml(displayDate)}の予定</h3>`;

  dailyList.innerHTML = dailyHeader + (selectedEvents.length
    ? selectedEvents.map(function (event) {
      return buildCalendarEventListItem(event, false);
    }).join('')
    : '<div class="no-evts">この日の予定はありません</div>');

  monthlyList.innerHTML = monthlyEvents.length
    ? monthlyEvents.map(function (event) {
      return buildCalendarEventListItem(event, true);
    }).join('')
    : '<div class="no-evts">今月の予定はありません</div>';
}

function renderCalendar() {
  const monthYear = document.getElementById('calendar-month-year');
  const grid = document.getElementById('calendar-grid-info');
  if (!monthYear || !grid) return;

  monthYear.textContent = `${currentMonthDate.getFullYear()}年${currentMonthDate.getMonth() + 1}月`;
  const firstDay = new Date(currentMonthDate.getFullYear(), currentMonthDate.getMonth(), 1);
  const startWeekday = firstDay.getDay();
  const daysInMonth = new Date(currentMonthDate.getFullYear(), currentMonthDate.getMonth() + 1, 0).getDate();
  const weekdayLabels = ['日', '月', '火', '水', '木', '金', '土'];
  const todayKey = formatCalendarDateKey(new Date());

  let html = weekdayLabels.map(function (label, idx) {
    let cls = 'cal-day-name';
    if (idx === 0) cls += ' sun';
    if (idx === 6) cls += ' sat';
    return `<div class="${cls}">${label}</div>`;
  }).join('');

  for (let i = 0; i < startWeekday; i++) {
    html += '<div class="cal-day empty"></div>';
  }

  for (let day = 1; day <= daysInMonth; day++) {
    const d = new Date(currentMonthDate.getFullYear(), currentMonthDate.getMonth(), day);
    const dateKey = formatCalendarDateKey(d);
    const dayEvents = getCalendarEventsByDate(dateKey);
    const isSelected = selectedDate.getFullYear() === d.getFullYear() &&
      selectedDate.getMonth() === d.getMonth() &&
      selectedDate.getDate() === d.getDate();
    const isToday = dateKey === todayKey;
    const weekday = d.getDay();

    const classes = ['cal-day'];
    if (weekday === 0) classes.push('sun');
    if (weekday === 6) classes.push('sat');
    if (isToday) classes.push('today');
    if (isSelected) classes.push('selected');

    html += `<button type="button" class="${classes.join(' ')}" onclick="selectCalendarDate(${d.getFullYear()}, ${d.getMonth()}, ${day})"><span class="cal-day-number">${day}</span>${buildCalendarDayMarkers(dayEvents)}</button>`;
  }

  grid.innerHTML = html;
  renderCalendarEventLists();
}

function selectCalendarDate(year, month, day) {
  selectedDate = new Date(year, month, day);
  currentMonthDate = new Date(year, month, 1);
  renderCalendar();
}

function prevMonth() {
  currentMonthDate = new Date(currentMonthDate.getFullYear(), currentMonthDate.getMonth() - 1, 1);
  selectedDate = new Date(currentMonthDate.getFullYear(), currentMonthDate.getMonth(), 1);
  renderCalendar();
}

function nextMonth() {
  currentMonthDate = new Date(currentMonthDate.getFullYear(), currentMonthDate.getMonth() + 1, 1);
  selectedDate = new Date(currentMonthDate.getFullYear(), currentMonthDate.getMonth(), 1);
  renderCalendar();
}

function openCalendarEventDetail(idx) {
  const event = calendarData[idx];
  if (!event) return;
  markContentItemSeen('calendar', event);
  const detail = document.getElementById('calendarEventDetailContent');
  if (!detail) return;
  const safeTitle = escapeHtml(event.title || '');
  const safeDate = escapeHtml(formatCalendarDisplayDate(event.date || ''));
  const safeDesc = renderManagedTextHtml(event.desc || '', {
    inlineImageClass: 'blog-inline-image blog-inline-image--detail',
    inlineImageAlt: (event.title || 'イベント画像')
  }) || '詳細はありません。';
  const safeColor = getCalendarSafeColor(event.color);
  const detailImageHtml = buildDetailImageGalleryHtml(event.imageUrls || event.image, event.title || 'Calendar Image');
  detail.innerHTML = `
    ${detailImageHtml}
    <div class="blog-detail-title" style="margin-bottom:12px;">${safeTitle}</div>
    <div class="blog-detail-body managed-rich-text" style="font-size:14px; line-height:1.8; color:var(--text-main);">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px;color:var(--text-light);">
        <span style="display:inline-block;width:10px;height:10px;border-radius:999px;background:${safeColor};flex-shrink:0;"></span>
        <span>${safeDate}</span>
      </div>
      <div>${safeDesc}</div>
    </div>
    <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:16px;">
      ${buildFavoriteActionMarkup('calendar', event)}
      <button class="btn secondary favorite-action-btn" type="button" onclick="downloadCalendarEventIcs(${idx})">端末カレンダーに追加</button>
      <button class="btn secondary favorite-action-btn" type="button" onclick="openGoogleCalendarEvent(${idx})">Googleカレンダーに追加</button>
    </div>
  `;
  renderCalendarEventLists();
  updateNavBadges();
  openModal('calendarEventDetailModal');
}

function escapeIcsText(value) {
  return String(value || '')
    .replace(/\\/g, '\\\\')
    .replace(/\n/g, '\\n')
    .replace(/,/g, '\\,')
    .replace(/;/g, '\\;');
}

function normalizeCalendarEventDateParts(dateValue) {
  if (!dateValue) return null;
  if (dateValue instanceof Date && !Number.isNaN(dateValue.getTime())) {
    return {
      year: dateValue.getFullYear(),
      month: dateValue.getMonth() + 1,
      day: dateValue.getDate()
    };
  }
  const raw = String(dateValue).trim();
  if (!raw) return null;
  const datePart = raw.split(/[ T]/)[0];
  const directMatch = datePart.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/);
  if (directMatch) {
    return {
      year: Number(directMatch[1]),
      month: Number(directMatch[2]),
      day: Number(directMatch[3])
    };
  }
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return null;
  return {
    year: parsed.getFullYear(),
    month: parsed.getMonth() + 1,
    day: parsed.getDate()
  };
}

function getCalendarEventDateRange(event) {
  const parts = normalizeCalendarEventDateParts(event && event.date);
  if (!parts) return null;
  const startDate = String(parts.year).padStart(4, '0')
    + String(parts.month).padStart(2, '0')
    + String(parts.day).padStart(2, '0');
  const endDateObj = new Date(parts.year, parts.month - 1, parts.day + 1);
  const endDate = endDateObj.getFullYear()
    + String(endDateObj.getMonth() + 1).padStart(2, '0')
    + String(endDateObj.getDate()).padStart(2, '0');
  return { startDate, endDate };
}

function openGoogleCalendarEvent(idx) {
  const event = calendarData[idx];
  if (!event) return;
  const range = getCalendarEventDateRange(event);
  if (!range) {
    showToast('このイベントはGoogleカレンダーへ追加できません');
    return;
  }
  const url = 'https://calendar.google.com/calendar/render?action=TEMPLATE'
    + '&text=' + encodeURIComponent(event.title || 'まゆみ助産院イベント')
    + '&dates=' + encodeURIComponent(range.startDate + '/' + range.endDate)
    + '&details=' + encodeURIComponent(event.desc || '');
  const opened = window.open(url, '_blank', 'noopener');
  if (!opened) {
    window.location.href = url;
  }
}

function downloadCalendarEventIcs(idx) {
  const event = calendarData[idx];
  if (!event) return;
  const range = getCalendarEventDateRange(event);
  if (!range) {
    showToast('このイベントは端末カレンダーへ追加できません');
    return;
  }
  const icsBody = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Mayumi App//JP',
    'BEGIN:VEVENT',
    'UID:' + escapeIcsText(buildContentItemKey('calendar', event)),
    'DTSTAMP:' + new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+Z$/, 'Z'),
    'DTSTART;VALUE=DATE:' + range.startDate,
    'DTEND;VALUE=DATE:' + range.endDate,
    'SUMMARY:' + escapeIcsText(event.title || 'まゆみ助産院イベント'),
    'DESCRIPTION:' + escapeIcsText(event.desc || ''),
    'END:VEVENT',
    'END:VCALENDAR'
  ].join('\r\n');
  const blob = new Blob([icsBody], { type: 'text/calendar;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = (event.title || 'mayumi-event') + '.ics';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(function () {
    URL.revokeObjectURL(url);
  }, 1500);
  showToast('端末カレンダー追加用のファイルを開きました');
}

async function loadCalendar() {
  const data = await getFromGAS('getCalendar');
  if (data && data.status === 'ok' && Array.isArray(data.events)) {
    calendarData = data.events.map(normalizeCalendarEventEntry);
  } else {
    calendarData = [];
  }
  renderCalendar();
  if (document.getElementById('page-notices').classList.contains('active')) {
    renderPushNotices();
  }
}

function formatCalendarDateKey(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
  return date.getFullYear() + '-' +
    String(date.getMonth() + 1).padStart(2, '0') + '-' +
    String(date.getDate()).padStart(2, '0');
}

function formatCalendarDisplayDate(dateValue) {
  return formatCustomerDateYmd(dateValue);
}

function formatCalendarCompactDate(dateValue) {
  return formatCustomerDateYmd(dateValue);
}

function getCalendarEventsByDate(dateKey) {
  return calendarData.filter(function (event) {
    const eventDate = String(event.date || '').split(/[ T]/)[0];
    return eventDate === dateKey;
  }).sort(compareCalendarEventsByDateAsc);
}

function compareCalendarEventsByDateAsc(a, b) {
  const aRange = getCalendarEventDateRange(a);
  const bRange = getCalendarEventDateRange(b);
  const aStart = String(aRange && aRange.startDate || '');
  const bStart = String(bRange && bRange.startDate || '');
  if (aStart !== bStart) return aStart.localeCompare(bStart);

  const aTitle = String(a && a.title || '');
  const bTitle = String(b && b.title || '');
  if (aTitle !== bTitle) return aTitle.localeCompare(bTitle, 'ja');

  const aDesc = String(a && a.desc || '');
  const bDesc = String(b && b.desc || '');
  return aDesc.localeCompare(bDesc, 'ja');
}

function isCalendarHolidayEvent(event) {
  const title = String(event && event.title || '');
  const desc = String(event && event.desc || '');
  return /休診|休み|休業/.test(title + ' ' + desc);
}

function isCalendarVisitEvent(event) {
  const title = String(event && event.title || '');
  return title.indexOf('往診') !== -1;
}

function getCalendarSafeColor(colorValue) {
  const color = String(colorValue || '').trim();
  if (/^#[0-9a-fA-F]{3,8}$/.test(color)) return color;
  if (/^[a-zA-Z]+$/.test(color)) return color;
  return '#e57373';
}

function buildCalendarDayMarkers(dayEvents) {
  if (!dayEvents || !dayEvents.length) return '';
  const holidayEvents = dayEvents.filter(isCalendarHolidayEvent);
  const visitEvents = dayEvents.filter(isCalendarVisitEvent);
  const normalEvents = dayEvents.filter(function (event) {
    return !isCalendarHolidayEvent(event) && !isCalendarVisitEvent(event);
  });

  let html = '';
  if (holidayEvents.length) {
    html += '<div class="cal-holiday-tag">休診</div>';
  }
  if (visitEvents.length) {
    visitEvents.forEach(function (ev) {
      const color = getCalendarSafeColor(ev.color);
      html += `<div class="cal-visit-tag" style="background:${color}">往診</div>`;
    });
  }

  if (normalEvents.length) {
    html += '<div class="cal-evt-container">';
    normalEvents.slice(0, 3).forEach(function (event) {
      html += `<span class="cal-evt-tag" style="background:${getCalendarSafeColor(event.color)};">イ</span>`;
    });
    if (normalEvents.length > 3) {
      html += '<span class="cal-evt-tag" style="background:#ccb3b3;">+</span>';
    }
    html += '</div>';
  }
  return html;
}

function buildCalendarEventListItem(event, showDate) {
  const eventIndex = calendarData.indexOf(event);
  const safeTitle = escapeHtml(event.title || '');
  const safeDesc = renderManagedTextHtml(event.desc || '', {
    stripInlineImages: true,
    preserveLineBreaks: false
  }) || '詳細はありません。';
  const safeDate = escapeHtml(formatCalendarCompactDate(event.date || ''));
  const safeImage = getContentDisplayImageUrl(event.image || '');
  const safeColor = getCalendarSafeColor(event.color);
  const dateHtml = showDate ? `<div class="cal-evt-date">${safeDate}</div>` : '';
  const imageHtml = safeImage
    ? `<div class="cal-evt-img-box"><img src="${safeImage}" alt="${safeTitle}" loading="lazy"></div>`
    : '<div class="cal-evt-img-box"><div class="cal-evt-img-placeholder">📅</div></div>';
  return `<div class="cal-evt-item" onclick="openCalendarEventDetail(${eventIndex})">${dateHtml}${imageHtml}<div class="cal-evt-color" style="background:${safeColor};"></div><div class="cal-evt-details"><div class="cal-evt-title">${safeTitle} ${buildUnreadBadgeHtml('calendar', event)}</div><div class="cal-evt-desc">${safeDesc}</div></div></div>`;
}

// ===== お知らせ一覧用 統合ハッシュ =====
// お知らせ一覧に表示する管理対象をまとめてハッシュ化。
// これが変わる → 拡声器バッジを表示
function computeNoticeListHash() {
  try {
    const visibleNoticePart = buildNoticeFeedItems().map(function (item) {
      return [
        item.kind || '',
        item.timestamp || '',
        item.dateLabel || '',
        item.category || '',
        item.title || '',
        item.body || '',
        item.image || ''
      ].join('_');
    }).join('|');
    return visibleNoticePart;
  } catch (e) {
    console.error('[同期] ハッシュ計算エラー:', e);
    return '';
  }
}

function computeProductListHash() {
  return (PRODUCTS || []).map(function (item) {
    return [
      item.updatedAt || '',
      item.category || '',
      item.name || '',
      item.price || '',
      item.description || '',
      item.icon || '',
      item.descriptionImage || ''
    ].join('_');
  }).join('|');
}


async function tryTask(taskName, taskFn) {
  try {
    console.log(`[同期] ${taskName} の取得中...`);
    return await taskFn();
  } catch (e) {
    console.error(`[同期エラー] ${taskName} の取得に失敗しました:`, e);
    return null;
  }
}

async function fetchLatestManagedContent(options) {
  const opts = options || {};
  const currentPage = document.querySelector('.page.active')?.id || '';

  const tasks = [
    tryTask('ニュース', () => loadBlog()),
    tryTask('商品', () => loadProducts()),
    tryTask('カレンダー', () => loadCalendar()),
    tryTask('通知', () => loadPushNotices()),
    tryTask('アンケート', () => loadSurveys()),
    tryTask('特典', () => loadStampRewards())
  ];

  if (opts.refreshSupportFaq) tasks.push(tryTask('FAQ', () => loadSupportFaq(true)));
  tasks.push(tryTask('メニュー', () => loadMenus({ silent: currentPage !== 'page-menu-list' })));
  if (opts.refreshOrderHistory || currentPage === 'page-mypage') tasks.push(tryTask('履歴', () => renderOrderHistory()));

  await Promise.all(tasks);

  try {
    updateNavBadges();
  } catch (e) {
    console.error('[同期] バッジ表示更新エラー:', e);
  }
}

async function refreshNoticeFeed() {
  await Promise.all([
    tryTask('ニュース', () => loadBlog()),
    tryTask('商品', () => loadProducts()),
    tryTask('カレンダー', () => loadCalendar()),
    tryTask('メニュー', () => loadMenus({ silent: true, allowMissingContainer: true }))
  ]);

  try {
    updateNavBadges();
  } catch (e) {
    console.error('[お知らせ同期] バッジ表示更新エラー:', e);
  }
}

async function refreshAppData() {
  const btn = document.getElementById('refresh-btn');
  const overlay = document.getElementById('refresh-overlay');
  if (btn) btn.classList.add('spinning');
  if (overlay) overlay.classList.add('active');

  try {
    console.log('[更新] アプリ情報の同期を開始します...');
    let shouldReloadForCodeUpdate = false;

    const versionGate = await ensureSupportedAppVersion();
    if (versionGate.blocked) {
      alert('アプリ本体の更新が必要です');
      return;
    }
    if (versionGate.needsWebUpdate) shouldReloadForCodeUpdate = true;

    if ('serviceWorker' in navigator) {
      try {
        const registrations = await navigator.serviceWorker.getRegistrations();
        for (let registration of registrations) {
          await registration.update();
          if (registration.waiting) {
            registration.waiting.postMessage({ type: 'SKIP_WAITING' });
            shouldReloadForCodeUpdate = true;
          }
        }
      } catch (e) {
        console.error('[更新] サービスワーカー更新失敗:', e);
      }
    }

    await fetchLatestManagedContent({
      refreshSupportFaq: true,
      refreshMenus: true,
      refreshOrderHistory: true
    });

    if (shouldReloadForCodeUpdate) {
      await applyPendingAppUpdate();
      return;
    } else {
      hideAppUpdateBanner(false);
      showToast('最新情報を反映しました ✨');
    }
  } catch (e) {
    console.error('[更新] 同期中に致命的なエラーが発生しました:', e);
    showToast('同期中に問題が発生しました。しばらく経ってから再度お試しください。');
  } finally {
    if (btn) btn.classList.remove('spinning');
    if (overlay) overlay.classList.remove('active');
  }
}

function renderBlogList(containerId, limit, filterType = null, filterCategory = '全て') {
  const el = document.getElementById(containerId);
  if (!el) return;

  let items = blogItems;
  if (filterType) {
    items = items.filter(i => getBlogItemType(i) === filterType);
  }
  if (filterCategory && filterCategory !== '全て') {
    items = items.filter(i => i.category === filterCategory);
  }

  items = items.slice(0, limit);

  if (!items.length) {
    el.innerHTML = '<div style="text-align:center;font-size:13px;color:var(--text-light);padding:28px 0">記事がありません</div>';
    return;
  }
  el.innerHTML = '';
  items.forEach(item => {
    const div = document.createElement('div');
    div.className = 'blog-card';
    const displayDate = formatCustomerDateYmd(item.date);

    let bodyHtml = '';
    if (item.body) {
      const previewText = String(item.body || '')
        .replace(/(^|\n)\s*📷\s*https?:\/\/\S+/g, '$1')
        .replace(/\n{3,}/g, '\n\n')
        .trim();
      if (previewText) {
        bodyHtml = `<div class="blog-body-preview managed-rich-text">${renderManagedTextHtml(previewText, { stripInlineImages: true })}</div>`;
      }
    }

    div.innerHTML = `<div class="blog-inner"><div class="blog-icon">${escapeHtml(item.icon || '📢')}</div><div style="flex:1"><div class="blog-meta"><span class="blog-date">${escapeHtml(displayDate)}</span><span class="blog-cat">${escapeHtml(item.category || '')}</span>${buildUnreadBadgeHtml('blog', item)}</div><div class="blog-title">${escapeHtml(item.title || '')}</div>${bodyHtml}</div></div>`;
    div.onclick = () => openBlogDetail(item);
    el.appendChild(div);
  });
}

function renderDividedBlogList() {
  const newsCat = document.getElementById('newsCategoryFilter')?.value || '全て';
  renderBlogList('newsPageList', 999, null, newsCat);
}

function normalizeManagedCategoryType(type) {
  const value = String(type || '').trim();
  return ['お知らせ', 'ブログ', 'メニュー', '通知'].indexOf(value) !== -1 ? value : '';
}

function normalizeCategoryLookupKey(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[ぁ-ん]/g, function (char) {
      return String.fromCharCode(char.charCodeAt(0) + 0x60);
    })
    .replace(/[\s\u3000_\-\/]/g, '');
}

function getManagedCategoryNamesByTypes(types) {
  const allowed = {};
  (types || []).forEach(function (type) {
    const normalized = normalizeManagedCategoryType(type);
    if (normalized) allowed[normalized] = true;
  });

  const names = [];
  const seen = {};
  (allBlogCategories || []).forEach(function (item) {
    const type = normalizeManagedCategoryType(item && item.type);
    const name = String(item && item.name || '').trim();
    if (!name || !allowed[type] || seen[name]) return;
    seen[name] = true;
    names.push(name);
  });
  return names;
}

function getManagedNoticeCategoryMap() {
  const defaults = [
    { kind: 'blog', fallback: 'NEWS', aliases: ['news', 'ニュース'] },
    { kind: 'calendar', fallback: 'カレンダー', aliases: ['カレンダー', 'calendar'] },
    { kind: 'product', fallback: 'ショップ', aliases: ['ショップ', 'shop'] },
    { kind: 'menu', fallback: 'ホーム', aliases: ['ホーム', 'home'] }
  ];
  const masterNames = getManagedCategoryNamesByTypes(['通知']);
  const entries = masterNames.map(function (name) {
    return { name: name, key: normalizeCategoryLookupKey(name) };
  });
  const assigned = {};
  const used = {};

  defaults.forEach(function (item) {
    const lookupKeys = item.aliases.map(normalizeCategoryLookupKey).concat([normalizeCategoryLookupKey(item.fallback)]);
    const matchedIndex = entries.findIndex(function (entry, index) {
      return !used[index] && lookupKeys.indexOf(entry.key) !== -1;
    });
    if (matchedIndex >= 0) {
      assigned[item.kind] = entries[matchedIndex].name;
      used[matchedIndex] = true;
    }
  });

  const remaining = entries.filter(function (entry, index) {
    return !used[index];
  }).map(function (entry) {
    return entry.name;
  });

  defaults.forEach(function (item) {
    if (!assigned[item.kind] && remaining.length) {
      assigned[item.kind] = remaining.shift();
    }
    if (!assigned[item.kind]) assigned[item.kind] = item.fallback;
  });

  return assigned;
}

function resolveManagedNoticeCategory(kind, fallback) {
  const map = getManagedNoticeCategoryMap();
  return map[kind] || fallback || '';
}

function normalizeNoticeVisibilityStatus(status) {
  return String(status || '').trim() === '非公開' ? '非公開' : '公開';
}

function isNoticeFeedEntryVisible(item) {
  return normalizeNoticeVisibilityStatus(item && item.noticeStatus) === '公開';
}

function updateBlogCategoryFilters() {
  const filter = document.getElementById('newsCategoryFilter');
  if (!filter) return;

  const previousValue = filter.value || '全て';
  const categoryNames = [];
  const seen = {};

  getManagedCategoryNamesByTypes(['お知らせ', 'ブログ']).forEach(function (name) {
    if (!name || seen[name]) return;
    seen[name] = true;
    categoryNames.push(name);
  });

  (blogItems || []).forEach(function (item) {
    const name = String(item && item.category || '').trim();
    if (!name || seen[name]) return;
    const type = getBlogItemType(item);
    if (type !== 'お知らせ' && type !== 'ブログ') return;
    seen[name] = true;
    categoryNames.push(name);
  });

  filter.innerHTML = '<option value="全て">カテゴリ: 全て</option>' + categoryNames.map(function (name) {
    return `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`;
  }).join('');

  filter.value = seen[previousValue] ? previousValue : '全て';
}

function renderPushNotices() {
  const container = document.getElementById('noticeList');
  if (!container) return;

  let items = buildNoticeFeedItems();

  updateNoticeCategoryFilter(items);
  const filterValue = document.getElementById('noticeCategoryFilter') ? document.getElementById('noticeCategoryFilter').value : '';

  if (filterValue) {
    items = items.filter(function (item) {
      return item.category === filterValue;
    });
  }

  if (items.length === 0) {
    container.innerHTML = '<div style="text-align:center; padding:40px 20px; color:var(--text-mid);">条件に一致するお知らせはありません</div>';
    return;
  }

  container.innerHTML = '';
  items.forEach(function (item) {
    const card = document.createElement('div');
    card.className = 'blog-card';
    card.innerHTML = `
      <div class="blog-inner">
        <div class="blog-icon">${escapeHtml(item.icon || '📢')}</div>
        <div style="flex:1">
          <div class="blog-meta">
            <span class="blog-cat">${escapeHtml(item.category)}</span>
            <span class="blog-date">${escapeHtml(item.dateLabel || '')}</span>
            ${buildUnreadBadgeHtml(item.kind, item)}
          </div>
          <div class="blog-title">${escapeHtml(item.title)}</div>
          <div class="blog-body-preview managed-rich-text">${renderManagedTextHtml(item.body || '', { stripInlineImages: true })}</div>
        </div>
      </div>
    `;
    card.onclick = function () { openNoticeFeedItem(item); };
    container.appendChild(card);
  });
}

function openNoticeFeedItem(item) {
  if (!item) return;
  if (item.kind === 'blog') {
    const blog = blogItems.find(function (entry) {
      return buildContentItemKey('blog', entry) === item.sourceKey;
    });
    if (blog) {
      openBlogDetail(blog);
      return;
    }
  }
  if (item.kind === 'calendar') {
    const calendarIndex = calendarData.findIndex(function (entry) {
      return buildContentItemKey('calendar', entry) === item.sourceKey;
    });
    if (calendarIndex >= 0) {
      openCalendarEventDetail(calendarIndex);
      return;
    }
  }
  if (item.kind === 'product') {
    const productIndex = PRODUCTS.findIndex(function (entry) {
      return buildContentItemKey('product', entry) === item.sourceKey;
    });
    if (productIndex >= 0) {
      switchPage('shop');
      openProductModal(productIndex);
      return;
    }
  }
  if (item.kind === 'menu') {
    const menuIndex = USER_MENUS.findIndex(function (entry) {
      return buildContentItemKey('menu', entry) === item.sourceKey;
    });
    if (menuIndex >= 0) {
      switchPage('menu-list');
      openMenuDetail(menuIndex);
      return;
    }
  }
  markContentItemSeen(item.kind || 'blog', item);
  openBlogDetail({
    category: item.category,
    title: item.title,
    date: item.dateLabel,
    body: item.body,
    image: item.image || '',
    imageUrls: item.imageUrls || [],
    linkUrl: item.linkUrl || '',
    linkButtonText: item.linkButtonText || '',
    hasEmbeddedImage: item.hasEmbeddedImage === true
  });
}

function updateMenuCategoryFilter() {
  const select = document.getElementById('menuCategoryFilter');
  if (!select) return;

  const currentValue = select.value || '全て';
  const categories = [];
  const seen = {};

  getManagedCategoryNamesByTypes(['メニュー']).forEach(function (name) {
    if (!name || seen[name]) return;
    seen[name] = true;
    categories.push(name);
  });

  (USER_MENUS || []).forEach(function (item) {
    const name = String(item && item.category || '').trim();
    if (!name || seen[name]) return;
    seen[name] = true;
    categories.push(name);
  });

  select.innerHTML = '<option value="全て">カテゴリ: 全て</option>' + categories.map(function (name) {
    return `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`;
  }).join('');
  select.value = seen[currentValue] ? currentValue : '全て';
}

function buildNoticeFeedItems() {
  const minimumTimestamp = getNoticeItemTimestamp(NOTICE_FEED_START_DATE);

  // Blog: GAS側で新しい順にソート済みなので、indexが小さいほど新しい
  const blogFeed = (blogItems || []).map(function (item, index) {
    const publishMeta = buildNoticePublishMeta({
      noticeListedAt: item.noticeListedAt,
      updatedAt: item.updatedAt,
      fallbackDate: item.date,
      useFallbackDateForSort: true,
      useFallbackDateForLabel: true,
      useFallbackDateForVisibility: true
    });
    return {
      kind: 'blog',
      sourceKey: buildContentItemKey('blog', item),
      sourceWeight: (blogItems.length - index),
      sortOrder: Number(item.sortOrder || 0),
      timestamp: publishMeta.timestamp,
      visibilityTimestamp: publishMeta.visibilityTimestamp,
      dateLabel: publishMeta.dateLabel,
      category: resolveManagedNoticeCategory('blog', 'NEWS'),
      title: item.title || '',
      body: item.body || '',
      image: getDisplayImageUrl(item.image || item.imageUrl || ''),
      imageUrls: normalizeManagedImageList(item.imageUrls || item.image || item.imageUrl),
      icon: item.icon || '📢',
      linkUrl: item.linkUrl || '',
      linkButtonText: item.linkButtonText || '',
      hasEmbeddedImage: item.hasEmbeddedImage === true
    };
  }).filter(isNoticeFeedEntryVisible);

  // Calendar: スプレッドシート追加順
  const calendarFeed = (calendarData || []).filter(function (event) {
    return !isCalendarHolidayEvent(event) && !isCalendarVisitEvent(event);
  }).map(function (event, index) {
    const publishMeta = buildNoticePublishMeta({
      noticeListedAt: event.noticeListedAt,
      updatedAt: event.updatedAt,
      fallbackDate: event.date,
      useFallbackDateForVisibility: true
    });
    return {
      kind: 'calendar',
      sourceKey: buildContentItemKey('calendar', event),
      sourceWeight: index,
      sortOrder: Number(event.sortOrder || 0),
      timestamp: publishMeta.timestamp,
      visibilityTimestamp: publishMeta.visibilityTimestamp,
      dateLabel: publishMeta.dateLabel,
      category: resolveManagedNoticeCategory('calendar', 'カレンダー'),
      title: event.title || '',
      body: event.desc || '',
      image: getDisplayImageUrl(event.image || ''),
      imageUrls: normalizeManagedImageList(event.imageUrls || event.image),
      icon: '📅'
    };
  }).filter(isNoticeFeedEntryVisible);

  const productFeed = (PRODUCTS || []).map(function (product, index) {
    const publishMeta = buildNoticePublishMeta({
      noticeListedAt: product.noticeListedAt,
      updatedAt: product.updatedAt
    });
    return {
      kind: 'product',
      sourceKey: buildContentItemKey('product', product),
      sourceWeight: (PRODUCTS.length - index),
      sortOrder: Number(product.sortOrder || 0),
      timestamp: publishMeta.timestamp,
      visibilityTimestamp: publishMeta.visibilityTimestamp,
      dateLabel: publishMeta.dateLabel,
      category: resolveManagedNoticeCategory('product', 'ショップ'),
      title: product.name || '',
      body: product.description || '',
      image: getDisplayImageUrl(product.icon || ''),
      imageUrls: normalizeManagedImageList(product.imageUrls || product.icon),
      icon: '🛍️'
    };
  }).filter(isNoticeFeedEntryVisible);

  // Home: メニュー一覧はホームから閲覧する導線のため、カテゴリは「ホーム」でまとめる
  const menuFeed = (USER_MENUS || []).map(function (menu, index) {
    const publishMeta = buildNoticePublishMeta({
      noticeListedAt: menu.noticeListedAt,
      updatedAt: menu.updatedAt,
      fallbackDate: menu.date,
      useFallbackDateForSort: true,
      useFallbackDateForLabel: true,
      useFallbackDateForVisibility: true
    });
    return {
      kind: 'menu',
      sourceKey: buildContentItemKey('menu', menu),
      sourceWeight: index,
      sortOrder: Number(menu.sortOrder || 0),
      timestamp: publishMeta.timestamp,
      visibilityTimestamp: publishMeta.visibilityTimestamp,
      dateLabel: publishMeta.dateLabel,
      category: resolveManagedNoticeCategory('menu', 'ホーム'),
      title: menu.name || '',
      body: menu.description || (menu.reservationStatus ? '予約状況: ' + menu.reservationStatus : ''),
      image: getDisplayImageUrl(menu.imageUrl || ''),
      imageUrls: normalizeManagedImageList(menu.imageUrls || menu.imageUrl),
      icon: '🍴'
    };
  }).filter(isNoticeFeedEntryVisible);

  return blogFeed.concat(calendarFeed, productFeed, menuFeed).filter(function (item) {
    const visibilityTimestamp = Number(item.visibilityTimestamp || item.timestamp || 0);
    if (!visibilityTimestamp) return true;
    return visibilityTimestamp >= minimumTimestamp;
  }).sort(function (a, b) {
    const diff = (b.timestamp || 0) - (a.timestamp || 0);
    if (diff !== 0) return diff;

    const sortA = a.sortOrder || 0;
    const sortB = b.sortOrder || 0;
    if (sortA !== sortB) return sortB - sortA;
    return (b.sourceWeight || 0) - (a.sourceWeight || 0);
  });
}

/**
 * お知らせ一覧のカテゴリ選択肢を更新する
 */
function updateNoticeCategoryFilter(items) {
  const select = document.getElementById('noticeCategoryFilter');
  if (!select) return;

  const current = select.value;
  const seen = {};
  const categories = [];
  (items || []).forEach(function (item) {
    const name = String(item && item.category || '').trim();
    if (!name || seen[name]) return;
    seen[name] = true;
    categories.push(name);
  });
  if (!categories.length) {
    Object.values(getManagedNoticeCategoryMap()).forEach(function (name) {
      if (!name || seen[name]) return;
      seen[name] = true;
      categories.push(name);
    });
  }

  const options = ['<option value="">全て</option>'];
  categories.forEach(cat => {
    options.push(`<option value="${escapeHtml(cat)}">${escapeHtml(cat)}</option>`);
  });

  select.innerHTML = options.join('');
  select.value = seen[current] ? current : '';
}

function buildBlogCategoryTypeMap(categories) {
  const map = {};
  (categories || []).forEach(function (item) {
    const name = String(item && item.name || '').trim();
    const type = normalizeManagedCategoryType(item && item.type);
    if (!name) return;
    if (type === 'お知らせ' || type === 'ブログ') {
      map[name] = type;
    }
  });
  return map;
}

function normalizeBlogItems(items, categories) {
  const categoryTypeMap = buildBlogCategoryTypeMap(categories);
  return (items || []).map(function (item) {
    const category = String(item && item.category || 'お知らせ').trim() || 'お知らせ';
    const body = String(item && item.body || '');
    const embeddedImage = extractEmbeddedImageUrl(body);
    const inferredType = getInferredBlogType(category, categoryTypeMap[category], item && item.type);
    const imageUrls = normalizeManagedImageList(item && (item.imageUrls || item.images || [item.image, item.imageUrl]));
    const effectiveImageUrls = imageUrls.length ? imageUrls : (embeddedImage ? [embeddedImage] : []);
    return {
      date: normalizeBlogDate(item && item.date),
      updatedAt: String(item && item.updatedAt || item && item.date || ''),
      noticeListedAt: String(item && item.noticeListedAt || ''),
      title: String(item && item.title || ''),
      category: category,
      type: inferredType,
      sortOrder: Number(item && item.sortOrder || 0),
      icon: String(item && item.icon || (inferredType === 'お知らせ' ? '📢' : '📝')),
      body: body,
      image: effectiveImageUrls[0] || '',
      imageUrls: effectiveImageUrls,
      linkUrl: String(item && (item.linkUrl || item.link_url) || '').trim(),
      linkButtonText: String(item && (item.linkButtonText || item.link_button_text) || '').trim(),
      hasEmbeddedImage: !!embeddedImage,
      noticeStatus: normalizeNoticeVisibilityStatus(item && item.noticeStatus)
    };
  }).filter(function (item) {
    return item.title;
  }).sort(function (a, b) {
    const timeA = parseLooseDateToTimestamp(a.updatedAt || a.date);
    const timeB = parseLooseDateToTimestamp(b.updatedAt || b.date);
    return timeB - timeA;
  });
}

function normalizeCalendarEventEntry(item) {
  const event = item && typeof item === 'object' ? item : {};
  const imageUrls = normalizeManagedImageList(event.imageUrls || event.images || event.image);
  return {
    rowIdx: event.rowIdx,
    date: String(event.date || ''),
    title: String(event.title || ''),
    desc: String(event.desc || ''),
    color: String(event.color || '#e57373'),
    category: String(event.category || ''),
    image: imageUrls[0] || getDisplayImageUrl(event.image) || '',
    imageUrls: imageUrls,
    updatedAt: String(event.updatedAt || ''),
    noticeListedAt: String(event.noticeListedAt || ''),
    publishAt: String(event.publishAt || ''),
    noticeStatus: normalizeNoticeVisibilityStatus(event.noticeStatus),
    sortOrder: Number(event.sortOrder || 0)
  };
}

function normalizeBlogDate(rawValue) {
  const raw = String(rawValue || '').trim();
  const match = raw.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})(?:\D+(\d{1,2}):(\d{1,2}))?$/);
  if (!match) return raw;
  return match[1] + '/' + String(match[2]).padStart(2, '0') + '/' + String(match[3]).padStart(2, '0');
}

function getInferredBlogType(category, categoryType, rawType) {
  if (rawType === 'お知らせ' || rawType === 'ブログ') return rawType;
  if (categoryType === 'お知らせ' || categoryType === 'ブログ') return categoryType;
  return category === 'お知らせ' || category === '休診情報' ? 'お知らせ' : 'ブログ';
}

function getBlogItemType(item) {
  return getInferredBlogType(item && item.category, item && item.type, item && item.type);
}

function extractEmbeddedImageUrl(body) {
  const match = String(body || '').match(/📷\s*(https?:\/\/\S+)/);
  return match ? match[1] : '';
}

function getDisplayImageUrl(rawValue) {
  const value = String(rawValue || '').trim();
  if (/^https?:\/\//i.test(value) || /^data:image\//i.test(value)) return value;
  return '';
}

function normalizeManagedImageList(rawValue) {
  if (Array.isArray(rawValue)) {
    return rawValue.reduce(function (acc, item) {
      return acc.concat(normalizeManagedImageList(item));
    }, []);
  }
  const raw = String(rawValue || '').trim();
  if (!raw) return [];
  if (raw.charAt(0) === '[') {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) return normalizeManagedImageList(parsed);
    } catch (error) {
      console.warn('[images] failed to parse image list', error);
    }
  }
  const parts = raw.split(/\r?\n+/).map(function (item) {
    return String(item || '').trim();
  }).filter(Boolean);
  const candidates = parts.length > 1 ? parts : [raw];
  return candidates.map(getDisplayImageUrl).filter(Boolean).filter(function (url, index, array) {
    return array.indexOf(url) === index;
  });
}

function extractDriveFileId(rawValue) {
  const value = String(rawValue || '').trim();
  if (!value) return '';
  const idPatterns = [
    /[?&]id=([a-zA-Z0-9_-]+)/,
    /\/d\/([a-zA-Z0-9_-]+)/,
    /\/thumbnail\?id=([a-zA-Z0-9_-]+)/,
    /\/uc\?(?:[^#]*&)?id=([a-zA-Z0-9_-]+)/
  ];
  for (let i = 0; i < idPatterns.length; i++) {
    const matched = value.match(idPatterns[i]);
    if (matched && matched[1]) return matched[1];
  }
  return '';
}

function getContentDisplayImageUrl(rawValue) {
  const value = getDisplayImageUrl(rawValue);
  if (!value) return '';
  const driveFileId = extractDriveFileId(value);
  if (!driveFileId) return value;
  return 'https://lh3.googleusercontent.com/d/' + driveFileId + '=w1600';
}

function buildDetailImageGalleryHtml(images, altBase, extraStyle) {
  const imageUrls = normalizeManagedImageList(images).map(function (url) {
    return getContentDisplayImageUrl(url);
  }).filter(Boolean);
  if (!imageUrls.length) return '';
  const style = extraStyle ? ` style="${extraStyle}"` : '';
  return `<div class="detail-image-gallery"${style}>${imageUrls.map(function (url, index) {
    return `<div class="blog-media-frame blog-media-frame--detail"><img src="${url}" class="blog-media-image blog-media-image--detail" alt="${escapeHtml((altBase || 'image') + ' ' + (index + 1))}"></div>`;
  }).join('')}</div>`;
}

function parseLooseDateToTimestamp(rawValue) {
  if (typeof rawValue === 'number' && Number.isFinite(rawValue)) return rawValue;
  const raw = String(rawValue || '').trim();
  if (!raw) return 0;
  const matched = raw.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})(?:\D+(\d{1,2}):(\d{1,2}))?$/);
  if (matched) {
    return new Date(
      Number(matched[1]),
      Number(matched[2]) - 1,
      Number(matched[3]),
      Number(matched[4] || 0),
      Number(matched[5] || 0),
      0,
      0
    ).getTime();
  }
  const parsed = Date.parse(raw);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatCustomerDateYmd(rawValue) {
  if (rawValue instanceof Date) {
    if (Number.isNaN(rawValue.getTime())) return '';
    return rawValue.getFullYear() + '/' +
      String(rawValue.getMonth() + 1).padStart(2, '0') + '/' +
      String(rawValue.getDate()).padStart(2, '0');
  }
  const raw = String(rawValue || '').trim();
  if (!raw) return '';
  const timestamp = parseLooseDateToTimestamp(rawValue);
  if (Number.isFinite(timestamp) && timestamp !== 0) {
    const date = new Date(timestamp);
    if (!Number.isNaN(date.getTime())) {
      return date.getFullYear() + '/' +
        String(date.getMonth() + 1).padStart(2, '0') + '/' +
        String(date.getDate()).padStart(2, '0');
    }
  }
  const dateOnly = raw.split(/[ T]/)[0];
  const match = dateOnly.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
  if (match) {
    return match[1] + '/' + String(match[2]).padStart(2, '0') + '/' + String(match[3]).padStart(2, '0');
  }
  return dateOnly.replace(/[.-]/g, '/');
}

function formatCustomerDateYmdHm(rawValue) {
  const timestamp = parseLooseDateToTimestamp(rawValue);
  if (!Number.isFinite(timestamp) || timestamp === 0) return '';
  const date = new Date(timestamp);
  if (Number.isNaN(date.getTime())) return '';
  return date.getFullYear() + '/' +
    String(date.getMonth() + 1).padStart(2, '0') + '/' +
    String(date.getDate()).padStart(2, '0') + ' ' +
    String(date.getHours()).padStart(2, '0') + ':' +
    String(date.getMinutes()).padStart(2, '0');
}

function getNoticeItemTimestamp(rawValue) {
  return parseLooseDateToTimestamp(rawValue);
}

function buildNoticePublishMeta(options) {
  const noticeListedAt = String(options && options.noticeListedAt || '').trim();
  const updatedAt = String(options && options.updatedAt || '').trim();
  const fallbackDate = String(options && options.fallbackDate || '').trim();
  const sortRaw = noticeListedAt || updatedAt || (options && options.useFallbackDateForSort ? fallbackDate : '');
  const labelRaw = noticeListedAt || updatedAt || (options && options.useFallbackDateForLabel ? fallbackDate : '');
  const visibilityRaw = noticeListedAt || updatedAt || (options && options.useFallbackDateForVisibility ? fallbackDate : '');
  return {
    timestamp: getNoticeItemTimestamp(sortRaw),
    visibilityTimestamp: getNoticeItemTimestamp(visibilityRaw),
    dateLabel: formatNoticeDateLabel(labelRaw)
  };
}

function hasExplicitTimeValue(rawValue) {
  if (typeof rawValue === 'number' && Number.isFinite(rawValue)) return true;
  return /\d{1,2}:\d{2}/.test(String(rawValue || '').trim());
}

function formatNoticeDateLabel(rawValue) {
  return formatCustomerDateYmd(rawValue);
}

function formatPushNoticeDate(rawValue) {
  return formatCustomerDateYmd(rawValue);
}

function buildUnreadBadgeHtml(kind, item) {
  return isContentItemUnread(kind, item) ? '<span class="item-new-badge">NEW</span>' : '';
}

function buildFavoriteActionMarkup(kind, item) {
  const entry = buildFavoriteEntry(kind, item);
  if (!entry) return '';
  return '<button class="btn secondary favorite-action-btn" type="button" onclick="toggleFavoriteByKey(\'' + encodeURIComponent(entry.key) + '\')">' +
    (isFavoriteKey(entry.key) ? '★ お気に入りから外す' : '☆ お気に入りに追加') +
    '</button>';
}

function buildFavoriteEntry(kind, item) {
  if (!item) return null;
  if (kind === 'product') {
    return {
      key: buildContentItemKey('product', item),
      kind: 'product',
      page: 'shop',
      title: item.name || '商品',
      subtitle: item.category || 'ショップ',
      body: item.description || '',
      image: item.icon || '',
      dateLabel: formatCustomerDateYmd(item.updatedAt || item.date || ''),
      contentIndex: PRODUCTS.indexOf(item)
    };
  }
  if (kind === 'blog') {
    return {
      key: buildContentItemKey('blog', item),
      kind: 'blog',
      page: 'blog',
      title: item.title || 'NEWS',
      subtitle: item.category || 'NEWS',
      body: item.body || '',
      image: item.image || item.imageUrl || '',
      dateLabel: formatCustomerDateYmd(item.updatedAt || item.date || ''),
      contentIndex: blogItems.indexOf(item)
    };
  }
  if (kind === 'menu') {
    return {
      key: buildContentItemKey('menu', item),
      kind: 'menu',
      page: 'menu-list',
      title: item.name || 'メニュー',
      subtitle: item.category || 'ホーム',
      body: item.description || '',
      image: item.imageUrl || '',
      dateLabel: formatCustomerDateYmd(item.updatedAt || item.date || ''),
      contentIndex: USER_MENUS.indexOf(item)
    };
  }
  if (kind === 'calendar') {
    return {
      key: buildContentItemKey('calendar', item),
      kind: 'calendar',
      page: 'calendar',
      title: item.title || 'イベント',
      subtitle: 'カレンダー',
      body: item.desc || '',
      image: item.image || '',
      dateLabel: formatCustomerDateYmd(item.updatedAt || item.date || ''),
      contentIndex: calendarData.indexOf(item)
    };
  }
  return null;
}

function toggleFavoriteByKey(key) {
  const resolvedKey = decodeURIComponent(String(key || ''));
  const existing = getFavoriteEntries().find(function (entry) {
    return entry && entry.key === resolvedKey;
  });
  if (existing) {
    saveFavoriteEntries(getFavoriteEntries().filter(function (entry) {
      return entry.key !== resolvedKey;
    }));
    showToast('お気に入りから外しました');
  } else {
    const entry = (PRODUCTS || []).map(function (item) { return buildFavoriteEntry('product', item); })
      .concat((blogItems || []).map(function (item) { return buildFavoriteEntry('blog', item); }))
      .concat((USER_MENUS || []).map(function (item) { return buildFavoriteEntry('menu', item); }))
      .concat((calendarData || []).map(function (item) { return buildFavoriteEntry('calendar', item); }))
      .find(function (item) { return item && item.key === resolvedKey; });
    if (!entry) return;
    toggleFavoriteEntry(entry);
    showToast('お気に入りに追加しました');
  }
  renderFavoriteList();
  refreshActiveDetailViews();
}

function refreshActiveDetailViews() {
  if (document.getElementById('blogDetailModal').classList.contains('open')) {
    const currentTitle = document.querySelector('#blogDetailContent .blog-detail-title');
    if (currentTitle) {
      const blog = blogItems.find(function (item) {
        return item.title === currentTitle.textContent;
      });
      if (blog) openBlogDetail(blog);
    }
  }
  if (document.getElementById('menuDetailModal').classList.contains('open')) {
    const titleNode = document.querySelector('#menuDetailContent .blog-detail-title');
    if (titleNode) {
      const menu = USER_MENUS.find(function (item) {
        return item.name === titleNode.textContent;
      });
      if (menu) openMenuDetail(USER_MENUS.indexOf(menu));
    }
  }
  if (document.getElementById('calendarEventDetailModal').classList.contains('open')) {
    const titleNode = document.querySelector('#calendarEventDetailContent .blog-detail-title');
    if (titleNode) {
      const event = calendarData.find(function (item) {
        return item.title === titleNode.textContent;
      });
      if (event) openCalendarEventDetail(calendarData.indexOf(event));
    }
  }
  if (document.getElementById('productModal').classList.contains('open') && currentProdIdx != null) {
    openProductModal(currentProdIdx);
  }
}
function openBlogDetail(item) {
  markContentItemSeen(item && item.kind ? item.kind : 'blog', item);
  const detailImageHtml = !item.hasEmbeddedImage
    ? buildDetailImageGalleryHtml(item.imageUrls || item.image, item.title || 'Blog Image')
    : '';

  const formattedBody = renderManagedTextHtml(item.body || '', {
    inlineImageClass: 'blog-inline-image blog-inline-image--detail',
    inlineImageAlt: (item.title || '記事画像')
  }) || '本文はありません。';
  const displayDate = formatCustomerDateYmd(item.date);
  
  // URLの正規化（http/httpsがなければ補完）
  const ensureAbsoluteUrl = (url) => {
    if (!url) return '';
    const trimmed = String(url).trim();
    if (!trimmed) return '';
    if (trimmed.startsWith('http://') || trimmed.startsWith('https://')) return trimmed;
    // もし google.com のような形式なら https:// を付与
    return 'https://' + trimmed;
  };
  const absoluteUrl = ensureAbsoluteUrl(item.linkUrl);

  const linkText = item.linkButtonText || '詳しく見る';
  const linkButtonHtml = absoluteUrl ? `
    <div style="margin-top:28px; text-align:center;">
      <button class="btn" style="width:100%; max-width:320px; padding:16px; background:var(--primary); color:white; border:none; border-radius:14px; font-weight:bold; font-size:16px; cursor:pointer; box-shadow:0 6px 16px rgba(0,0,0,0.12);" onclick="window.open('${absoluteUrl}', '_blank', 'noopener')">
        ${escapeHtml(linkText)} 🔗
      </button>
      <div style="margin-top:10px; font-size:11px; color:var(--text-light); word-break:break-all; line-height:1.4;">
        リンク先: ${escapeHtml(absoluteUrl)}<br>
        <span style="font-size:9px;">※タップすると外部ブラウザで開きます</span>
      </div>
    </div>` : '';

  document.getElementById('blogDetailContent').innerHTML = `
    <span class="blog-detail-cat">${escapeHtml(item.category || '')}</span>
    <div class="blog-detail-title">${escapeHtml(item.title || '')}</div>
    <div class="blog-detail-date">${escapeHtml(displayDate)}</div>
    ${detailImageHtml}
    <div class="blog-detail-body managed-rich-text">${formattedBody}</div>
    ${linkButtonHtml}
    <div style="margin-top:20px;">${buildFavoriteActionMarkup('blog', item)}</div>
    <div style="height:80px;"></div>`;
  if (document.getElementById('page-blog').classList.contains('active')) {
    renderDividedBlogList();
  } else {
    renderBlogList('homeNewsList', 3);
  }
  if (document.getElementById('page-notices').classList.contains('active')) {
    renderPushNotices();
  }
  updateNavBadges();
  openModal('blogDetailModal');
}

// ===== ショップ =====
function openProductModal(idx) {
  currentProdIdx = idx; modalQty = 1;
  const p = PRODUCTS[idx];
  markContentItemSeen('product', p);
  const img = document.getElementById('prodModalImg');
  img.className = 'prod-img';
  const productGalleryHtml = buildDetailImageGalleryHtml(p.imageUrls || p.icon, p.name || 'Product Image');

  // 1. iconがURL(http/data)ならそれを最優先、次にimgKeyMap、最後に絵文字
  if (productGalleryHtml) {
    img.style.background = '#f8f5f0';
    img.innerHTML = productGalleryHtml;
  } else if (p.icon && (p.icon.startsWith('http') || p.icon.startsWith('data:'))) {
    img.style.background = '#f8f5f0';
    img.innerHTML = '<img src="' + getContentDisplayImageUrl(p.icon) + '" alt="" class="prod-img-media">';
  } else if (p.imgKey && PROD_IMAGES[p.imgKey]) {
    img.style.background = '#f8f5f0';
    img.innerHTML = '<img src="' + PROD_IMAGES[p.imgKey] + '" alt="" class="prod-img-media">';
  } else {
    img.style.background = '';
    img.textContent = p.icon || '🌿';
  }
  document.getElementById('prodModalName').textContent = p.name;
  document.getElementById('prodModalPrice').innerHTML = buildProductPriceMarkup(p, 1, { mode: 'unit', includeTax: true, showPeriod: true });

  let descHtml = renderManagedTextHtml(p.description || '', {
    inlineImageClass: 'blog-inline-image blog-inline-image--detail',
    inlineImageAlt: (p.name || '商品画像')
  });
  const descriptionGalleryHtml = buildDetailImageGalleryHtml(p.descriptionImageUrls || p.descriptionImage, (p.name || 'Product') + ' Description Image', 'margin-top:12px;');
  if (descriptionGalleryHtml) {
    descHtml += descriptionGalleryHtml;
  }
  const prodModalDesc = document.getElementById('prodModalDesc');
  prodModalDesc.classList.add('managed-rich-text');
  prodModalDesc.innerHTML = descHtml || '商品説明はありません';
  prodModalDesc.style.display = 'none'; // 初期非表示
  const favoriteWrap = document.getElementById('prodModalFavoriteWrap');
  if (favoriteWrap) favoriteWrap.innerHTML = buildFavoriteActionMarkup('product', p);
  document.getElementById('modalQtyDisp').textContent = 1;
  updateNavBadges();
  openModal('productModal');
}

function toggleModalDesc() {
  const desc = document.getElementById('prodModalDesc');
  const isHidden = window.getComputedStyle(desc).display === 'none';
  desc.style.display = isHidden ? 'block' : 'none';
  event.target.textContent = isHidden ? '× 閉じる' : '商品説明をみる';
}
function changeQty(d) {
  modalQty = Math.max(1, modalQty + d);
  document.getElementById('modalQtyDisp').textContent = modalQty;
}
function addToCart() {
  const ex = cart.find(c => c.idx === currentProdIdx);
  if (ex) ex.qty += modalQty; else cart.push({ idx: currentProdIdx, qty: modalQty });
  closeModal('productModal');
  updateCartUI();
  showToast('カートに追加しました 🛒');
}

// ===== カート =====
function updateCartUI() {
  const count = cart.reduce((s, c) => s + c.qty, 0);
  const badge = document.getElementById('cartBadge');
  badge.style.display = count > 0 ? 'flex' : 'none';
  badge.textContent = count;
  if (document.getElementById('page-cart').classList.contains('active')) renderCart();
}
function renderCart() {
  const empty = document.getElementById('cartEmpty');
  const content = document.getElementById('cartContent');
  if (!cart.length) { empty.style.display = 'block'; content.style.display = 'none'; return; }
  empty.style.display = 'none'; content.style.display = 'block';
  const list = document.getElementById('cartList'); list.innerHTML = '';
  let total = 0;
  cart.forEach((c, ci) => {
    const p = PRODUCTS[c.idx];
    const el = document.createElement('div'); el.className = 'cart-row';

    // 表示用画像の決定
    let imgSrcHtml = p.icon;
    if (p.icon && (p.icon.startsWith('http') || p.icon.startsWith('data:'))) {
      imgSrcHtml = `<img src="${getContentDisplayImageUrl(p.icon)}" class="cart-thumb-image" alt="">`;
    } else if (p.imgKey && PROD_IMAGES[p.imgKey]) {
      imgSrcHtml = `<img src="${PROD_IMAGES[p.imgKey]}" class="cart-thumb-image" alt="">`;
    }

    const pricing = getProductPricing(p, c.qty);
    let subtotal = pricing.total;
    let priceNote = buildProductPriceMarkup(p, c.qty, { mode: 'total', includeTax: false, showPeriod: false });
    total += subtotal;

    el.innerHTML = `<div class="cart-thumb ${p.bg}">${imgSrcHtml}</div><div class="cart-info"><div class="cart-item-name">${p.name}</div><div class="cart-item-price">${priceNote}</div><div class="cart-qty-row"><button class="qty-ctrl-btn" onclick="chgCartQty(${ci},-1)">－</button><span class="qty-ctrl-num">${c.qty}</span><button class="qty-ctrl-btn" onclick="chgCartQty(${ci},1)">＋</button></div></div><div class="cart-del" onclick="rmCartItem(${ci})">🗑</div>`;
    list.appendChild(el);
  });
  document.getElementById('cartSubtotal').textContent = '¥' + total.toLocaleString();
  document.getElementById('cartTotal').textContent = '¥' + total.toLocaleString();
  updateCheckoutBtn();
}
function chgCartQty(ci, d) {
  if (isOrderSubmitting) return;
  cart[ci].qty = Math.max(1, cart[ci].qty + d);
  updateCartUI();
  renderCart();
}
function rmCartItem(ci) {
  if (isOrderSubmitting) return;
  cart.splice(ci, 1);
  updateCartUI();
  renderCart();
}
function updateCheckoutBtn() {
  const btn = document.getElementById('checkoutBtn');
  if (!btn) return;
  btn.disabled = !cart.length || isOrderSubmitting;
  btn.textContent = isOrderSubmitting ? '送信中...' : 'ご注文を確定する';
}
function proceedCheckout() {
  if (isOrderSubmitting) return;
  finalizeOrder('現金払い');
}
async function finalizeOrder(payLabel) {
  if (isOrderSubmitting || !cart.length) return;
  if (!_profile || !_profile.memberId) {
    showToast('先にプロフィールを登録してください');
    switchPage('mypage');
    return;
  }

  const cartSnapshot = cart.map(c => ({ idx: c.idx, qty: c.qty }));
  const total = cartSnapshot.reduce((sum, c) => {
    const prod = PRODUCTS[c.idx];
    return sum + getProductPricing(prod, c.qty).total;
  }, 0);
  // GAS側の getCurrentTime() と同じ形式: M/d HH:mm
  const now = new Date();
  const ts = formatCustomerDateYmd(now);
  const orderId = 'ORD-' + Date.now();
  const order = { id: orderId, items: cartSnapshot, total, payment: payLabel, status: 'pending', time: ts, memberId: _profile.memberId };

  isOrderSubmitting = true;
  updateCheckoutBtn();

  try {
    // GASの handleOrder が期待する payload 構造:
    // { type:'order', customerName, items:[{name,qty,price}], total, payment, orderId, time, memberId }
    const activityProfile = getUserActivityProfilePayload();
    const res = await postToGAS({
      type: 'order',
      customerName: CUSTOMER_NAME || activityProfile.name,
      name: activityProfile.name,
      memberId: _profile.memberId,
      phone: activityProfile.phone,
      birthday: activityProfile.birthday,
      address: activityProfile.address,
      items: order.items.map(c => {
        const prod = PRODUCTS[c.idx];
        const effectivePrice = getProductPricing(prod, c.qty).unitPrice;
        return {
          name: prod.name,
          qty: c.qty,
          price: effectivePrice
        };
      }),
      total: total,
      payment: payLabel,
      orderId: orderId,
      time: ts
    });

    if (!res || (res.status !== 'ok' && res.status !== 'queued')) {
      showToast('注文の送信に失敗しました');
      return;
    }

    if (res.status === 'queued') {
      order.status = 'sync_pending';
    }
    orders.unshift(order);
    cart = [];
    updateCartUI();
    document.getElementById('orderCompleteMsg').innerHTML = res.status === 'queued'
      ? `お支払い：${payLabel}<br><br>通信回復後に注文を送信します。注文内容はこの端末に保持されています。`
      : `お支払い：${payLabel}<br><br>ご来院時にスタッフにお声がけください。`;
    openModal('orderCompleteModal');
  } catch (e) {
    showToast('注文の送信に失敗しました');
  } finally {
    isOrderSubmitting = false;
    updateCheckoutBtn();
  }
}
function finishOrder() { closeModal('orderCompleteModal'); renderOrderHistory(); switchPage('mypage'); }

function setAppDialogButtonVariant(button, variant) {
  if (!button) return;
  button.className = 'btn ' + (variant === 'danger' ? 'danger' : 'primary');
}

function showAppDialog(options) {
  const modal = document.getElementById('appDialogModal');
  const title = document.getElementById('appDialogTitle');
  const message = document.getElementById('appDialogMessage');
  const confirmBtn = document.getElementById('appDialogConfirmBtn');
  const cancelBtn = document.getElementById('appDialogCancelBtn');
  if (!modal || !title || !message || !confirmBtn || !cancelBtn) return Promise.resolve(false);
  if (activeAppDialogResolver) return Promise.resolve(false);

  const settings = options || {};
  const kind = settings.kind === 'confirm' ? 'confirm' : 'alert';
  title.textContent = settings.title || 'ご案内';
  message.textContent = settings.message || '';
  confirmBtn.textContent = settings.confirmLabel || (kind === 'confirm' ? 'OK' : '閉じる');
  setAppDialogButtonVariant(confirmBtn, settings.confirmVariant || 'primary');

  if (kind === 'confirm') {
    cancelBtn.style.display = 'block';
    cancelBtn.textContent = settings.cancelLabel || 'キャンセル';
  } else {
    cancelBtn.style.display = 'none';
    cancelBtn.textContent = 'キャンセル';
  }

  openModal('appDialogModal');
  return new Promise(function (resolve) {
    activeAppDialogResolver = resolve;
  });
}

function resolveAppDialog(result) {
  const resolver = activeAppDialogResolver;
  activeAppDialogResolver = null;
  closeModal('appDialogModal');
  if (resolver) resolver(result);
}

function showAppAlert(message, options) {
  const settings = options || {};
  return showAppDialog({
    kind: 'alert',
    title: settings.title || 'ご案内',
    message: message,
    confirmLabel: settings.confirmLabel || '閉じる',
    confirmVariant: settings.confirmVariant || 'primary'
  });
}

function showAppConfirm(message, options) {
  const settings = options || {};
  return showAppDialog({
    kind: 'confirm',
    title: settings.title || '確認',
    message: message,
    confirmLabel: settings.confirmLabel || 'OK',
    cancelLabel: settings.cancelLabel || 'キャンセル',
    confirmVariant: settings.confirmVariant || 'primary'
  });
}

// ===== 注文履歴 =====
function shouldHideOrderFromHistory(order) {
  if (!order) return true;
  const status = String(order.status || '').trim().toLowerCase();
  return status === 'cancelled' ||
    status === 'キャンセル済' ||
    status === 'キャンセル' ||
    status.indexOf('キャンセル') !== -1;
}

function renderOrderHistoryError(message) {
  const list = document.getElementById('orderHistoryList');
  if (!list) return;
  list.innerHTML = `
        <div style="text-align:center;padding:26px 12px;color:var(--text-light);line-height:1.7">
          <div style="font-size:13px;margin-bottom:12px">${message}</div>
          <button class="btn secondary" id="orderHistoryRetryBtn" onclick="renderOrderHistory()" style="padding:8px 14px">再読み込み</button>
        </div>
      `;
}

async function renderOrderHistory() {
  const list = document.getElementById('orderHistoryList');
  if (!_profile) { list.innerHTML = '<div style="text-align:center;font-size:13px;color:var(--text-light);padding:26px 0">プロフィールを登録すると履歴が表示されます</div>'; return; }

  const previousOrders = Array.isArray(orders) ? orders.slice() : [];
  list.innerHTML = '<div style="text-align:center;padding:26px 0"><span class="loading-spinner"></span> 読み込み中...</div>';

  // GASから最新の履歴を取得（同期）
  const res = await getFromGAS('getCustomerOrders', { memberId: _profile.memberId });
  if (res && res.status === 'ok') {
    orders = (res.orders || []).filter(function (order) {
      return !shouldHideOrderFromHistory(order);
    });
    renderOrderHistoryUI();
    return;
  }

  orders = previousOrders;
  const hasCachedOrders = orders.some(function (order) {
    return !shouldHideOrderFromHistory(order);
  });
  if (hasCachedOrders) {
    showToast('注文履歴の更新に失敗しました');
    renderOrderHistoryUI();
    return;
  }

  renderOrderHistoryError('注文履歴の取得に失敗しました。通信状況をご確認ください。');
}

function renderOrderHistoryUI() {
  const list = document.getElementById('orderHistoryList');
  if (!_profile) {
    list.innerHTML = '<div style="text-align:center;font-size:13px;color:var(--text-light);padding:26px 0">プロフィールを登録すると履歴が表示されます</div>';
    return;
  }
  const visibleOrders = orders.filter(function (order) {
    return !shouldHideOrderFromHistory(order);
  });

  if (!visibleOrders.length) {
    list.innerHTML = '<div style="text-align:center;font-size:13px;color:var(--text-light);padding:26px 0">注文履歴はありません</div>';
    return;
  }
  list.innerHTML = '';
  visibleOrders.forEach(o => {
    const isCancelled = o.status === 'cancelled' || o.status === 'キャンセル済' || o.status === 'キャンセル';
    const isActionLocked = isCancelSubmitting || isReceiptSubmitting;
    const receiptButtonLabel = isReceiptSubmitting && receiptSubmittingOrderId === o.id ? '更新中...' : '受け取りました';
    const isSyncPending = o.status === 'sync_pending';
    const sl = isSyncPending ? '送信待ち' : (o.status === 'pending' ? '受付中' : (isCancelled ? 'キャンセル済' : '完了'));
    const sc = isSyncPending ? 's-pending' : (o.status === 'pending' ? 's-pending' : (isCancelled ? 's-cancelled' : 's-done'));

    const names = o.items.map(c => {
      const itemName = c.name || (productItems.find(p => p.id === c.idx) ? productItems.find(p => p.id === c.idx).name : '不明な商品');
      return itemName + ' ×' + c.qty;
    }).join('、');

    const d = document.createElement('div');
    d.className = 'order-history-item';
    d.innerHTML = `
          <div style="display:flex;justify-content:space-between;align-items:flex-start">
            <div>
              <span class="status-chip ${sc}">${sl}</span>
              <span style="font-size:10px;color:var(--text-light);margin-left:8px">${formatCustomerDateYmd(o.time)}</span>
            </div>
          </div>
          <div style="font-size:12px;color:var(--text-dark);margin:8px 0;line-height:1.5">${names}</div>
          <div style="display:flex;justify-content:space-between;align-items:center;margin-top:5px">
            <span style="font-size:11px;color:var(--text-light)">${o.payment}</span>
            <span style="font-size:14px;font-weight:700;color:var(--brown)">¥${Number(o.total).toLocaleString()}</span>
          </div>
          <div style="margin-top:8px; display:flex; gap:8px;">
            ${o.status === 'pending' ? `<button class="btn danger" style="flex:1;padding:8px" onclick="openCancelModal('${o.id}')" ${isActionLocked ? 'disabled' : ''}>キャンセルする</button>` : ''}
            ${(!o.checked && !isCancelled && !isSyncPending) ? `<button class="btn primary" style="flex:1;padding:8px;background:var(--sage)" onclick="confirmReceipt('${o.id}')" ${isActionLocked ? 'disabled' : ''}>${receiptButtonLabel}</button>` : ''}
            ${(o.checked) ? `<div style="flex:1;padding:8px;background:#f0f0f0;color:#888;text-align:center;border-radius:8px;font-size:12px;font-weight:bold;">受取完了</div>` : ''}
          </div>
        `;
    list.appendChild(d);
  });
}

async function confirmReceipt(orderId) {
  if (isReceiptSubmitting) return;
  const targetOrder = orders.find(function (order) { return order.id === orderId; });
  if (!targetOrder || shouldHideOrderFromHistory(targetOrder)) return;
  const confirmed = await showAppConfirm('商品を受け取りましたか？\nステータスを受取済に更新します。', {
    title: '受取報告',
    confirmLabel: '受取済にする',
    cancelLabel: '戻る',
    confirmVariant: 'primary'
  });
  if (!confirmed) return;

  isReceiptSubmitting = true;
  receiptSubmittingOrderId = orderId;
  renderOrderHistoryUI();
  showToast('更新中...');
  try {
    const res = await postToGAS({ type: 'confirmReceipt', orderId: orderId });
    if (res && res.status === 'ok') {
      showToast('受取を報告しました🌿');
      orders = orders.filter(o => o.id !== orderId);
    } else {
      showToast('更新に失敗しました');
    }
  } finally {
    isReceiptSubmitting = false;
    receiptSubmittingOrderId = null;
    renderOrderHistoryUI();
  }
}

function updateCancelModalState() {
  const confirmBtn = document.getElementById('cancelConfirmBtn');
  const backBtn = document.getElementById('cancelBackBtn');
  const isBusy = isCancelSubmitting || isReceiptSubmitting;
  if (confirmBtn) {
    confirmBtn.disabled = isBusy;
    confirmBtn.textContent = isCancelSubmitting ? 'キャンセル中...' : 'キャンセルする';
  }
  if (backBtn) backBtn.disabled = isCancelSubmitting;
}
function openCancelModal(id) {
  if (isCancelSubmitting || isReceiptSubmitting) return;
  cancelOrderId = id;
  updateCancelModalState();
  openModal('cancelModal');
}
async function confirmCancel() {
  if (isCancelSubmitting || isReceiptSubmitting || !cancelOrderId) return;

  const targetOrderId = cancelOrderId;
  const targetOrder = orders.find(o => o.id === targetOrderId);
  if (!targetOrder) {
    cancelOrderId = null;
    closeModal('cancelModal');
    return;
  }

  isCancelSubmitting = true;
  updateCancelModalState();
  showToast('キャンセル処理中...');

  try {
    const res = await postToGAS({ type: 'cancel', orderId: targetOrder.id });
    if (res && res.status === 'ok') {
      orders = orders.filter(o => o.id !== targetOrderId);
      showToast('キャンセルしました');
      await renderOrderHistory();
    } else {
      showToast('キャンセルの送信に失敗しました');
    }
  } finally {
    isCancelSubmitting = false;
    cancelOrderId = null;
    updateCancelModalState();
    closeModal('cancelModal');
    if (!document.getElementById('cancelModal').classList.contains('open')) {
      renderOrderHistoryUI();
    }
  }
}

// ===== モーダル / ページ / トースト =====
function openModal(id) { document.getElementById(id).classList.add('open'); }
function closeModal(id, e) { if (e && e.target !== document.getElementById(id)) return; document.getElementById(id).classList.remove('open'); }
function switchPage(name) {
  if (name === 'blog') {
    const hash = blogItems.map(function (i) {
      return [i.updatedAt || '', i.date || '', i.title || '', i.body || ''].join('_');
    }).join('|');
    localStorage.setItem('last_seen_blog_hash', hash);
    document.getElementById('badge-blog').style.display = 'none';
  }
  if (name === 'calendar') {
    const hash = calendarData.map(function (i) {
      return [i.updatedAt || '', i.date || '', i.title || '', i.desc || ''].join('_');
    }).join('|');
    localStorage.setItem('last_seen_calendar_hash', hash);
    document.getElementById('badge-calendar').style.display = 'none';
  }
  if (name === 'notices') {
    // 拡声器ページを開いた山時点で統合ハッシュを保存 → バッジを溈ませる
    localStorage.setItem('last_seen_notices_hash', computeNoticeListHash());
    document.getElementById('badge-notices').style.display = 'none';
    if (isDataLoaded) {
      const noticeList = document.getElementById('noticeList');
      if (noticeList) {
        noticeList.innerHTML = '<div style="text-align:center; padding:40px 20px; color:var(--text-mid);">最新のお知らせ一覧を確認中...</div>';
      }
      refreshNoticeFeed().catch(function (e) {
        console.error('notice feed refresh error:', e);
      }).finally(function () {
        renderPushNotices();
      });
    } else {
      renderPushNotices();
    }
  }
  if (name === 'home') {
    updateNavBadges();
  }

  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('page-' + name).classList.add('active');
  const nav = document.getElementById('nav-' + name);
  if (nav) nav.classList.add('active');
  window.scrollTo({ top: 0, behavior: 'smooth' });
  if (name === 'cart') renderCart();
  if (name === 'mypage') {
    renderOrderHistory();
    renderSurveys();
    renderFavoriteList();
    renderUserDevices();
    renderRetryQueueStatus();
    renderAccessibilitySettings();
    renderCurrentDeviceGuidance();
    loadUserDevices().catch(function (error) {
      console.error('loadUserDevices switchPage error:', error);
    });
    if (typeof syncNativePushStatus === 'function') syncNativePushStatus();
  }
  if (name === 'survey') {
    renderSurveys();
  }
  if (name === 'blog') renderDividedBlogList();
  if (name === 'shop') {
    localStorage.setItem('last_seen_product_hash', computeProductListHash());
    var sb = document.getElementById('badge-shop');
    if (sb) sb.style.display = 'none';
  }
  if (name === 'calendar') {
    renderCalendar();
  }
  if (name === 'menu-list') {
    loadMenus();
  }
}

const PUSH_ENABLED_STORAGE_KEY = 'push_enabled';
const NATIVE_PUSH_PLAYER_ID_STORAGE_KEY = 'mayumi_native_push_player_id';
const NATIVE_PUSH_TOKEN_STORAGE_KEY = 'mayumi_native_push_token';

function getStoredPushPreference() {
  return localStorage.getItem(PUSH_ENABLED_STORAGE_KEY);
}

function isPushEnabled() {
  return getStoredPushPreference() === 'true';
}

function getStoredNativePushPlayerId() {
  return localStorage.getItem(NATIVE_PUSH_PLAYER_ID_STORAGE_KEY) || '';
}

function setStoredNativePushPlayerId(playerId) {
  if (playerId) localStorage.setItem(NATIVE_PUSH_PLAYER_ID_STORAGE_KEY, playerId);
  else localStorage.removeItem(NATIVE_PUSH_PLAYER_ID_STORAGE_KEY);
}

function getStoredNativePushToken() {
  return localStorage.getItem(NATIVE_PUSH_TOKEN_STORAGE_KEY) || '';
}

function setStoredNativePushToken(token) {
  if (token) localStorage.setItem(NATIVE_PUSH_TOKEN_STORAGE_KEY, token);
  else localStorage.removeItem(NATIVE_PUSH_TOKEN_STORAGE_KEY);
}

function getCurrentPushSubscriptionValue(enabled) {
  if (!enabled) return false;
  return getStoredNativePushPlayerId() || getStoredNativePushToken() || true;
}

function getOneSignalPushSubscription(oneSignal) {
  if (!oneSignal || !oneSignal.User || !oneSignal.User.PushSubscription) return null;
  return oneSignal.User.PushSubscription;
}

async function getOneSignalPermissionState(oneSignal) {
  if (!oneSignal || !oneSignal.Notifications) return false;
  try {
    const permission = oneSignal.Notifications.permission;
    if (permission && typeof permission.then === 'function') {
      return !!(await permission);
    }
    return !!permission;
  } catch (e) {
    console.error('OneSignal permission read error:', e);
    return false;
  }
}

function getOneSignalSubscriptionValue(oneSignal) {
  const sub = getOneSignalPushSubscription(oneSignal);
  if (!sub) return true;
  return sub.id || sub.token || true;
}

async function syncWebPushState(options) {
  const opts = options || {};
  const OS = opts.oneSignal || window.OneSignalRef;
  const sub = getOneSignalPushSubscription(OS);
  if (!OS || !OS.Notifications || !sub) return false;

  const permissionGranted = await getOneSignalPermissionState(OS);
  if (!permissionGranted) {
    await applyPushEnabledState(false, { syncProfile: opts.syncProfile === true });
    return false;
  }

  try {
    if (opts.targetEnabled === false && sub.optedIn && typeof sub.optOut === 'function') {
      await sub.optOut();
    } else if (opts.targetEnabled === true && !sub.optedIn && typeof sub.optIn === 'function') {
      await sub.optIn();
    } else if (opts.reconcile !== false) {
      const savedPreference = getStoredPushPreference();
      if (savedPreference === 'false' && sub.optedIn && typeof sub.optOut === 'function') {
        await sub.optOut();
      } else if (savedPreference === 'true' && !sub.optedIn && typeof sub.optIn === 'function') {
        await sub.optIn();
      }
    }
  } catch (e) {
    console.error('OneSignal push sync error:', e);
  }

  const enabled = !!sub.optedIn;
  await applyPushEnabledState(enabled, {
    syncProfile: opts.syncProfile === true,
    subscriptionValue: enabled ? getOneSignalSubscriptionValue(OS) : false
  });
  return enabled;
}

async function syncPushPreferenceToProfile(enabled, subscriptionValue) {
  if (!_profile || !_profile.memberId) return;
  try {
    await postToGAS({
      type: 'updateUser',
      memberId: _profile.memberId,
      pushSubscription: enabled ? (subscriptionValue !== undefined ? subscriptionValue : getCurrentPushSubscriptionValue(true)) : false
    });
  } catch (e) {
    console.log('syncPushPreferenceToProfile error:', e);
  }
}

async function applyPushEnabledState(enabled, options) {
  const nextEnabled = enabled === true;
  const opts = options || {};

  localStorage.setItem(PUSH_ENABLED_STORAGE_KEY, nextEnabled ? 'true' : 'false');
  if (opts.clearNativePlayerId) setStoredNativePushPlayerId('');
  if (opts.clearNativeToken) setStoredNativePushToken('');
  updatePushUI(nextEnabled ? 'on' : 'off');

  if (opts.syncProfile !== false) {
    await syncPushPreferenceToProfile(nextEnabled, opts.subscriptionValue);
  }
  if (_profile && _profile.memberId) {
    syncCurrentDeviceSession({ silent: true }).catch(function (error) {
      console.error('syncCurrentDeviceSession after push setting error:', error);
    });
  }
}

// 起動時にlocalStorageから通知ボタンの状態を即座に復元
(function () {
  var saved = getStoredPushPreference();
  if (saved === 'true') updatePushUI('on');
})();

// ===== 初期化処理 (以前の定義は削除し、後半の定義に統一) =====

function updatePushUI(state) {
  const btn = document.getElementById('push-btn');
  if (!btn) return;
  if (state === 'on') {
    btn.textContent = '通知オン 🔔（タップでオフ）';
    btn.classList.add('secondary');
    btn.classList.remove('primary');
  } else {
    btn.textContent = '通知オフ（タップでオン）';
    btn.classList.add('primary');
    btn.classList.remove('secondary');
  }
  btn.disabled = false;
}

function updatePasscodeLoginUI() {
  const btn = document.getElementById('passcode-login-btn');
  const status = document.getElementById('passcodeLoginStatus');
  const desc = document.getElementById('passcodeLoginHelp');
  const available = !!(_profile && hasConfiguredLocalPasscode());
  const enabled = available && isPasscodeLoginEnabled();

  if (status) status.textContent = available ? (enabled ? 'オン' : 'オフ') : '未設定';
  if (desc) {
    desc.textContent = available
      ? (enabled
        ? 'アプリを開くたびにパスコード入力が必要です。'
        : '次回から起動時のパスコード入力を省略できます。')
      : 'まずはパスコードを設定してください。';
  }
  if (!btn) return;

  if (!available) {
    btn.textContent = '未設定';
    btn.disabled = true;
    btn.classList.add('secondary');
    btn.classList.remove('primary');
    return;
  }

  if (enabled) {
    btn.textContent = 'オン（タップでオフ）';
    btn.classList.add('secondary');
    btn.classList.remove('primary');
  } else {
    btn.textContent = 'オフ（タップでオン）';
    btn.classList.add('primary');
    btn.classList.remove('secondary');
  }
  btn.disabled = false;
}

async function togglePasscodeLoginSetting() {
  if (!_profile || !hasConfiguredLocalPasscode()) {
    showToast('まずパスコードを設定してください');
    return;
  }

  const nextEnabled = !isPasscodeLoginEnabled();
  const isConfirmed = await showAppConfirm(
    nextEnabled
      ? '起動時のパスコード入力をオンにしますか？\n次回からアプリを開くたびにログインが必要になります。'
      : '起動時のパスコード入力をオフにしますか？\n次回からアプリを開いたときのログインを省略できます。',
    {
      title: 'ログイン設定',
      confirmLabel: nextEnabled ? 'オンにする' : 'オフにする',
      cancelLabel: '戻る',
      confirmVariant: nextEnabled ? 'primary' : 'danger'
    }
  );
  if (!isConfirmed) return;

  setPasscodeLoginEnabled(nextEnabled);
  isPasscodeAuthenticated = true;
  if (!nextEnabled) {
    closePasscodeOverlay();
    markPasscodeUnlockSkippedOnce();
  }
  syncCurrentDeviceSession({ silent: true }).catch(function (error) {
    console.error('syncCurrentDeviceSession after togglePasscodeLoginSetting error:', error);
  });
  showToast(nextEnabled ? '起動時のパスコード入力をオンにしました' : '起動時のパスコード入力をオフにしました');
}


// 更新バッジの制御
function updateNavBadges() {
  if (!isDataLoaded) return; // 初回ロード完了までバッジ更新をスキップ
  let totalCount = 0;
  const currentPage = document.querySelector('.page.active')?.id;

  // 拡声器（お知らせ全体）の未読チェック
  // ブログ / カレンダー / 商品 / Pushのどれかで変更があればドットを表示
  const noticeBadge = document.getElementById('badge-notices');
  if (noticeBadge) {
    const currentNoticeHash = computeNoticeListHash();
    const lastSeenNoticeHash = localStorage.getItem('last_seen_notices_hash') || '';
    if (lastSeenNoticeHash && currentNoticeHash !== lastSeenNoticeHash && currentPage !== 'page-notices') {
      noticeBadge.style.display = 'block';
      totalCount++;
    } else {
      noticeBadge.style.display = 'none';
      if (currentPage === 'page-notices' || !lastSeenNoticeHash) {
        localStorage.setItem('last_seen_notices_hash', currentNoticeHash);
      }
    }
  }

  // ブログの更新チェック
  const blogBadge = document.getElementById('badge-blog');
  if (blogBadge && blogItems && blogItems.length > 0) {
    const lastSeenHash = localStorage.getItem('last_seen_blog_hash');
    const currentHash = blogItems.map(function (i) {
      return [i.updatedAt || '', i.date || '', i.title || '', i.body || ''].join('_');
    }).join('|');
    if (lastSeenHash && lastSeenHash !== currentHash && currentPage !== 'page-blog') {
      blogBadge.style.display = 'block';
      totalCount++;
    } else {
      blogBadge.style.display = 'none';
      if (currentPage === 'page-blog' || !lastSeenHash) {
        localStorage.setItem('last_seen_blog_hash', currentHash);
      }
    }
  }

  // カレンダーの更新チェック
  const calBadge = document.getElementById('badge-calendar');
  if (calBadge && calendarData && calendarData.length > 0) {
    const lastSeenHash = localStorage.getItem('last_seen_calendar_hash');
    const currentHash = calendarData.map(function (i) {
      return [i.updatedAt || '', i.date || '', i.title || '', i.desc || ''].join('_');
    }).join('|');
    if (lastSeenHash && lastSeenHash !== currentHash && currentPage !== 'page-calendar') {
      calBadge.style.display = 'block';
      totalCount++;
    } else {
      calBadge.style.display = 'none';
      if (currentPage === 'page-calendar' || !lastSeenHash) {
        localStorage.setItem('last_seen_calendar_hash', currentHash);
      }
    }
  }

  // ショップ商品の更新チェック
  const shopBadge = document.getElementById('badge-shop');
  if (shopBadge && PRODUCTS && PRODUCTS.length > 0) {
    const lastSeenHash = localStorage.getItem('last_seen_product_hash');
    const currentHash = computeProductListHash();
    if (lastSeenHash && lastSeenHash !== currentHash && currentPage !== 'page-shop') {
      shopBadge.style.display = 'block';
      totalCount++;
    } else {
      shopBadge.style.display = 'none';
      if (currentPage === 'page-shop' || !lastSeenHash) {
        localStorage.setItem('last_seen_product_hash', currentHash);
      }
    }
  }

  // App Badge API (ホーム画面のアイコン通知数)
  if ('setAppBadge' in navigator) {
    if (totalCount > 0) {
      navigator.setAppBadge(totalCount).catch(function (err) { console.log('Badge error:', err); });
    } else {
      navigator.clearAppBadge().catch(function (err) { console.log('Badge error:', err); });
    }
  }

  // Push UIの初期状態回復
  if (isPushEnabled()) {
    updatePushUI('on');
  }
}

function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2300);
}

function escapeHtml(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeRegExp(text) {
  return String(text || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function markManagedTextFormatTokens(text) {
  return String(text || '')
    .replace(/<\s*(strong|b)\s*>/gi, '[[FMT_B_OPEN]]')
    .replace(/<\s*\/\s*(strong|b)\s*>/gi, '[[FMT_B_CLOSE]]')
    .replace(/<\s*u\s*>/gi, '[[FMT_U_OPEN]]')
    .replace(/<\s*\/\s*u\s*>/gi, '[[FMT_U_CLOSE]]');
}

function restoreManagedTextFormatTokens(text) {
  return String(text || '')
    .replace(/\[\[FMT_B_OPEN\]\]/g, '<strong>')
    .replace(/\[\[FMT_B_CLOSE\]\]/g, '</strong>')
    .replace(/\[\[FMT_U_OPEN\]\]/g, '<u>')
    .replace(/\[\[FMT_U_CLOSE\]\]/g, '</u>');
}

function renderManagedTextHtml(text, options) {
  const opts = options || {};
  let normalized = String(text || '').replace(/\r\n?/g, '\n');
  const inlineImages = [];

  if (opts.stripInlineImages) {
    normalized = normalized.replace(/(^|\n)\s*📷\s*https?:\/\/\S+/g, '$1');
  } else {
    normalized = normalized.replace(/(^|\n)\s*📷\s*(https?:\/\/\S+)/g, function (match, prefix, url) {
      const imageUrl = getContentDisplayImageUrl(url);
      if (!imageUrl) return prefix || '';
      const token = `[[FMT_IMG_${inlineImages.length}]]`;
      inlineImages.push({ token: token, url: imageUrl });
      return (prefix || '') + token;
    });
  }

  normalized = markManagedTextFormatTokens(normalized);
  let safe = restoreManagedTextFormatTokens(escapeHtml(normalized));
  if (opts.preserveLineBreaks !== false) {
    safe = safe.replace(/\n/g, '<br>');
  }

  const inlineImageClass = escapeHtml(String(opts.inlineImageClass || 'blog-inline-image'));
  const inlineImageAlt = String(opts.inlineImageAlt || '画像');
  inlineImages.forEach(function (item, index) {
    safe = safe.replace(
      new RegExp(escapeRegExp(item.token), 'g'),
      `<img src="${escapeHtml(item.url)}" class="${inlineImageClass}" alt="${escapeHtml(inlineImageAlt + ' ' + (index + 1))}">`
    );
  });

  return safe;
}

function normalizeSupportText(text) {
  return String(text || '')
    .toLowerCase()
    .normalize('NFKC')
    .replace(/[！!？?。、,.・/／\\|｜()\[\]{}（）「」『』【】<>＜＞:：;；"'`´~〜\-_=+*&^%$#@]/g, '')
    .replace(/\s+/g, '');
}

function getSupportKeywordList(item) {
  return String((item && item.keywords) || '')
    .split(/[\n,、，\/\s]+/)
    .map(function (part) { return part.trim(); })
    .filter(Boolean);
}

function scoreSupportFaq(messageNorm, item) {
  if (!messageNorm || !item) return 0;

  let score = 0;
  let keywordHits = 0;
  const questionNorm = normalizeSupportText(item.question);
  const categoryNorm = normalizeSupportText(item.category);

  if (questionNorm) {
    if (messageNorm === questionNorm) score += 40;
    if (questionNorm.length >= 8 && messageNorm.indexOf(questionNorm) !== -1) score += 18;
    if (messageNorm.length >= 8 && questionNorm.length >= 8 && questionNorm.indexOf(messageNorm) !== -1) score += 10;
  }
  if (categoryNorm && messageNorm.indexOf(categoryNorm) !== -1) {
    score += 3;
  }

  getSupportKeywordList(item).forEach(function (keyword) {
    const norm = normalizeSupportText(keyword);
    if (!norm) return;
    if (messageNorm.indexOf(norm) !== -1) {
      keywordHits += 1;
      score += norm.length >= 4 ? 6 : 3;
    }
  });

  if (keywordHits >= 2) score += keywordHits * 2;
  return score;
}

function getSupportFaqSource() {
  return supportFaqItems.length ? supportFaqItems : SUPPORT_FAQ_FALLBACK;
}

function getCombinedSupportKnowledge() {
  return getSupportFaqSource().concat(SUPPORT_APP_GUIDE);
}

function getDefaultSupportQuestions(limit) {
  return Array.from(new Set(
    getCombinedSupportKnowledge()
      .map(function (item) { return item.question; })
      .filter(Boolean)
  )).slice(0, limit || 4);
}

function isLikelyAppSupportQuestion(messageNorm) {
  return SUPPORT_APP_KEYWORDS.some(function (keyword) {
    return messageNorm.indexOf(normalizeSupportText(keyword)) !== -1;
  });
}

function hasSupportKeyword(messageNorm, keyword) {
  const keywordNorm = normalizeSupportText(keyword);
  return !!keywordNorm && messageNorm.indexOf(keywordNorm) !== -1;
}

function hasAnySupportKeyword(messageNorm, keywords) {
  return (keywords || []).some(function (keyword) {
    return hasSupportKeyword(messageNorm, keyword);
  });
}

function hasAllSupportKeywordGroups(messageNorm, keywordGroups) {
  return (keywordGroups || []).every(function (group) {
    return hasAnySupportKeyword(messageNorm, group);
  });
}

function isSupportBlogArchiveQuestion(messageNorm) {
  return hasAnySupportKeyword(messageNorm, ['ブログ']) &&
    hasAnySupportKeyword(messageNorm, ['過去', '履歴', '一覧', '記事', '以前', 'バックナンバー']);
}

function isSupportOrderHistoryQuestion(messageNorm) {
  return hasAnySupportKeyword(messageNorm, ['注文履歴']) ||
    (hasAnySupportKeyword(messageNorm, ['注文', '購入', 'ショップ', '受取', '受け取り']) &&
      hasAnySupportKeyword(messageNorm, ['履歴', '状況', '状態', '受付中', '完了', 'キャンセル']));
}

function getMergedSupportSuggestions(preferred) {
  const merged = [];
  (preferred || []).forEach(function (question) {
    if (question && merged.indexOf(question) === -1) merged.push(question);
  });
  getDefaultSupportQuestions(6).forEach(function (question) {
    if (question && merged.indexOf(question) === -1) merged.push(question);
  });
  return merged.slice(0, 4);
}

function buildSupportText(answer) {
  if (Array.isArray(answer)) {
    return answer.filter(Boolean).join('\n');
  }
  return String(answer || '');
}

function buildSupportReply(answer, suggestions, matched, topic) {
  return {
    matched: matched !== false,
    answer: buildSupportText(answer),
    suggestions: getMergedSupportSuggestions(suggestions),
    topic: topic || ''
  };
}

function getSupportIntent(messageNorm) {
  if (hasAnySupportKeyword(messageNorm, ['できない', '使えない', '表示されない', '出ない', '開けない', '届かない', '読めない', '失敗', 'エラー', '不具合', '重い', '遅い', 'フリーズ', '消えない'])) {
    return 'trouble';
  }
  if (hasAnySupportKeyword(messageNorm, ['今', '現在', 'いま', '状況', '状態', '確認', '残り', '何個', 'いくつ', '何枚', '登録済', '未登録'])) {
    return 'status';
  }
  if (hasAnySupportKeyword(messageNorm, ['どこ', 'どこで', 'どこから', '場所', '開き方', '表示場所'])) {
    return 'where';
  }
  if (hasAnySupportKeyword(messageNorm, ['できる', 'できますか', '可能', '使える', '対応'])) {
    return 'can';
  }
  if (hasAnySupportKeyword(messageNorm, ['とは', '何', '意味'])) {
    return 'what';
  }
  return 'how';
}

function detectSupportTopic(messageNorm) {
  if (!messageNorm) return '';
  if (hasAnySupportKeyword(messageNorm, ['アップデートが必要', 'アップデート', '更新', '最新版', 'バージョン'])) return 'update';
  if (hasAnySupportKeyword(messageNorm, ['予約', 'よやく', 'アポイント'])) return 'reservation';
  if (hasAnySupportKeyword(messageNorm, ['注文履歴', '受け取りました', '受取報告', '受取完了']) || isSupportOrderHistoryQuestion(messageNorm)) return 'order-history';
  if (hasAnySupportKeyword(messageNorm, ['カート', '買い物かご'])) return 'cart';
  if (hasAnySupportKeyword(messageNorm, ['ショップ', '商品', '注文', '購入', '支払い', '現金'])) return 'shop';
  if (hasAnySupportKeyword(messageNorm, ['特典', 'プレゼント', '期限', 'チケット'])) return 'reward';
  if (hasAnySupportKeyword(messageNorm, ['スタンプ', 'qr', 'qrコード', 'qrcode', 'カメラ', '新しいカード', '枚目'])) return 'stamp';
  if (hasAnySupportKeyword(messageNorm, ['通知', 'push', 'プッシュ'])) return 'notification';
  if (hasAnySupportKeyword(messageNorm, ['プロフィール', '会員id', 'memberid', '会員番号', 'アイコン', 'アバター', 'バナー', '画像'])) return 'profile';
  if (hasAnySupportKeyword(messageNorm, ['メニュー', '施術'])) return 'menu';
  if (hasAnySupportKeyword(messageNorm, ['カレンダー', 'イベント', '予定', '日程'])) return 'calendar';
  if (hasAnySupportKeyword(messageNorm, ['news', 'ニュース', 'お知らせ', 'ブログ', 'お知らせ一覧', '通知一覧', '📢', '新着', 'つぶやき', 'まゆみのブログ', 'まゆみのつぶやき'])) return 'news';
  if (hasAnySupportKeyword(messageNorm, ['line', 'ライン', 'instagram', 'facebook', 'ホームページ', '公式サイト', 'sns', '問い合わせ', 'お問い合わせ'])) return 'links';
  if (hasAnySupportKeyword(messageNorm, ['チャット', 'サポート', 'ボット', '相談'])) return 'support-chat';
  if (hasAnySupportKeyword(messageNorm, ['ホーム', 'トップ'])) return 'home';
  if (hasAnySupportKeyword(messageNorm, ['アプリ', '使い方', 'できること', '機能'])) return 'app-overview';
  return '';
}

function isSupportFollowUpQuestion(messageNorm) {
  if (!messageNorm) return false;
  if (detectSupportTopic(messageNorm)) return false;
  return hasAnySupportKeyword(messageNorm, ['それ', 'その', 'さっき', '前の', '今の', 'この機能', 'この画面', 'どこ', 'どこから', '方法', 'やり方', '見方', '開き方', 'できますか', 'できる', '確認', '変更', '解除', 'オン', 'オフ', 'どうやって']);
}

function resolveSupportTopic(messageNorm) {
  const explicitTopic = detectSupportTopic(messageNorm);
  if (explicitTopic) return explicitTopic;
  if (isSupportFollowUpQuestion(messageNorm) && lastSupportTopic) {
    return lastSupportTopic;
  }
  return '';
}

function buildFeatureSupportReply(topic, answer, suggestions) {
  return buildSupportReply(answer, suggestions, true, topic);
}

function getFeatureSupportReply(messageNorm) {
  const topic = resolveSupportTopic(messageNorm);
  if (!topic) return null;
  const intent = getSupportIntent(messageNorm);

  if (topic === 'update') {
    if (hasAnySupportKeyword(messageNorm, ['アップデートが必要', 'ストア'])) {
      return buildFeatureSupportReply(
        topic,
        [
          '「アップデートが必要です」と表示された場合は、案内の「アップデートする」を押してください。',
          'App Store が開いたらアプリ本体を更新し、更新後にアプリを開き直してください。',
          '通常の商品・お知らせ更新は画面上部の🔄ボタンで反映できますが、必須更新表示が出た場合はアプリ更新が必要です。'
        ],
        ['最新情報への更新方法を知りたい', '通知をオンにする方法を知りたい']
      );
    }
    if (intent === 'trouble') {
      return buildFeatureSupportReply(
        topic,
        [
          '更新で変わらない場合は、まず画面上部の🔄ボタンで最新情報を再取得してください。',
          'それでも「アップデートが必要です」と表示される場合は、受付にて配布されている最新版のインストール方法をご確認ください。',
          '商品・お知らせ・カレンダーなどの最新情報は🔄ボタン、アプリそのものの更新は院内の案内に従う、と覚えていただくと分かりやすいです。'
        ],
        ['最新情報への更新方法を知りたい', 'アプリのアップデート方法は？']
      );
    }
    return buildFeatureSupportReply(
      topic,
      [
        '更新には2種類あります。',
        '1. 商品・お知らせ・カレンダーなどの最新情報は、画面上部の🔄ボタンで再取得でき、その場で最新状態に反映されます。',
        '2. 「アップデートが必要です」と大きく表示された場合は、当院から案内されている方法で最新版を入れ直してください。',
        '通常の情報更新（🔄ボタン）は再インストール不要です。'
      ],
      ['最新情報への更新方法を知りたい', 'アプリのアップデート方法は？']
    );
  }

  if (topic === 'reservation') {
    return buildFeatureSupportReply(
      topic,
      hasAnySupportKeyword(messageNorm, ['変更', 'キャンセル', '取り消し'])
        ? [
          '現在、アプリから予約の変更やキャンセルはできません。',
          '予約変更・キャンセル・個別相談は、公式LINEからご連絡ください。',
          '公式LINEはホーム画面またはマイページの「公式サイト・SNS」から開けます。'
        ]
        : [
          '現在、アプリから直接予約する機能はありません。',
          'ご予約や個別相談は公式LINEをご利用ください。',
          'ホーム画面またはマイページの「公式サイト・SNS」から公式LINEを開けます。'
        ],
      ['公式LINEやSNSの開き方を知りたい', 'メニュー一覧の見方を知りたい']
    );
  }

  if (topic === 'profile') {
    if (hasAnySupportKeyword(messageNorm, ['アイコン', 'アバター', 'バナー', '画像', '写真'])) {
      return buildFeatureSupportReply(
        topic,
        [
          'プロフィール画像の変更方法です。',
          '1. マイページの「✏️ プロフィールを編集」を開きます。',
          '2. アイコン画像またはバナー画像の変更ボタンから画像を選びます。',
          '3. 保存すると、マイページやホーム画面の表示に反映されます。'
        ],
        ['プロフィールの変更方法を知りたい', 'ホーム画面の見方を知りたい']
      );
    }
    if (intent === 'status') {
      return buildFeatureSupportReply(
        topic,
        _profile
          ? [
            `プロフィールは登録済みです。お名前は${_profile.name || '未設定'}、会員IDは${_profile.memberId || '未発行'}です。`,
            '変更したい場合はマイページの「プロフィールを編集」から更新できます。'
          ]
          : [
            '現在プロフィールは未登録です。',
            '初回起動時の案内、またはマイページの「プロフィールを編集」から登録してください。'
          ],
        ['プロフィールの登録方法を知りたい', 'プロフィールの変更方法を知りたい']
      );
    }
    return buildFeatureSupportReply(
      topic,
      _profile
        ? [
          'プロフィールの変更方法です。',
          '1. マイページを開きます。',
          '2. 「✏️ プロフィールを編集」を押します。',
          '3. お名前、電話番号、生年月日、住所、画像を変更して保存します。'
        ]
        : [
          'プロフィールの登録方法です。',
          '1. 初回起動時は「はじめて登録する方はこちら」を選びます。',
          '2. お名前、電話番号、生年月日、住所を入力して保存します。',
          '3. 以前登録したことがある方は、新規登録ではなく「以前登録した方はこちら ↺」から復元してください。'
        ],
      ['プロフィールの登録方法を知りたい', 'プロフィールの変更方法を知りたい']
    );
  }

  if (topic === 'notification') {
    if (intent === 'trouble') {
      return buildFeatureSupportReply(
        topic,
        [
          '通知が届かない場合は、次を確認してください。',
          '1. マイページの通知設定がオンになっているか。',
          '2. iPhone やブラウザの設定で通知が許可されているか。',
          '3. インターネット接続が安定しているか。',
          isPushEnabled() ? '現在この端末のアプリ内設定はオンです。端末側の通知許可もご確認ください。' : '現在この端末のアプリ内設定はオフです。まずアプリ内でオンにしてください。'
        ],
        ['通知のオンオフ方法を知りたい', '最新のお知らせの見方を知りたい']
      );
    }
    return buildFeatureSupportReply(
      topic,
      hasAnySupportKeyword(messageNorm, ['オフ', '解除'])
        ? [
          '通知はオフにもできます。',
          'マイページの「🔔 通知設定」ボタンを押すと、オンとオフを切り替えられます。',
          '端末側の通知許可をオフにしたい場合は、iPhone の通知設定からも変更できます。'
        ]
        : [
          '通知設定の方法です。',
          '1. マイページを開きます。',
          '2. 「🔔 通知設定」ボタンを押してオンまたはオフに切り替えます。',
          '3. 端末側で通知が拒否されている場合は、iPhone の通知設定もご確認ください。'
        ],
      ['通知を受け取りたい', '通知をオフにしたい']
    );
  }

  if (topic === 'order-history') {
    if (hasAnySupportKeyword(messageNorm, ['キャンセル'])) {
      return buildFeatureSupportReply(
        topic,
        [
          '注文キャンセルの方法です。',
          '1. マイページの「📋 ご注文履歴」を開きます。',
          '2. 受付中の注文に表示される「キャンセルする」を押します。',
          '3. キャンセル後、その注文は履歴から表示されなくなります。'
        ],
        ['注文履歴の確認方法を知りたい', '商品の注文方法を知りたい']
      );
    }
    if (hasAnySupportKeyword(messageNorm, ['受け取りました', '受取報告', '受取完了'])) {
      return buildFeatureSupportReply(
        topic,
        [
          '受け取り完了の報告方法です。',
          '1. マイページの「📋 ご注文履歴」を開きます。',
          '2. 対象注文の「受け取りました」を押します。',
          '3. 更新後、その注文は履歴に残らなくなります。'
        ],
        ['注文履歴の確認方法を知りたい', '商品の注文方法を知りたい']
      );
    }
    if (intent === 'trouble') {
      return buildFeatureSupportReply(
        topic,
        [
          '注文履歴が見えない場合は、次を確認してください。',
          '1. プロフィール登録が完了しているか。',
          '2. 画面上部の🔄ボタンで最新情報を再取得したか。',
          '3. 受付中または受け取り前の注文があるか。',
          'キャンセル済み・受取済みの注文は履歴に表示されません。'
        ],
        ['プロフィールの登録方法を知りたい', '最新情報への更新方法を知りたい']
      );
    }
    return buildFeatureSupportReply(
      topic,
      [
        '注文履歴の確認方法です。',
        '1. 下部メニューからマイページを開きます。',
        '2. 「📋 ご注文履歴」を確認します。',
        '3. 受付中の注文はキャンセル、受け取り前の注文は「受け取りました」で更新できます。'
      ],
      ['注文をキャンセルしたい', '受け取り完了の報告方法を知りたい']
    );
  }

  if (topic === 'cart') {
    return buildFeatureSupportReply(
      topic,
      intent === 'status'
        ? (cart.length
          ? [
            `現在のカートは${cart.reduce(function (sum, item) { return sum + Number(item.qty || 0); }, 0)}点です。`,
            '画面右上の🛒から中身を確認し、個数変更や削除ができます。'
          ]
          : [
            '現在カートには商品が入っていません。',
            'ショップで商品を選び、「カートに追加する」を押すと入ります。'
          ])
        : [
          'カートの使い方です。',
          '1. ショップで商品を選び「カートに追加する」を押します。',
          '2. 画面右上の🛒を押すとカートを開けます。',
          '3. カートでは個数変更や削除ができます。'
        ],
      ['カートの中身を確認するには？', '商品の注文方法を知りたい']
    );
  }

  if (topic === 'shop') {
    if (hasAnySupportKeyword(messageNorm, ['支払い', '決済', '現金'])) {
      return buildFeatureSupportReply(
        topic,
        [
          '商品注文のお支払い方法は現在「現金払い」です。',
          'ご来院時に受付でお支払いください。'
        ],
        ['商品の注文方法を知りたい', '受け取り方法は？']
      );
    }
    if (hasAnySupportKeyword(messageNorm, ['受取', '受け取り', '配送', '宅配'])) {
      return buildFeatureSupportReply(
        topic,
        [
          '商品はすべて院内受け取りです。',
          '注文後、ご来院時にスタッフへお声がけください。'
        ],
        ['商品の注文方法を知りたい', '注文履歴の確認方法を知りたい']
      );
    }
    return buildFeatureSupportReply(
      topic,
      [
        '商品の注文手順です。',
        '1. 下部メニューの「🛍 ショップ」を開きます。',
        '2. 商品をタップして詳細を確認します。',
        '3. 個数を選んで「カートに追加する🛒」を押します。',
        '4. 画面右上の🛒からカートを開き、「ご注文を確定する」で完了です。'
      ],
      ['カートの使い方を知りたい', '注文履歴の確認方法を知りたい']
    );
  }

  if (topic === 'stamp') {
    if (hasAnySupportKeyword(messageNorm, ['新しいカード', '次のカード', 'カード開始', 'カード取得'])) {
      return buildFeatureSupportReply(
        topic,
        [
          '新しいスタンプカードの始め方です。',
          '1. スタンプが10個たまると、ホーム画面に「🎯 特典ガチャを回す」が表示されます。',
          '2. 特典ガチャを回すと、「🌸 新しいスタンプカードを取得」が表示されます。',
          '3. そのボタンを押すと次のカードが始まり、スタンプが0個に戻ってカード番号が1つ進みます。'
        ],
        ['スタンプが10個たまったらどうなりますか？', 'スタンプ特典の使い方を知りたい']
      );
    }
    if (hasAnySupportKeyword(messageNorm, ['qr', 'qrコード', 'qrcode', 'カメラ', '読み取り'])) {
      return buildFeatureSupportReply(
        topic,
        intent === 'trouble'
          ? [
            'QRコードが読み取れない場合は、次を確認してください。',
            '1. カメラ権限が許可されているか。',
            '2. QRコードに汚れや傷がないか。',
            '3. その日のスタンプをすでに取得していないか。',
            '4. 10個達成済みなら先に特典ガチャを回し、その後に新しいカードを始めてください。'
          ]
          : [
            'スタンプ取得の手順です。',
            '1. ホーム画面の「📷 カメラを起動して読み取る」を押します。',
            '2. カメラ権限の案内が出た場合は「許可」を選びます。',
            '3. 院内のQRコードを読み取ります。',
            '4. 読み取りに成功するとスタンプが1つ追加されます。'
          ],
        ['スタンプは1日何個取得できますか？', '新しいカードを始めるには？']
      );
    }
    if (hasAnySupportKeyword(messageNorm, ['10個', '達成', 'たまった', '満了'])) {
      return buildFeatureSupportReply(
        topic,
        [
          'スタンプが10個たまるとホーム画面から特典ガチャを回せます。',
          'ガチャ結果はマイページの「🎁 スタンプ・特典履歴」で確認できます。',
          'ガチャ後は、ホーム画面の「🌸 新しいスタンプカードを取得」から次のカードを始められます。'
        ],
        ['スタンプ特典の使い方を知りたい', '新しいカードを始めるには？']
      );
    }
    return buildFeatureSupportReply(
      topic,
      [
        'スタンプ機能のご案内です。',
        '1. スタンプはホーム画面で確認します。',
        '2. 院内QRコードを読むと1日1回まで追加されます。',
        `3. 現在のカードは${stampCardNum}枚目、スタンプは${stampCount}個です。`,
        '4. 10個たまると特典ガチャを回せます。'
      ],
      ['スタンプの集め方を知りたい', 'スタンプ特典の使い方を知りたい']
    );
  }

  if (topic === 'reward') {
    const availableRewards = (EARNED_REWARDS || []).filter(function (reward) {
      return reward && !reward.used;
    });
    if (hasAnySupportKeyword(messageNorm, ['期限', 'いつまで'])) {
      return buildFeatureSupportReply(
        topic,
        [
          '特典の受取期限は、スタンプ10個を達成した日から1か月です。',
          'マイページの「🎁 スタンプ・特典履歴」に期限が表示されます。',
          '受け取りについては受付へ直接お問い合わせください。'
        ],
        ['スタンプ特典の使い方を知りたい', '特典はどこで確認できますか？']
      );
    }
    if (intent === 'status') {
      return buildFeatureSupportReply(
        topic,
        availableRewards.length
          ? [
            `現在、未使用の特典が${availableRewards.length}件あります。`,
            'マイページの「🎁 スタンプ・特典履歴」で確認できます。',
            '特典は達成当日から使用でき、一度使用すると再使用できません。'
          ]
          : [
            '現在、未使用の特典はありません。',
            'スタンプが10個たまると特典ガチャを回せるようになり、結果がマイページに表示されます。'
          ],
        ['特典はどこで確認できますか？', 'スタンプの集め方を知りたい']
      );
    }
    return buildFeatureSupportReply(
      topic,
      [
        'スタンプ特典の使い方です。',
        '1. マイページの「🎁 スタンプ・特典履歴」を開きます。',
        '2. 獲得済みの特典を確認し、必要なら「使用する」を押します。',
        '3. 特典は達成当日から使用できます。',
        '4. 一度使用すると再使用できません。受け取りは受付へ直接お問い合わせください。'
      ],
      ['特典に有効期限はありますか？', 'スタンプが10個たまったらどうなりますか？']
    );
  }

  if (topic === 'menu') {
    if (hasAnySupportKeyword(messageNorm, ['カテゴリ', '絞り込み', '全て'])) {
      return buildFeatureSupportReply(
        topic,
        [
          'メニュー一覧ではカテゴリ絞り込みができます。',
          '1. ホーム画面の「メニュー一覧を見る🍴」を押します。',
          '2. 右上のカテゴリ選択から見たいカテゴリを選びます。',
          '3. 「全て」を選ぶとすべてのメニューが表示されます。'
        ],
        ['メニュー一覧の見方を知りたい', '予約の方法を知りたい']
      );
    }
    return buildFeatureSupportReply(
      topic,
      [
        'メニュー一覧の見方です。',
        '1. ホーム画面の「メニュー一覧を見る🍴」を押します。',
        '2. 右上のカテゴリ選択で絞り込みできます。',
        '3. 各メニューをタップすると詳細や画像を確認できます。',
        '4. 実際のご予約や個別相談は公式LINEからお願いします。'
      ],
      ['予約の方法を知りたい', 'ホーム画面の見方を知りたい']
    );
  }

  if (topic === 'calendar') {
    return buildFeatureSupportReply(
      topic,
      [
        'カレンダーの見方です。',
        '1. 下部メニューの「📅 カレンダー」を開きます。',
        '2. 左右の矢印で月を切り替えます。',
        '3. 日付を押すと、その日の予定やイベントを確認できます。',
        '4. 休＝休診日、往＝往診日、イ＝イベントです。'
      ],
      ['イベントカレンダーの見方を知りたい', '最新のお知らせの見方を知りたい']
    );
  }

  if (topic === 'news') {
    if (hasAnySupportKeyword(messageNorm, ['お知らせ一覧', '通知一覧', '拡声器', '📢'])) {
      return buildFeatureSupportReply(
        topic,
        [
          'お知らせ一覧の見方です。',
          '1. 画面上部の📢ボタンを押します。',
          '2. NEWS、カレンダー、ショップ、ホームの更新情報を新しい順に確認できます。',
          '3. カテゴリ選択で絞り込みできます。'
        ],
        ['お知らせ一覧の見方を知りたい', 'NEWSページの使い方を知りたい']
      );
    }
    if (hasAnySupportKeyword(messageNorm, ['カテゴリ', '絞り込み', '全て'])) {
      return buildFeatureSupportReply(
        topic,
        [
          'NEWSページではカテゴリの絞り込みができます。',
          '1. 下部メニューの「💬 NEWS」を開きます。',
          '2. 右上のカテゴリ選択から見たいカテゴリを選びます。',
          '3. 「全て」を選ぶとすべての記事が表示されます。'
        ],
        ['NEWSページの使い方を知りたい', 'まゆみのつぶやきはどこで見られますか？']
      );
    }
    if (hasAnySupportKeyword(messageNorm, ['まゆみのブログ', 'ブログ'])) {
      return buildFeatureSupportReply(
        topic,
        [
          '「まゆみのブログ」は、マイページの「🔗 公式サイト・SNS」から閲覧できる外部ブログです。',
          '以前から続いている院長のブログ記事をじっくり読むことができます。',
          'NEWSにある「まゆみのつぶやき」とは内容が異なりますので、ぜひ両方チェックしてみてくださいね。'
        ],
        ['まゆみのブログとは何ですか？', 'まゆみのつぶやきはどこで見られますか？']
      );
    }
    if (hasAnySupportKeyword(messageNorm, ['まゆみのつぶやき', 'つぶやき'])) {
      return buildFeatureSupportReply(
        topic,
        [
          '「まゆみのつぶやき」は、NEWSページ内の特定のカテゴリです。',
          '1. 下部メニューの「💬 NEWS」を開きます。',
          '2. 右上のカテゴリ選択から「まゆみのつぶやき」を選んでください。',
          'アプリ内で手軽に読める院長からのメッセージや、最新の活動報告などが掲載されています。'
        ],
        ['まゆみのつぶやきはどこで見られますか？', 'まゆみのブログとは何ですか？']
      );
    }
    return buildFeatureSupportReply(
      topic,
      isSupportBlogArchiveQuestion(messageNorm)
        ? [
          '過去の更新情報は、画面上部の📢ボタンから開く「お知らせ一覧」で確認できます。',
          'NEWS、カレンダー、ショップ、ホームの更新情報を新しい順でまとめて見られます。'
        ]
        : [
          'お知らせの見方です。',
          '1. 下部メニューの「💬 NEWS」で記事一覧を確認できます。',
          '2. 画面上部の📢ボタンでは、NEWS、カレンダー、ショップ、ホームの更新情報をまとめて確認できます。',
          '3. 新着があるとアイコンに赤いドットが表示されます。'
        ],
      ['ブログや過去のお知らせの見方を知りたい', 'まゆみのブログとは何ですか？']
    );
  }

  if (topic === 'links') {
    return buildFeatureSupportReply(
      topic,
      [
        '公式サイトやSNSの開き方です。',
        '1. ホーム画面またはマイページの「🔗 公式サイト・SNS」を開きます。',
        '2. 公式ホームページ、Instagram、Facebook、公式LINEを選んで開けます。',
        '3. 診療や個別相談は公式LINEをご利用ください。'
      ],
      ['公式LINEやSNSの開き方を知りたい', '予約の方法を知りたい']
    );
  }

  if (topic === 'support-chat') {
    return buildFeatureSupportReply(
      topic,
      [
        'このチャットでは、現在アプリに実装されている機能の使い方をご案内しています。',
        'たとえば、登録方法、復元方法、パスコード、注文方法、注文履歴、スタンプ、特典、通知設定、NEWS、お知らせ一覧、更新方法などに回答できます。',
        '診療相談や個別のご予約内容の相談は公式LINEをご利用ください。'
      ],
      ['このアプリでできることを知りたい', '予約の方法を知りたい']
    );
  }

  if (topic === 'home') {
    return buildFeatureSupportReply(
      topic,
      [
        'ホーム画面でできることです。',
        '1. スタンプQRの読み取り',
        '2. 現在のスタンプカード確認',
        '3. おすすめ商品や最新のNEWS確認',
        '4. メニュー一覧への移動',
        '5. 公式サイト・SNSへのアクセス'
      ],
      ['スタンプの集め方を知りたい', 'メニュー一覧の見方を知りたい']
    );
  }

  if (topic === 'app-overview') {
    return getSupportOverviewReply();
  }

  return null;
}

function getSupportStatusReply(messageNorm) {
  const asksCurrent = hasAnySupportKeyword(messageNorm, ['今', '現在', 'いま', '状況', '状態', '確認']);
  const asksStatusOnly = asksCurrent || hasAnySupportKeyword(messageNorm, ['なってる', 'なっている', 'されてる', 'されている', '登録済', '未登録']);

  if (hasAnySupportKeyword(messageNorm, ['スタンプ']) && hasAnySupportKeyword(messageNorm, ['残り', '何個', 'いくつ', '現在', '今', '状況', '状態'])) {
    const remain = Math.max(0, 10 - stampCount);
    return buildSupportReply(
      stampCount >= 10
        ? [
          `現在は${stampCardNum}枚目のカードでスタンプ${stampCount}個です。10個達成済みです。`,
          hasCurrentCardReward()
            ? '次の流れ: 1. マイページの「スタンプ・特典履歴」を確認 2. 必要なら特典を使用 3. ホームの「新しいスタンプカードを取得」で次のカードを開始してください。'
            : '次の流れ: 1. ホームの「特典ガチャを回す」でごほうびを受け取る 2. マイページの「スタンプ・特典履歴」で確認 3. ガチャ後に「新しいスタンプカードを取得」で次のカードを開始してください。',
          '補足: 来院スタンプは1日1回までです。'
        ]
        : [
          `現在は${stampCardNum}枚目のカードでスタンプ${stampCount}個です。あと${remain}個で特典ガチャを回せます。`,
          '追加方法: ホーム画面の「カメラを起動して読み取る」から院内QRコードを読み取ってください。',
          '補足: 来院スタンプは1日1回までです。'
        ],
      ['スタンプの集め方を知りたい', 'スタンプ特典の使い方を知りたい'],
      true,
      'stamp'
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['カート', '買い物かご']) && (asksCurrent || hasAnySupportKeyword(messageNorm, ['何個', 'いくつ', '合計']))) {
    const itemCount = cart.reduce(function (sum, item) { return sum + Number(item.qty || 0); }, 0);
    const total = cart.reduce(function (sum, item) {
      const product = PRODUCTS[item.idx];
      return sum + (product ? getProductPricing(product, Number(item.qty || 0)).total : 0);
    }, 0);
    return buildSupportReply(
      itemCount > 0
        ? [
          `現在のカートは${itemCount}点、合計は¥${total.toLocaleString()}です。`,
          'カート画面では個数変更や削除ができます。',
          '内容を確認したら「ご注文を確定する」を押してください。'
        ]
        : [
          '現在カートには商品が入っていません。',
          '下部メニューの「ショップ」で商品を選ぶとカートに追加できます。'
        ],
      ['カートの使い方を知りたい', '商品の注文方法を知りたい'],
      true,
      'cart'
    );
  }

  if ((hasAnySupportKeyword(messageNorm, ['プロフィール']) && asksStatusOnly) || hasAnySupportKeyword(messageNorm, ['会員id', 'memberid', '会員番号'])) {
    return buildSupportReply(
      _profile
        ? [
          `プロフィールは登録済みです。お名前は${_profile.name || '未設定'}、会員IDは${_profile.memberId || '未発行'}です。`,
          '変更したい場合は、マイページの「プロフィールを編集」からお名前、電話番号、生年月日、住所、画像を更新できます。'
        ]
        : [
          '現在プロフィールは未登録です。',
          '初回起動時の案内、またはマイページの「プロフィールを編集」から登録してください。'
        ],
      ['プロフィールの登録方法を知りたい', 'プロフィールの変更方法を知りたい'],
      true,
      'profile'
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['通知', 'push', 'プッシュ']) && asksStatusOnly) {
    const isEnabled = isPushEnabled();
    return buildSupportReply(
      isEnabled
        ? [
          '現在この端末では通知がオンの状態です。',
          'オフにしたい場合は、マイページの通知設定ボタンから切り替えられます。'
        ]
        : [
          '現在この端末では通知はオフです。',
          'マイページの通知設定ボタンからオンにできます。'
        ],
      ['通知のオンオフ方法を知りたい', '最新のお知らせの見方を知りたい'],
      true,
      'notification'
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['パスコード', 'ログイン時']) && asksStatusOnly) {
    return buildSupportReply(
      hasConfiguredLocalPasscode()
        ? [
          `この端末のパスコード設定は完了しており、起動時のパスコード入力は現在${isPasscodeLoginEnabled() ? 'オン' : 'オフ'}です。`,
          '変更したい場合は、マイページの「ログイン時のパスコード」からオン・オフを切り替えられます。'
        ]
        : [
          'この端末ではまだパスコード設定が完了していません。',
          '初回設定画面、または既存会員向けの設定画面で4桁または6桁のパスコードを設定してください。'
        ],
      ['起動時のパスコード設定について知りたい', 'パスコードの変更方法を知りたい'],
      true,
      'passcode'
    );
  }

  return null;
}

function getSupportFaqReply(messageNorm) {
  const source = getSupportFaqSource();
  const ranked = source.map(function (item) {
    return {
      item: item,
      score: scoreSupportFaq(messageNorm, item)
    };
  }).sort(function (a, b) {
    if (b.score !== a.score) return b.score - a.score;
    return Number(b.item.priority || 0) - Number(a.item.priority || 0);
  });

  if (ranked[0] && ranked[0].score >= 10) {
    return buildSupportReply(
      ranked[0].item.answer,
      ranked
        .filter(function (row) { return row.score >= 10 && row.item.question !== ranked[0].item.question; })
        .slice(0, 3)
        .map(function (row) { return row.item.question; })
    );
  }

  return null;
}

function getSupportOverviewReply() {
  return buildSupportReply(
    [
      'このアプリでは、現在次の機能をご利用いただけます。',
      '1. ホーム: スタンプQR読み取り、現在のスタンプカード確認、最新NEWS、メニュー一覧、公式サイト・SNSの確認',
      '2. ショップ: 商品確認、カート追加、ご注文の確定',
      '3. カレンダー: 月ごとの予定確認とイベント詳細の確認',
      '4. NEWS: 記事一覧の確認、カテゴリ絞り込み',
      '5. 画面上部の📢: NEWS、カレンダー、ショップ、ホームの更新情報を一覧確認',
      '6. マイページ: プロフィール、会員ID、通知設定、注文履歴、特典、パスコード、引き継ぎコードの確認',
      '7. ログイン・復元: 起動時パスコード、データ引き継ぎ・復元、パスコード再設定',
      '知りたい画面名や操作名をそのまま送っていただければ、手順を具体的にご案内します。'
    ],
    ['このアプリでできることを知りたい', '商品の注文方法を知りたい', 'スタンプの集め方を知りたい', 'プロフィールの登録方法を知りたい'],
    true,
    'app-overview'
  );
}

function getBuiltInSupportReply(messageNorm) {
  if (hasAllSupportKeywordGroups(messageNorm, [['初回', '初めて', 'はじめて', '最初', '始める'], ['登録', 'プロフィール', '開始', 'ようこそ']])) {
    return buildSupportReply(
      [
        '初回起動時は、まず「以前登録した方はこちら」と「はじめて登録する方はこちら」の選択画面が表示されます。',
        '今回が初めての方は、そのままプロフィール登録へ進み、お名前などを入力してください。',
        '以前登録したことがある方は、新規登録ではなく復元からお入りください。'
      ],
      ['初回起動時はどちらを選べばいいですか？', 'プロフィールの登録方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['以前登録した方', 'はじめて登録する方', 'どちら'])) {
    return buildSupportReply(
      [
        '最初の選択画面のご案内です。',
        '1. 以前登録したことがある方、再インストール後の方、機種変更後の方は「以前登録した方はこちら ↺」を選びます。',
        '2. 今回が初めての方は「はじめて登録する方はこちら」を選びます。',
        '3. 以前登録した方が新規登録へ進むと、別の会員IDが作られることがあります。'
      ],
      ['データの引き継ぎ・復元方法を知りたい', '会員登録が重複しないようにする方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['会員id', 'memberid', '会員番号'])) {
    return buildSupportReply(
      _profile
        ? `会員IDはマイページ上部に表示されます。現在の会員IDは${_profile.memberId || '未発行'}です。`
        : '会員IDはプロフィール登録後に発行され、マイページ上部に表示されます。現在はプロフィール未登録です。',
      ['プロフィールの登録方法を知りたい', 'プロフィールの変更方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['パスコード']) && hasAnySupportKeyword(messageNorm, ['毎回', '起動時', 'オン', 'オフ', '省略'])) {
    return buildSupportReply(
      [
        'ログイン時のパスコード設定です。',
        '1. パスコードを一度設定したあと、マイページの「ログイン時のパスコード」からオン・オフを切り替えられます。',
        `2. 現在は${isPasscodeLoginEnabled() ? 'オン' : 'オフ'}です。`,
        '3. オフにすると、次回からアプリ起動時のパスコード入力を省略できます。'
      ],
      ['起動時のパスコード設定について知りたい', 'パスコードの変更方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['アイコン', 'アバター', 'バナー画像', '背景画像']) && hasAnySupportKeyword(messageNorm, ['変更', '編集', '設定'])) {
    return buildSupportReply(
      'プロフィール編集画面で、アイコン画像とバナー背景画像を変更できます。画像を選ぶとアプリ内で圧縮され、保存後にプロフィールやバナーへ反映されます。',
      ['プロフィールの変更方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['プロフィール']) && hasAnySupportKeyword(messageNorm, ['登録', '登録方法', '会員登録']) && !hasAnySupportKeyword(messageNorm, ['変更', '編集'])) {
    return buildSupportReply(
      [
        'プロフィール登録の方法です。',
        '1. 初回起動時は「はじめて登録する方はこちら」を選びます。',
        '2. お名前、電話番号、生年月日、住所を入力して保存してください。',
        '3. すでに登録済みの方は新規登録ではなく「以前登録した方はこちら ↺」から復元してください。',
        '補足: お名前は必須です。'
      ],
      ['初回起動時はどちらを選べばいいですか？', 'データの引き継ぎ・復元方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['プロフィール']) && hasAnySupportKeyword(messageNorm, ['登録', '変更', '編集', '名前', '電話', '生年月日', '住所'])) {
    return buildSupportReply(
      [
        'プロフィール変更の方法です。',
        '1. マイページを開きます。',
        '2. 「プロフィールを編集」を押します。',
        '3. お名前、電話番号、生年月日、住所を入力または修正して保存します。',
        '補足: 保存後はマイページやバナー表示に反映されます。'
      ],
      ['プロフィールの登録方法を知りたい', 'プロフィールの変更方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['通知', 'push', 'プッシュ']) && hasAnySupportKeyword(messageNorm, ['オン', 'オフ', '許可', '設定'])) {
    return buildSupportReply(
      [
        '通知設定の方法です。',
        '1. マイページを開きます。',
        '2. 「通知設定」ボタンを押して、通知をオンまたはオフに切り替えます。',
        '3. 端末側で拒否されている場合は、iPhoneやブラウザの通知設定もご確認ください。',
        '補足: 通知内容はブログ更新や重要なお知らせです。'
      ],
      ['通知のオンオフ方法を知りたい', '最新のお知らせの見方を知りたい']
    );
  }

  if (isSupportBlogArchiveQuestion(messageNorm)) {
    return buildSupportReply(
      [
        '過去のブログやお知らせの確認方法です。',
        '1. 画面上部の📢ボタンを押します。',
        '2. 「お知らせ一覧」で NEWS、カレンダー、ショップ、ホームの更新情報を新しい順に確認できます。',
        '3. NEWSページでは記事一覧を、📢のお知らせ一覧では更新情報の一覧を見られます。'
      ],
      ['ブログや過去のお知らせの見方を知りたい', 'お知らせ一覧の見方を知りたい']
    );
  }

  if (isSupportOrderHistoryQuestion(messageNorm)) {
    return buildSupportReply(
      [
        '注文履歴の確認方法です。',
        '1. 下部メニューからマイページを開きます。',
        '2. 「ご注文履歴」を開きます。',
        '3. 受付中の注文や、受け取り前の注文を確認できます。',
        '補足: プロフィール未登録の場合は履歴は表示されません。'
      ],
      ['商品の注文方法を知りたい', '注文履歴の確認方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['キャンセル']) && hasAnySupportKeyword(messageNorm, ['注文', '履歴'])) {
    return buildSupportReply(
      [
        '注文キャンセルの方法です。',
        '1. マイページの「ご注文履歴」を開きます。',
        '2. 受付中の注文に表示される「キャンセルする」を押します。',
        '3. 確認後、キャンセルした注文は履歴から表示されなくなります。'
      ],
      ['注文履歴の確認方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['受け取りました']) || (hasAnySupportKeyword(messageNorm, ['受取', '受け取り']) && hasAnySupportKeyword(messageNorm, ['注文', '履歴', '商品']))) {
    return buildSupportReply(
      [
        '受け取り完了の報告方法です。',
        '1. マイページの「ご注文履歴」を開きます。',
        '2. 対象注文の「受け取りました」を押します。',
        '3. 受取完了として更新されます。',
        '補足: キャンセル済みの注文には表示されません。'
      ],
      ['注文履歴の確認方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['支払い', '決済', '現金'])) {
    return buildSupportReply(
      [
        '現在の支払い方法についてご案内します。',
        '商品注文の支払い方法は現在「現金払い」です。',
        '注文確定後、ご来院時にスタッフへお声がけください。'
      ],
      ['商品の注文方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['注文', '購入']) || hasAllSupportKeywordGroups(messageNorm, [['ショップ', '商品'], ['方法', 'やり方', '買い方', '流れ']])) {
    return buildSupportReply(
      [
        '商品の注文手順です。',
        '1. 下部メニューの「ショップ」を開きます。',
        '2. 商品カードをタップして詳細を開きます。',
        '3. 個数を選んで「カートに追加する」を押します。',
        '4. カートで内容を確認し、「ご注文を確定する」を押します。',
        '5. 注文後はマイページの「ご注文履歴」で状態を確認できます。'
      ],
      ['カートの使い方を知りたい', '注文履歴の確認方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['カート', '買い物かご']) && hasAnySupportKeyword(messageNorm, ['追加', '入れる', '個数', '削除', '変更', '見る'])) {
    return buildSupportReply(
      [
        'カートの使い方です。',
        '1. 商品をカートに追加します。',
        '2. カート画面で個数変更や削除ができます。',
        '3. 商品がない場合は空画面になり、「ショップへ」から戻れます。'
      ],
      ['カートの使い方を知りたい', '商品の注文方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['ショップ', '商品'])) {
    return buildSupportReply(
      [
        'ショップの見方です。',
        '1. 下部メニューの「ショップ」を開きます。',
        '2. 商品カードをタップすると詳細が開きます。',
        '3. 「商品説明をみる」で説明確認、個数選択、「カートに追加する」で購入準備ができます。'
      ],
      ['商品の注文方法を知りたい', 'カートの使い方を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['新しいカード', '次のカード', 'カードを取得', 'カード開始'])) {
    return buildSupportReply(
      [
        '新しいスタンプカードの始め方です。',
        '1. スタンプが10個たまると、ホームに「特典ガチャを回す」ボタンが表示されます。',
        '2. 特典ガチャを回すと、「新しいスタンプカードを取得」ボタンが表示されます。',
        '3. そのボタンを押すと次のスタンプカードが始まります。'
      ],
      ['スタンプ特典の使い方を知りたい', 'スタンプの集め方を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['特典', 'プレゼント', '期限', '使用', 'チケット'])) {
    return buildSupportReply(
      [
        'スタンプ特典の使い方です。',
        '1. マイページの「スタンプ・特典履歴」を開きます。',
        '2. 獲得済み特典を確認し、必要な場合は「使用する」を押します。',
        '3. 特典は達成当日から使用できます。',
        '4. 受取期限は獲得から1か月です。',
        '補足: 一度使用すると再度は使用できません。受け取りの際は受付へ直接お問い合わせください'
      ],
      ['スタンプ特典の使い方を知りたい', 'スタンプの集め方を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['1日', '一日', '何回']) && hasAnySupportKeyword(messageNorm, ['スタンプ'])) {
    return buildSupportReply(
      '来院スタンプは1日1回までです。同じ日に再度読み取ると「本日はすでにスタンプを取得済みです」と表示されます。',
      ['スタンプの集め方を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['qr', 'qrコード', 'qrcode', 'カメラ', 'スキャン', '読み取り'])) {
    return buildSupportReply(
      [
        'スタンプ取得の手順です。',
        '1. ホーム画面の「カメラを起動して読み取る」を押します。',
        '2. カメラ許可の確認が出た場合は「許可」を選びます。',
        '3. 院内のQRコードを読み取ります。',
        '4. 読み取りが成功するとスタンプが追加されます。',
        '補足: すでに拒否している場合は「設定を開く」から許可してください。プロフィール未登録の場合は先に登録してください。'
      ],
      ['スタンプの集め方を知りたい', 'スタンプ特典の使い方を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['スタンプ'])) {
    return buildSupportReply(
      [
        'スタンプ機能のご案内です。',
        '1. スタンプはホーム画面で管理します。',
        '2. 院内QRコードを読むと1日1回まで追加されます。',
        '3. 10個たまると特典ガチャを回せます。',
        '4. ガチャ結果はマイページで確認できます。'
      ],
      ['スタンプの集め方を知りたい', 'スタンプ特典の使い方を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['予約', 'よやく', 'アポイント', '予約したい', '予約方法'])) {
    return buildSupportReply(
      [
        'ご予約についてのご案内です。',
        '当院のご予約は「公式LINE」にて承っております。',
        'お手数ですが、公式LINEの友達登録をしていただき、トーク画面から予約をご依頼ください。',
        '公式LINEへは、ホーム画面またはマイページの「公式サイト・SNS」からアクセスできます。'
      ],
      ['公式LINEやSNSの開き方を知りたい', 'メニュー一覧の見方を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['メニュー', '施術'])) {
    return buildSupportReply(
      [
        'メニュー一覧の見方です。',
        '1. ホームの「メニュー一覧を見る🍴」を押します。',
        '2. 右上のカテゴリ選択で絞り込みできます。',
        '3. 各メニューをタップすると詳細や画像を確認できます。',
        '補足: 実際のご予約は公式LINEからお願いいたします。'
      ],
      ['メニュー一覧の見方を知りたい', '予約の方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['カレンダー', 'イベント', '予定', '日程'])) {
    return buildSupportReply(
      [
        'カレンダーの見方です。',
        '1. 下部メニューの「カレンダー」を開きます。',
        '2. 左右の矢印で月を切り替えます。',
        '3. 日付をタップするとその日の予定を確認できます。',
        '4. 下部には今月のイベント一覧も表示されます。'
      ],
      ['イベントカレンダーの見方を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['お知らせ一覧', '通知一覧', '拡声器'])) {
    return buildSupportReply(
      [
        'お知らせ一覧の見方です。',
        '1. 画面上部の📢ボタンを押します。',
        '2. NEWS、カレンダー、ショップ、ホームの更新情報を新しい順に確認できます。',
        '3. カテゴリ選択で絞り込みできます。',
        '補足: 新着があると📢に赤いドットが表示されます。'
      ],
      ['お知らせ一覧の見方を知りたい', '最新のお知らせの見方を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['news', 'ニュース', 'お知らせ'])) {
    return buildSupportReply(
      [
        'NEWSページの見方です。',
        '1. 下部メニューの「NEWS」を開きます。',
        '2. 記事を一覧で確認できます。',
        '3. 記事をタップすると詳細が開きます。',
        '4. 右上のカテゴリ選択で絞り込みできます。',
        '補足: 画面上部の📢は更新情報一覧、NEWSは記事一覧です。'
      ],
      ['NEWSページの使い方を知りたい', 'お知らせ一覧の見方を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['更新', '再読み込み', 'リロード', '最新情報'])) {
    return buildSupportReply(
      [
        '最新情報の更新方法です。',
        '1. 画面上部の🔄ボタンを押します。',
        '2. ブログ、お知らせ、カレンダー、商品、メニュー一覧、FAQ、注文履歴などの最新情報を再取得します。',
        '3. 再インストールせず、その場で最新状態に反映します。'
      ],
      ['最新情報への更新方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['line', 'ライン', 'instagram', 'facebook', 'ホームページ', '公式サイト', 'sns', '問い合わせ', 'お問い合わせ'])) {
    return buildSupportReply(
      [
        '公式サイトやSNSの開き方です。',
        '1. ホーム画面またはマイページの「公式サイト・SNS」を開きます。',
        '2. 公式ホームページ、Instagram、Facebook、公式LINEを選んで開けます。',
        '補足: 診療や個別相談は公式LINEをご利用ください。'
      ],
      ['公式LINEやSNSの開き方を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['ホーム', 'トップ'])) {
    return buildSupportReply(
      [
        'ホーム画面でできることです。',
        '1. スタンプQRの読み取り',
        '2. 現在のスタンプカード確認',
        '3. メニュー一覧への移動',
        '4. 最新のお知らせ確認',
        '5. 公式サイトやSNSへのアクセス'
      ],
      ['ホーム画面の見方を知りたい', 'スタンプの集め方を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['アプリ', '使い方', 'できること', '機能', '何ができる'])) {
    return getSupportOverviewReply();
  }

  // --- トラブルシューティング ---
  if (hasAnySupportKeyword(messageNorm, ['表示されない', 'おかしい', '崩れ', '不具合', 'バグ'])) {
    return buildSupportReply(
      [
        '画面表示のトラブルについてです。',
        '以下をお試しください：',
        '1. 画面右上の🔄ボタンで情報を再取得',
        '2. アプリを一度閉じて再起動',
        '3. インターネット接続を確認'
      ],
      ['最新情報への更新方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['エラー', '失敗', '同期', '問題'])) {
    return buildSupportReply(
      [
        'データ取得エラーについてです。',
        'インターネット接続が不安定な可能性があります。',
        'Wi-Fiや4G/5G回線が安定している環境で🔄ボタンを押して再度お試しください。',
        '一部データの取得に失敗しても、他のデータは正常に更新されます。'
      ],
      ['最新情報への更新方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['重い', '遅い', 'フリーズ', '動かない'])) {
    return buildSupportReply(
      [
        'アプリの動作が遅い場合の対処法です。',
        '1. 他のアプリを閉じてメモリを開放してください。',
        '2. 端末を再起動してみてください。',
        '3. 🔄ボタンでアプリを最新状態に更新してください。'
      ],
      ['最新情報への更新方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['引き継ぎ', '機種変更', 'データ移行'])) {
    return buildSupportReply(
      [
        '端末の引き継ぎについてです。',
        'ログイン画面、または初回画面の「以前登録した方はこちら ↺」から進めます。',
        '引き継ぎコードがある場合は、引き継ぎコードと新しいパスコードを入力してください。',
        '引き継ぎコードがない場合は、お名前に加えて電話番号・生年月日・現在のパスコードのうち1つ以上を入力すると復元できます。',
        '機種変更前は、マイページの「↺ 引き継ぎコードの発行」からコードを出しておくとスムーズです。'
      ],
      ['引き継ぎコードの発行方法を知りたい', 'パスコードを忘れたときの再設定方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['削除', 'アンインストール', '復元', '再インストール'])) {
    return buildSupportReply(
      [
        'アプリの削除・復元についてです。',
        '再インストール後は、ログイン画面や初回画面で「以前登録した方はこちら ↺」を選ぶと会員情報を戻せます。',
        '引き継ぎコードがある場合はコードで復元できます。ない場合も、お名前と電話番号・生年月日・現在のパスコードのうち1つ以上で復元できます。',
        '再インストール前にマイページで引き継ぎコードを発行しておくと、よりスムーズです。'
      ],
      ['データの引き継ぎ・復元方法を知りたい', '引き継ぎコードの発行方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['個人情報', 'プライバシー', 'セキュリティ'])) {
    return buildSupportReply(
      [
        '個人情報の取り扱いについてです。',
        'プロフィールに登録いただいた情報は、まゆみ助産院の業務（注文管理・会員管理など）のためにのみ使用されます。',
        '第三者への提供は行いません。'
      ],
      ['プロフィールの登録方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['対応端末', 'iphone', 'android', 'ブラウザ', 'safari'])) {
    return buildSupportReply(
      [
        'アプリの対応端末についてです。',
        'iPhone と Android のどちらでもご利用いただけます。',
        'ホーム画面へ追加する場合は、iPhone は Safari、Android は Chrome から行ってください。',
        'すでに会員登録済みの方は、ホーム画面追加後に新規登録せず「以前登録した方はこちら」からお入りください。'
      ],
      ['ホーム画面への追加方法を知りたい', '会員登録が重複しないようにする方法を知りたい']
    );
  }

  if (hasAnySupportKeyword(messageNorm, ['チャットボット', 'ボット', 'サポート']) && hasAnySupportKeyword(messageNorm, ['何', 'できる', '使い方', '開き方'])) {
    return buildSupportReply(
      [
        'チャットサポートについてです。',
        'アプリの使い方に関する質問にチャット形式で回答するサポート機能です。',
        '注文方法、スタンプの集め方、通知設定、プロフィール登録などの操作方法を案内します。',
        '画面右下の「💬 使い方相談」ボタン、またはマイページの「🤖 使い方サポート」からご利用いただけます。'
      ],
      ['このアプリでできることを知りたい']
    );
  }

  return null;
}

async function loadSupportFaq(force) {
  if (!force && supportFaqItems.length) {
    renderSupportSuggestions(getDefaultSupportQuestions(4));
    return supportFaqItems;
  }

  const res = await getFromGAS('getSupportFaq');
  if (res && res.status === 'ok' && Array.isArray(res.faqs) && res.faqs.length) {
    supportFaqItems = res.faqs.map(function (item) {
      return {
        rowIdx: item && item.rowIdx,
        category: String(item && item.category || ''),
        question: String(item && item.question || ''),
        keywords: String(item && item.keywords || ''),
        answer: String(item && item.answer || ''),
        priority: Number(item && item.priority || 0),
        updatedAt: String(item && item.updatedAt || '')
      };
    });
  } else {
    supportFaqItems = SUPPORT_FAQ_FALLBACK.slice();
  }

  if (document.getElementById('page-notices').classList.contains('active')) {
    renderPushNotices();
  }
  renderSupportSuggestions(getDefaultSupportQuestions(4));
  return supportFaqItems;
}

function getLocalSupportReply(message) {
  const messageNorm = normalizeSupportText(message);
  const statusReply = getSupportStatusReply(messageNorm);
  const featureReply = getFeatureSupportReply(messageNorm);
  const builtInReply = getBuiltInSupportReply(messageNorm);
  const faqReply = getSupportFaqReply(messageNorm);

  if (statusReply) {
    return statusReply;
  }
  if (featureReply) return featureReply;
  if (builtInReply) return builtInReply;
  if (faqReply) return faqReply;

  if (isLikelyAppSupportQuestion(messageNorm)) {
    return getSupportOverviewReply();
  }

  return buildSupportReply(
    [
      'このチャットでは、現在アプリに実装されている機能の使い方をご案内しています。',
      'ご案内できる主な内容: ホーム、予約、スタンプ、メニュー一覧、ショップ、カレンダー、NEWS、通知設定、プロフィール、注文履歴、特典。',
      '知りたい画面名や操作名をそのまま送ってください。例: 「予約の方法」「注文履歴の見方」「通知をオンにする方法」',
      '診療や個別相談は公式LINEをご利用ください。'
    ],
    ['予約の方法を知りたい', 'このアプリでできることを知りたい', '商品の注文方法を知りたい', 'スタンプの集め方を知りたい'],
    false
  );
}

function persistSupportChatHistory() {
  const clean = supportChatHistory
    .filter(function (item) { return !item.pending; })
    .slice(-20)
    .map(function (item) {
      return { role: item.role, text: item.text };
    });
  supportChatHistory = clean;
  try {
    localStorage.setItem('mayumi_support_chat_history', JSON.stringify(clean));
  } catch (e) { }
}

function persistSupportTopic() {
  try {
    if (lastSupportTopic) {
      localStorage.setItem('mayumi_support_chat_topic', lastSupportTopic);
    } else {
      localStorage.removeItem('mayumi_support_chat_topic');
    }
  } catch (e) { }
}

function ensureSupportChatWelcome() {
  if (supportChatHistory.length) return;
  lastSupportTopic = '';
  persistSupportTopic();
  supportChatHistory = [{
    role: 'bot',
    text: 'アプリの使い方をご案内します。機能ごとに、どこから使うか・何ができるか・できないこと・困ったときの対処まで、できるだけ具体的にお答えします。ホーム、予約、スタンプ、メニュー一覧、ショップ、カレンダー、NEWS、通知設定、プロフィール、注文履歴などをそのまま質問してください。'
  }];
  persistSupportChatHistory();
}

function updateSupportChatSendingState() {
  const input = document.getElementById('supportChatInput');
  const sendBtn = document.getElementById('supportChatSendBtn');
  if (input) input.disabled = isSupportChatSending;
  if (sendBtn) {
    sendBtn.disabled = isSupportChatSending;
    sendBtn.textContent = isSupportChatSending ? '送信中...' : '送信';
  }

  document.querySelectorAll('#supportChatSuggestions .support-chip').forEach(function (chip) {
    chip.disabled = isSupportChatSending;
  });
}

function renderSupportSuggestions(questions) {
  const container = document.getElementById('supportChatSuggestions');
  if (!container) return;

  const sourceQuestions = (questions && questions.length ? questions : getDefaultSupportQuestions(4))
    .filter(Boolean);
  const unique = Array.from(new Set(sourceQuestions)).slice(0, 4);
  container.innerHTML = '';

  unique.forEach(function (question) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'support-chip';
    btn.textContent = question;
    btn.disabled = isSupportChatSending;
    btn.addEventListener('click', function () {
      askSuggestedSupportQuestion(question);
    });
    container.appendChild(btn);
  });
}

function renderSupportChatMessages() {
  const container = document.getElementById('supportChatMessages');
  if (!container) return;

  container.innerHTML = '';
  supportChatHistory.forEach(function (message) {
    const bubble = document.createElement('div');
    bubble.className = 'support-msg ' + (message.role === 'user' ? 'user' : 'bot');
    if (message.pending) bubble.className += ' pending';

    const label = document.createElement('div');
    label.className = 'support-msg-label';
    label.textContent = message.role === 'user' ? 'あなた' : 'サポート';

    const body = document.createElement('div');
    body.className = 'support-msg-body';
    if (message.pending) {
      body.innerHTML = '<span class="support-loading"><span class="support-loading-dot"></span><span class="support-loading-dot"></span><span class="support-loading-dot"></span><span>回答を準備しています...</span></span>';
    } else {
      body.innerHTML = escapeHtml(message.text).replace(/\n/g, '<br>');
    }

    bubble.appendChild(label);
    bubble.appendChild(body);
    container.appendChild(bubble);
  });

  container.scrollTop = container.scrollHeight;
}

async function openSupportChat() {
  ensureSupportChatWelcome();
  openModal('supportChatModal');
  renderSupportChatMessages();
  renderSupportSuggestions();
  updateSupportChatSendingState();
  await loadSupportFaq();

  setTimeout(function () {
    const input = document.getElementById('supportChatInput');
    if (input) input.focus();
  }, 120);
}

function askSuggestedSupportQuestion(question) {
  if (isSupportChatSending) return;
  const input = document.getElementById('supportChatInput');
  if (input) input.value = question;
  sendSupportChat(question);
}


async function sendSupportChat(prefilledText) {
  if (isSupportChatSending) return;

  const input = document.getElementById('supportChatInput');
  const message = String(typeof prefilledText === 'string' ? prefilledText : (input ? input.value : '')).trim();
  if (!message) return;

  ensureSupportChatWelcome();
  isSupportChatSending = true;
  updateSupportChatSendingState();
  supportChatHistory.push({ role: 'user', text: message });
  if (input) input.value = '';

  const pendingMessage = { role: 'bot', text: '', pending: true };
  supportChatHistory.push(pendingMessage);
  renderSupportSuggestions([]);
  renderSupportChatMessages();

  try {
    const startedAt = Date.now();
    if (!supportFaqItems.length) {
      await loadSupportFaq();
    }

    const localReply = getLocalSupportReply(message);
    const elapsed = Date.now() - startedAt;
    if (elapsed < 250) {
      await new Promise(function (resolve) { setTimeout(resolve, 250 - elapsed); });
    }

    supportChatHistory = supportChatHistory.filter(function (item) { return item !== pendingMessage; });
    supportChatHistory.push({ role: 'bot', text: localReply.answer });
    lastSupportTopic = localReply.topic || resolveSupportTopic(normalizeSupportText(message)) || lastSupportTopic;
    renderSupportSuggestions(localReply.suggestions || []);
    persistSupportChatHistory();
    persistSupportTopic();
    renderSupportChatMessages();
  } catch (err) {
    supportChatHistory = supportChatHistory.filter(function (item) { return item !== pendingMessage; });
    supportChatHistory.push({
      role: 'bot',
      text: '回答の準備中にエラーが発生しました。もう一度お試しいただくか、知りたい操作名を短く送ってください。'
    });
    renderSupportSuggestions(getDefaultSupportQuestions(4));
    persistSupportChatHistory();
    renderSupportChatMessages();
  } finally {
    isSupportChatSending = false;
    updateSupportChatSendingState();
  }
}


// ===== プロフィール管理 =====
// デフォルトアバター画像
const DEFAULT_AVATAR = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSI4MCIgaGVpZ2h0PSI4MCIgdmlld0JveD0iMCAwIDgwIDgwIj48cmVjdCB3aWR0aD0iODAiIGhlaWdodD0iODAiIHJ4PSI0MCIgZmlsbD0iI2I1YzlhOCIvPjx0ZXh0IHg9IjQwIiB5PSI0OCIgdGV4dC1hbmNob3I9Im1pZGRsZSIgZmlsbD0id2hpdGUiIGZvbnQtc2l6ZT0iMzIiPvCfkak8L3RleHQ+PC9zdmc+';
let _avatarData = null;
try { _avatarData = localStorage.getItem('mayumi_avatar') || null; } catch (e) { }
let _bannerData = null;
try { _bannerData = localStorage.getItem('mayumi_banner') || null; } catch (e) { }

// 画像リサイズ・圧縮ヘルパー
function compressImage(file, maxWidth, maxHeight, quality, callback) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const img = new Image();
    img.onload = function () {
      const canvas = document.createElement('canvas');
      let width = img.width;
      let height = img.height;
      if (width > height) {
        if (width > maxWidth) {
          height *= maxWidth / width;
          width = maxWidth;
        }
      } else {
        if (height > maxHeight) {
          width *= maxHeight / height;
          height = maxHeight;
        }
      }
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, width, height);
      callback(canvas.toDataURL('image/jpeg', quality));
    };
    img.onerror = () => callback(null);
    img.src = e.target.result;
  };
  reader.onerror = () => callback(null);
  reader.readAsDataURL(file);
}

function handleAvatarChange(input) {
  const file = input.files[0];
  if (!file) return;
  showToast('画像を処理中...');
  compressImage(file, 800, 800, 0.7, (dataUrl) => {
    if (!dataUrl) {
      showToast('画像の読み込みに失敗しました');
      return;
    }
    _avatarData = dataUrl;
    localStorage.removeItem('mayumi_avatar_url');
    // プレビュー更新
    const preview = document.getElementById('avatarPreview');
    if (preview) preview.src = _avatarData;
  });
}

function handleBannerChange(input) {
  const file = input.files[0];
  if (!file) return;
  showToast('画像を処理中...');
  compressImage(file, 1200, 1000, 0.7, (dataUrl) => {
    if (!dataUrl) {
      showToast('画像の読み込みに失敗しました');
      return;
    }
    _bannerData = dataUrl;
    // プレビュー更新
    const preview = document.getElementById('bannerPreview');
    if (preview) preview.style.backgroundImage = `url(${_bannerData})`;
  });
}

function clearPasscodeOverlayError() {
  const input = document.getElementById('lockPasscode');
  const error = document.getElementById('lockPasscodeErr');
  if (input) input.classList.remove('error');
  if (error) error.classList.remove('show');
}

function prefillForgotPasscodeForm(fieldMap) {
  const settings = fieldMap || {};
  const nameInput = document.getElementById(settings.nameId);
  const phoneInput = document.getElementById(settings.phoneId);
  const birthdayInput = document.getElementById(settings.birthdayId);
  const passcodeInput = document.getElementById(settings.passcodeId);
  const sourcePhoneInput = settings.sourcePhoneId ? document.getElementById(settings.sourcePhoneId) : null;
  const profileName = _profile && _profile.name ? _profile.name : '';
  const profilePhone = _profile && _profile.phone ? _profile.phone : '';
  const profileBirthday = _profile && _profile.birthday ? _profile.birthday : '';

  if (nameInput) {
    nameInput.value = normalizeNameInput(nameInput.value) || profileName;
  }
  if (phoneInput) {
    const currentPhone = normalizePhoneInput(phoneInput.value);
    const sourcePhone = sourcePhoneInput ? normalizePhoneInput(sourcePhoneInput.value) : '';
    phoneInput.value = currentPhone || sourcePhone || profilePhone;
  }
  if (birthdayInput) {
    birthdayInput.value = normalizeDateOnlyInput(birthdayInput.value) || profileBirthday;
  }
  if (passcodeInput) {
    passcodeInput.value = '';
  }
}

function prefillRestoreAccountForm(fieldMap) {
  const settings = fieldMap || {};
  const nameInput = document.getElementById(settings.nameId);
  const phoneInput = document.getElementById(settings.phoneId);
  const birthdayInput = document.getElementById(settings.birthdayId);
  const passcodeInput = document.getElementById(settings.passcodeId);
  const transferCodeInput = document.getElementById(settings.transferCodeId);
  const newPasscodeInput = document.getElementById(settings.newPasscodeId);
  const profileName = _profile && _profile.name ? _profile.name : '';
  const profilePhone = _profile && _profile.phone ? _profile.phone : '';
  const profileBirthday = _profile && _profile.birthday ? _profile.birthday : '';

  if (nameInput) nameInput.value = normalizeNameInput(nameInput.value) || profileName;
  if (phoneInput) phoneInput.value = normalizePhoneInput(phoneInput.value) || profilePhone;
  if (birthdayInput) birthdayInput.value = normalizeDateOnlyInput(birthdayInput.value) || profileBirthday;
  if (passcodeInput) passcodeInput.value = '';
  if (transferCodeInput) transferCodeInput.value = '';
  if (newPasscodeInput) newPasscodeInput.value = '';
}

function updateTransferCodeDisplay(details) {
  const card = document.getElementById('transferCodeResultCard');
  const codeEl = document.getElementById('transferCodeValue');
  const expiryEl = document.getElementById('transferCodeExpiry');
  const code = details && details.transferCode ? String(details.transferCode) : '';
  const expiryText = details && details.expiresAt ? formatCustomerDateYmdHm(details.expiresAt) : (details && details.expiresAtLabel ? details.expiresAtLabel : '');

  if (codeEl) codeEl.textContent = code || '--------';
  if (expiryEl) expiryEl.textContent = expiryText ? ('有効期限：' + expiryText) : '有効期限：---';
  if (card) card.style.display = code ? 'block' : 'none';
}

async function applyRecoveredUserState(user, passcode, successMessage) {
  const u = user || {};
  let recoveredRewards = [];
  let recoveredStampHistory = [];
  try {
    if (Array.isArray(u.rewards)) {
      recoveredRewards = normalizeRewardList(u.rewards);
    } else {
      recoveredRewards = normalizeRewardList(JSON.parse(String(u.rewards || '[]')));
    }
  } catch (e) {
    recoveredRewards = [];
  }
  try {
    if (Array.isArray(u.stampHistory)) {
      recoveredStampHistory = normalizeStampHistoryList(u.stampHistory);
    } else {
      recoveredStampHistory = normalizeStampHistoryList(JSON.parse(String(u.stampHistory || '[]')));
    }
  } catch (e) {
    recoveredStampHistory = [];
  }
  const profile = {
    memberId: u.memberId,
    name: u.name,
    kana: u.kana,
    phone: u.phone,
    avatar: u.avatar,
    memo: u.memo,
    status: u.status,
    birthday: u.birthday,
    address: u.address,
    regDate: u.regDate
  };

  _profile = profile;
  CUSTOMER_NAME = profile && profile.name ? profile.name : '';
  stampCount = Number(u.stampCount || 0) || 0;
  stampCardNum = Number(u.stampCardNum || 1) || 1;
  EARNED_REWARDS = recoveredRewards;
  STAMP_HISTORY = recoveredStampHistory;
  const recoveredLastStampAt = deriveLastStampAt({
    lastStampAt: u.lastStampAt,
    lastStampDate: u.lastStampDate,
    stampHistory: recoveredStampHistory
  }, recoveredStampHistory);
  localStorage.setItem('mayumi_profile', JSON.stringify(profile));
  localStorage.setItem('mayumi_stamp', String(stampCount));
  localStorage.setItem('mayumi_stamp_card', String(stampCardNum));
  localStorage.setItem('mayumi_earned_rewards', JSON.stringify(EARNED_REWARDS));
  localStorage.setItem(STAMP_HISTORY_STORAGE_KEY, JSON.stringify(STAMP_HISTORY));
  if (u.lastStampDate) localStorage.setItem('mayumi_last_stamp_date', u.lastStampDate);
  else localStorage.removeItem('mayumi_last_stamp_date');
  if (recoveredLastStampAt) localStorage.setItem(LAST_STAMP_AT_STORAGE_KEY, recoveredLastStampAt);
  else localStorage.removeItem(LAST_STAMP_AT_STORAGE_KEY);
  if (u.stampAchievedAt) localStorage.setItem('mayumi_stamp_10_date', u.stampAchievedAt);
  else localStorage.removeItem('mayumi_stamp_10_date');
  await storeLocalPasscode(passcode);
  isPasscodeAuthenticated = true;
  markPasscodeUnlockSkippedOnce();

  try { closePasscodeOverlay(); } catch (e) { }
  try { closeSetupModal(); } catch (e) { }
  const onboarding = document.getElementById('onboardingScreen');
  const migrationOverlay = document.getElementById('migrationOverlay');
  if (onboarding) onboarding.classList.remove('show');
  if (migrationOverlay) migrationOverlay.classList.remove('open');
  updateProfileUI();
  updateStampUI();
  renderEarnedRewards();
  cacheDeviceSessions(Array.isArray(u.deviceSessions) ? u.deviceSessions : []);
  syncCurrentDeviceSession({ silent: true }).catch(function (error) {
    console.error('syncCurrentDeviceSession after recovery error:', error);
  });
  const homePage = document.getElementById('page-home');
  if (homePage) homePage.classList.add('active');
  switchPage('home');
  fetchLatestManagedContent({
    refreshSupportFaq: false,
    refreshOrderHistory: true
  }).catch(function (e) {
    console.error('recovery refresh error:', e);
  });
  showToast(successMessage || 'ログインしました🌿');
}

function togglePasscodeOverlayView(view) {
  const login = document.getElementById('passcodeLoginContent');
  const restore = document.getElementById('passcodeRestoreContent');
  const reset = document.getElementById('passcodeResetContent');
  if (login) login.style.display = view === 'login' ? 'block' : 'none';
  if (restore) restore.style.display = view === 'restore' ? 'block' : 'none';
  if (reset) reset.style.display = view === 'reset' ? 'block' : 'none';
  if (view === 'restore') {
    prefillRestoreAccountForm({
      nameId: 'passcodeRestoreName',
      phoneId: 'passcodeRestorePhone',
      birthdayId: 'passcodeRestoreBirthday',
      passcodeId: 'passcodeRestorePasscode',
      transferCodeId: 'passcodeRestoreTransferCode',
      newPasscodeId: 'passcodeRestoreNewPasscode'
    });
  }
  if (view === 'reset') {
    prefillForgotPasscodeForm({
      nameId: 'passcodeResetName',
      phoneId: 'passcodeResetPhone',
      birthdayId: 'passcodeResetBirthday',
      passcodeId: 'passcodeResetNewPasscode',
      sourcePhoneId: 'passcodeRestorePhone'
    });
  }
  clearPasscodeOverlayError();
}

function openPasscodeOverlay() {
  if (!_profile || !hasConfiguredLocalPasscode()) return;
  const overlay = document.getElementById('passcodeOverlay');
  const input = document.getElementById('lockPasscode');
  const resetName = document.getElementById('passcodeResetName');
  const resetPhone = document.getElementById('passcodeResetPhone');
  const resetBirthday = document.getElementById('passcodeResetBirthday');
  const resetPasscode = document.getElementById('passcodeResetNewPasscode');
  if (input) input.value = '';
  prefillRestoreAccountForm({
    nameId: 'passcodeRestoreName',
    phoneId: 'passcodeRestorePhone',
    birthdayId: 'passcodeRestoreBirthday',
    passcodeId: 'passcodeRestorePasscode',
    transferCodeId: 'passcodeRestoreTransferCode',
    newPasscodeId: 'passcodeRestoreNewPasscode'
  });
  if (resetName) resetName.value = _profile.name || '';
  if (resetPhone) resetPhone.value = _profile.phone || '';
  if (resetBirthday) resetBirthday.value = _profile.birthday || '';
  if (resetPasscode) resetPasscode.value = '';
  togglePasscodeOverlayView('login');
  if (overlay) overlay.classList.add('open');
  setTimeout(function () {
    if (input) input.focus();
  }, 60);
}

function closePasscodeOverlay() {
  const overlay = document.getElementById('passcodeOverlay');
  if (overlay) overlay.classList.remove('open');
  clearPasscodeOverlayError();
}

function shouldRequirePasscodeLock() {
  if (!_profile || !hasConfiguredLocalPasscode()) return false;
  if (!isPasscodeLoginEnabled()) return false;
  const onboarding = document.getElementById('onboardingScreen');
  const setupOverlay = document.getElementById('setupOverlay');
  const migrationOverlay = document.getElementById('migrationOverlay');
  if (onboarding && onboarding.classList.contains('show')) return false;
  if (setupOverlay && setupOverlay.classList.contains('open')) return false;
  if (migrationOverlay && migrationOverlay.classList.contains('open')) return false;
  return true;
}

async function unlockAppWithPasscode() {
  const input = document.getElementById('lockPasscode');
  const error = document.getElementById('lockPasscodeErr');
  const btn = document.getElementById('unlockBtn');
  const passcode = normalizePasscodeInput(input ? input.value : '');

  if (!isValidPasscodeValue(passcode)) {
    if (input) input.classList.add('error');
    if (error) error.classList.add('show');
    showToast('4桁または6桁の数字で入力してください');
    return;
  }

  if (btn) {
    btn.disabled = true;
    btn.textContent = '確認中...';
  }
  clearPasscodeOverlayError();

  try {
    const isMatched = await verifyLocalPasscode(passcode);
    if (!isMatched) {
      if (input) input.classList.add('error');
      if (error) error.classList.add('show');
      showToast('パスコードが一致しません');
      return;
    }
    isPasscodeAuthenticated = true;
    closePasscodeOverlay();
    showToast('ログインしました🌿');
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = 'ログインする';
    }
  }
}

async function restoreAccountByForm(options) {
  const settings = options || {};
  const nameInput = document.getElementById(settings.nameId);
  const phoneInput = document.getElementById(settings.phoneId);
  const birthdayInput = document.getElementById(settings.birthdayId);
  const passcodeInput = document.getElementById(settings.passcodeId);
  const transferCodeInput = document.getElementById(settings.transferCodeId);
  const newPasscodeInput = document.getElementById(settings.newPasscodeId);
  const btn = document.getElementById(settings.buttonId);
  const name = normalizeNameInput(nameInput ? nameInput.value : '');
  const phone = normalizePhoneInput(phoneInput ? phoneInput.value : '');
  const birthday = normalizeDateOnlyInput(birthdayInput ? birthdayInput.value : '');
  const passcode = normalizePasscodeInput(passcodeInput ? passcodeInput.value : '');
  const transferCode = normalizeTransferCodeInput(transferCodeInput ? transferCodeInput.value : '');
  const newPasscode = normalizePasscodeInput(newPasscodeInput ? newPasscodeInput.value : '');
  const finalPasscode = newPasscode || passcode;

  if (!transferCode && !name) {
    showToast('お名前を入力してください');
    return;
  }
  if (!transferCode && !phone && !birthday && !passcode) {
    showToast('電話番号・生年月日・現在のパスコードのうち1つ以上を入力してください');
    return;
  }
  if (passcode && !isValidPasscodeValue(passcode)) {
    showToast('現在のパスコードは4桁または6桁の数字で入力してください');
    return;
  }
  if (transferCode && !isValidTransferCodeValue(transferCode)) {
    showToast('引き継ぎコードは8桁の数字で入力してください');
    return;
  }
  if (!isValidPasscodeValue(finalPasscode)) {
    showToast('この端末で使うパスコードを4桁または6桁の数字で入力してください');
    return;
  }

  const isConfirmed = await showAppConfirm('入力された情報でログインしますか？\n現在の端末のデータは上書きされます。', {
    title: 'ログイン・復元',
    confirmLabel: 'ログインする',
    cancelLabel: '戻る'
  });
  if (!isConfirmed) return;

  if (btn) {
    btn.disabled = true;
    btn.dataset.originalText = btn.dataset.originalText || btn.textContent;
    btn.textContent = 'ログイン中...';
  }
  showToast('会員データを確認しています...');

  try {
    const res = await postToGAS({
      type: 'recoverAccount',
      name: name,
      phone: phone,
      birthday: birthday,
      passcode: passcode,
      transferCode: transferCode,
      newPasscode: newPasscode
    });

    if (res && res.status === 'ok' && res.user) {
      const u = res.user;
      await applyRecoveredUserState(u, finalPasscode, 'おかえりなさい、' + u.name + ' 様！\nデータを復元しました🌿');
      return;
    }

    await showAppAlert(res && res.message ? res.message : '一致する会員情報が見つかりませんでした。入力内容を確認してください。', {
      title: '復元エラー',
      confirmLabel: '閉じる'
    });
  } catch (e) {
    console.error('restoreAccount error:', e);
    showToast('通信エラーが発生しました');
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = btn.dataset.originalText || 'ログインして復元する ↺';
    }
  }
}

async function resetForgottenPasscodeByForm(options) {
  const settings = options || {};
  const nameInput = document.getElementById(settings.nameId);
  const phoneInput = document.getElementById(settings.phoneId);
  const birthdayInput = document.getElementById(settings.birthdayId);
  const passcodeInput = document.getElementById(settings.passcodeId);
  const btn = document.getElementById(settings.buttonId);
  const name = normalizeNameInput(nameInput ? nameInput.value : '');
  const phone = normalizePhoneInput(phoneInput ? phoneInput.value : '');
  const birthday = normalizeDateOnlyInput(birthdayInput ? birthdayInput.value : '');
  const newPasscode = normalizePasscodeInput(passcodeInput ? passcodeInput.value : '');

  if (!name || !phone || !birthday || !newPasscode) {
    showToast('お名前・電話番号・生年月日・新しいパスコードを入力してください');
    return;
  }
  if (!isValidPasscodeValue(newPasscode)) {
    showToast('新しいパスコードは4桁または6桁の数字で入力してください');
    return;
  }

  const isConfirmed = await showAppConfirm('入力された情報でパスコードを再設定しますか？\n再設定後、この端末へログインします。', {
    title: 'パスコードの再設定',
    confirmLabel: '再設定する',
    cancelLabel: '戻る'
  });
  if (!isConfirmed) return;

  if (btn) {
    btn.disabled = true;
    btn.dataset.originalText = btn.dataset.originalText || btn.textContent;
    btn.textContent = '再設定中...';
  }
  showToast('会員情報を確認しています...');

  try {
    const res = await postToGAS({
      type: 'resetForgottenPasscode',
      memberId: _profile && _profile.memberId ? _profile.memberId : '',
      name: name,
      phone: phone,
      birthday: birthday,
      newPasscode: newPasscode
    });

    if (res && res.status === 'ok' && res.user) {
      const u = res.user;
      await applyRecoveredUserState(u, newPasscode, 'パスコードを再設定しました。\n' + u.name + ' 様としてログインしました🌿');
      return;
    }

    await showAppAlert(res && res.message ? res.message : '本人確認ができませんでした。入力内容を確認してください。', {
      title: '再設定エラー',
      confirmLabel: '閉じる'
    });
  } catch (e) {
    console.error('resetForgottenPasscode error:', e);
    showToast('通信エラーが発生しました');
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = btn.dataset.originalText || '再設定してログインする';
    }
  }
}

function restoreAccountFromLockScreen() {
  return restoreAccountByForm({
    nameId: 'passcodeRestoreName',
    phoneId: 'passcodeRestorePhone',
    birthdayId: 'passcodeRestoreBirthday',
    passcodeId: 'passcodeRestorePasscode',
    transferCodeId: 'passcodeRestoreTransferCode',
    newPasscodeId: 'passcodeRestoreNewPasscode',
    buttonId: 'passcodeRestoreBtn'
  });
}

function resetForgottenPasscodeFromLockScreen() {
  return resetForgottenPasscodeByForm({
    nameId: 'passcodeResetName',
    phoneId: 'passcodeResetPhone',
    birthdayId: 'passcodeResetBirthday',
    passcodeId: 'passcodeResetNewPasscode',
    buttonId: 'passcodeResetBtn'
  });
}

function toggleSetupView(view) {
  const choice = document.getElementById('setupChoiceContent');
  const setup = document.getElementById('setupFormContent');
  const restore = document.getElementById('restoreFormContent');
  const change = document.getElementById('changePasscodeContent');
  const forgot = document.getElementById('forgotPasscodeContent');

  if (choice) choice.style.display = view === 'choice' ? 'block' : 'none';

  if (view === 'restore') {
    if (setup) setup.style.display = 'none';
    if (restore) restore.style.display = 'block';
    if (change) change.style.display = 'none';
    if (forgot) forgot.style.display = 'none';
    prefillRestoreAccountForm({
      nameId: 'restoreName',
      phoneId: 'restorePhone',
      birthdayId: 'restoreBirthday',
      passcodeId: 'restorePasscode',
      transferCodeId: 'restoreTransferCode',
      newPasscodeId: 'restoreNewPasscode'
    });
  } else if (view === 'forgot') {
    if (setup) setup.style.display = 'none';
    if (restore) restore.style.display = 'none';
    if (change) change.style.display = 'none';
    if (forgot) forgot.style.display = 'block';
    prefillForgotPasscodeForm({
      nameId: 'forgotPasscodeName',
      phoneId: 'forgotPasscodePhone',
      birthdayId: 'forgotPasscodeBirthday',
      passcodeId: 'forgotPasscodeNew',
      sourcePhoneId: 'restorePhone'
    });
  } else if (view === 'change') {
    if (setup) setup.style.display = 'none';
    if (restore) restore.style.display = 'none';
    if (change) change.style.display = 'block';
    if (forgot) forgot.style.display = 'none';
  } else if (view === 'choice') {
    if (setup) setup.style.display = 'none';
    if (restore) restore.style.display = 'none';
    if (change) change.style.display = 'none';
    if (forgot) forgot.style.display = 'none';
  } else {
    if (setup) setup.style.display = 'block';
    if (restore) restore.style.display = 'none';
    if (change) change.style.display = 'none';
    if (forgot) forgot.style.display = 'none';
  }
}

function openChangePasscode() {
  document.getElementById('newPasscode').value = '';
  toggleSetupView('change');
  const cancelBtn = document.getElementById('setupCancelBtn');
  if (cancelBtn) cancelBtn.style.display = 'inline-block';
  openModal('setupOverlay');
}

async function saveNewPasscode() {
  if (!_profile || !_profile.memberId) return;
  const newPass = normalizePasscodeInput(document.getElementById('newPasscode').value);

  if (!isValidPasscodeValue(newPass)) {
    showToast('4桁または6桁の数字を入力してください');
    return;
  }

  const btn = document.getElementById('changePasscodeBtn');
  btn.disabled = true;
  btn.textContent = '保存中...';

  try {
    const data = {
      type: 'updateUser',
      memberId: _profile.memberId,
      passcode: newPass
    };
    const res = await postToGAS(data);
    if (res && res.status === 'ok') {
      await storeLocalPasscode(newPass);
      isPasscodeAuthenticated = true;
      syncCurrentDeviceSession({ silent: true }).catch(function (error) {
        console.error('syncCurrentDeviceSession after saveNewPasscode error:', error);
      });
      updateTransferCodeDisplay(null);
      showToast('パスコードを変更しました🌿');
      closeSetupModal();
    } else if (res && res.status === 'queued') {
      await storeLocalPasscode(newPass);
      isPasscodeAuthenticated = true;
      showToast('パスコードを変更しました。通信回復後に同期します');
      closeSetupModal();
    } else {
      showToast('変更に失敗しました');
    }
  } catch (e) {
    console.error(e);
    showToast('通信エラーが発生しました');
  } finally {
    btn.disabled = false;
    btn.textContent = '変更を保存';
  }
}

async function issueTransferCode() {
  if (!_profile || !_profile.memberId) {
    showToast('会員情報を確認できません');
    return;
  }

  const isConfirmed = await showAppConfirm('この端末の会員情報で引き継ぎコードを発行しますか？\n新しい端末の「データの引き継ぎ・復元」で使えます。', {
    title: '引き継ぎコードの発行',
    confirmLabel: '発行する',
    cancelLabel: '戻る'
  });
  if (!isConfirmed) return;

  const btn = document.getElementById('issueTransferCodeBtn');
  if (btn) {
    btn.disabled = true;
    btn.dataset.originalText = btn.dataset.originalText || btn.textContent;
    btn.textContent = '発行中...';
  }

  try {
    const res = await postToGAS({
      type: 'issueTransferCode',
      memberId: _profile.memberId
    });

    if (res && res.status === 'ok' && res.transferCode) {
      updateTransferCodeDisplay(res);
      await showAppAlert(
        '引き継ぎコード：' + res.transferCode + '\n有効期限：' + (res.expiresAtLabel || formatCustomerDateYmdHm(res.expiresAt) || '---') + '\n\n新しい端末の「データの引き継ぎ・復元」で入力してください。',
        {
          title: '引き継ぎコードを発行しました',
          confirmLabel: '閉じる'
        }
      );
      return;
    }

    await showAppAlert(res && res.message ? res.message : '引き継ぎコードの発行に失敗しました。', {
      title: '発行エラー',
      confirmLabel: '閉じる'
    });
  } catch (e) {
    console.error('issueTransferCode error:', e);
    showToast('通信エラーが発生しました');
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = btn.dataset.originalText || '引き継ぎコードを発行する';
    }
  }
}


async function restoreAccount() {
  return restoreAccountByForm({
    nameId: 'restoreName',
    phoneId: 'restorePhone',
    birthdayId: 'restoreBirthday',
    passcodeId: 'restorePasscode',
    transferCodeId: 'restoreTransferCode',
    newPasscodeId: 'restoreNewPasscode',
    buttonId: 'restoreBtn'
  });
}

async function resetForgottenPasscode() {
  return resetForgottenPasscodeByForm({
    nameId: 'forgotPasscodeName',
    phoneId: 'forgotPasscodePhone',
    birthdayId: 'forgotPasscodeBirthday',
    passcodeId: 'forgotPasscodeNew',
    buttonId: 'forgotPasscodeBtn'
  });
}

async function saveProfile() {
  const nameEl = document.getElementById('setupName');
  const nameErr = document.getElementById('setupNameErr');
  const name = nameEl.value.trim();
  const phone = document.getElementById('setupPhone').value.trim();
  const birthday = document.getElementById('setupBirthday').value;
  const address = document.getElementById('setupAddress').value.trim();
  const isNew = !_profile;
  let passcode = '';

  if (isNew) {
    const passcodeEl = document.getElementById('setupPasscode');
    const passcodeErr = document.getElementById('setupPasscodeErr');
    passcode = normalizePasscodeInput(passcodeEl.value);
    if (!isValidPasscodeValue(passcode)) {
      passcodeEl.classList.add('error');
      passcodeErr.classList.add('show');
      passcodeEl.focus();
      return;
    }
    passcodeEl.classList.remove('error');
    passcodeErr.classList.remove('show');
  }

  // バリデーション
  if (!name) {
    nameEl.classList.add('error');
    nameErr.classList.add('show');
    nameEl.focus();
    return;
  }
  nameEl.classList.remove('error');
  nameErr.classList.remove('show');

  if (isNew) {
    const candidates = await checkExistingMemberCandidates(name, phone, birthday);
    const movedToRestore = await promptRecoveryCandidates(candidates, {
      name: name,
      phone: phone,
      birthday: birthday
    });
    if (movedToRestore) return;
  }

  const now = new Date();
  // 登録/更新日時を yyyy/M/d H:mm 形式にする
  const regDate = formatCustomerDateYmd(now);

  const memberId = isNew
    ? 'MYM-' + String(Math.floor(Math.random() * 9000) + 1000)
    : (_profile.memberId || 'MYM-' + String(Math.floor(Math.random() * 9000) + 1000));

  _profile = { name, phone, birthday, address, regDate, memberId };
  try {
    localStorage.setItem('mayumi_profile', JSON.stringify(_profile));
    localStorage.setItem('member_id', memberId); // 互換性のため個別にも保存
    if (_avatarData) localStorage.setItem('mayumi_avatar', _avatarData);
    if (_bannerData) localStorage.setItem('mayumi_banner', _bannerData);
  } catch (e) { }
  if (isNew && passcode) {
    await storeLocalPasscode(passcode);
    isPasscodeAuthenticated = true;
  }
  CUSTOMER_NAME = name;

  // UIを先に更新する
  closeSetupModal();
  if (isNew) {
    const homePage = document.getElementById('page-home');
    if (homePage) homePage.classList.add('active');
    switchPage('home');
  }
  updateProfileUI();
  showToast(isNew ? '登録完了しました 🌿' : 'プロフィールを更新しました ✅');

  // GASへユーザー情報を同期（非同期で実行し、待たずに次へ進む）
  (async function () {
    let avatarUrl = localStorage.getItem('mayumi_avatar_url');
    const avatarFallback = (_avatarData && _avatarData.startsWith('data:image/')) ? _avatarData : '';
    if (!avatarUrl && _avatarData && _avatarData.startsWith('data:image')) {
      try {
        const mimeMatch = _avatarData.match(/data:(image\/.+);base64,/);
        const mimeType = mimeMatch ? mimeMatch[1] : 'image/png';
        const ext = mimeType.split('/')[1] || 'png';
        const fname = "avatar_" + memberId + "." + ext;
        const res = await postToGAS({
          type: 'uploadImage',
          filename: fname,
          mimeType: mimeType,
          base64: _avatarData.split(',')[1]
        });
        if (res && res.url) {
          avatarUrl = res.url;
          localStorage.setItem('mayumi_avatar_url', avatarUrl);
        }
      } catch (e) { console.log('Avatar upload error:', e); }
    }

    await postToGAS({
      type: 'updateUser',
      memberId: memberId,
      avatar: avatarUrl || avatarFallback,
      name: name,
      phone: phone,
      birthday: birthday,
      address: address,
      kana: '',
      memo: '',
      passcode: isNew ? passcode : undefined,
      pushSubscription: getCurrentPushSubscriptionValue(isPushEnabled())
    });
    await syncCurrentDeviceSession({ silent: true });
    await syncRewardStatus(true);

    // OneSignal タグ設定
    if (window.OneSignalRef) {
      // OneSignalRef.User.addTag("status", status);
    }
  })();
}

function openMigrationModal() {
  const phoneInput = document.getElementById('migrationPhone');
  const passcodeInput = document.getElementById('migrationPasscode');
  const overlay = document.getElementById('migrationOverlay');
  if (phoneInput) phoneInput.value = _profile && _profile.phone ? _profile.phone : '';
  if (passcodeInput) passcodeInput.value = '';
  isPasscodeAuthenticated = false;
  if (overlay) overlay.classList.add('open');
  setTimeout(function () {
    if (passcodeInput) passcodeInput.focus();
  }, 60);
}

async function saveMigrationPasscode() {
  const phone = document.getElementById('migrationPhone').value.trim();
  const passcode = normalizePasscodeInput(document.getElementById('migrationPasscode').value);
  const btn = document.getElementById('migrationBtn');

  if (!phone || !isValidPasscodeValue(passcode)) {
    showToast('電話番号と4桁または6桁のパスコードを入力してください');
    return;
  }

  btn.disabled = true;
  btn.textContent = '保存中...';

  try {
    const data = {
      type: 'updateUser',
      memberId: _profile.memberId,
      phone: phone,
      passcode: passcode
    };
    const res = await postToGAS(data);
    if (res && res.status === 'ok') {
      _profile.phone = phone;
      localStorage.setItem('mayumi_profile', JSON.stringify(_profile));
      await storeLocalPasscode(passcode);
      isPasscodeAuthenticated = true;
      syncCurrentDeviceSession({ silent: true }).catch(function (error) {
        console.error('syncCurrentDeviceSession after saveMigrationPasscode error:', error);
      });
      showToast('パスコード設定を完了しました🌿');
      document.getElementById('migrationOverlay').classList.remove('open');
      const homePage = document.getElementById('page-home');
      if (homePage) homePage.classList.add('active');
      switchPage('home');
      updateProfileUI();
    } else if (res && res.status === 'queued') {
      _profile.phone = phone;
      localStorage.setItem('mayumi_profile', JSON.stringify(_profile));
      await storeLocalPasscode(passcode);
      isPasscodeAuthenticated = true;
      showToast('設定を保存しました。通信回復後に同期します');
      document.getElementById('migrationOverlay').classList.remove('open');
      const homePage = document.getElementById('page-home');
      if (homePage) homePage.classList.add('active');
      switchPage('home');
      updateProfileUI();
    } else {
      showToast('設定に失敗しました');
    }
  } catch (e) {
    console.error(e);
    showToast('通信エラーが発生しました');
  } finally {
    btn.disabled = false;
    btn.textContent = '設定してアプリを使う';
  }
}

function openEditProfile() {
  if (!_profile) { openSetupModal(false); return; }
  toggleSetupView('setup'); // Ensure setup view is visible and not restore or change
  // 編集モード：現在の値をセット
  document.getElementById('setupName').value = _profile.name || '';
  document.getElementById('setupPhone').value = _profile.phone || '';
  document.getElementById('setupBirthday').value = _profile.birthday || '';
  document.getElementById('setupAddress').value = _profile.address || '';
  // アバタープレビュー設定
  const preview = document.getElementById('avatarPreview');
  if (preview) preview.src = _avatarData || DEFAULT_AVATAR;
  // バナープレビュー設定
  const bannerPreview = document.getElementById('bannerPreview');
  if (bannerPreview) {
    bannerPreview.style.backgroundImage = _bannerData ? `url(${_bannerData})` : '';
  }
  document.getElementById('setupTitle').textContent = 'プロフィールを編集';
  document.getElementById('setupDesc').textContent = '項目を入力して、プロフィールを変更できます。';
  document.getElementById('setupBtn').textContent = '変更を保存する ✅';
  document.getElementById('setupCancelBtn').style.display = 'block';
  document.getElementById('avatarEditWrap').style.display = 'block';

  // Hide passcode field when editing (passcode has its own modal)
  const passcodeEl = document.getElementById('setupPasscode');
  if (passcodeEl && passcodeEl.parentElement) {
    passcodeEl.parentElement.style.display = 'none';
  }

  // Hide login switch link
  const restoreLinkWrap = document.getElementById('setupRestoreLinkWrap');
  if (restoreLinkWrap) restoreLinkWrap.style.display = 'none';

  document.getElementById('setupOverlay').classList.add('open');
}

function openSetupModal(isFirst) {
  const shouldShowChoice = !_profile;
  toggleSetupView(shouldShowChoice ? 'choice' : 'setup');
  document.getElementById('setupName').value = '';
  document.getElementById('setupPhone').value = '';
  document.getElementById('setupPasscode').value = '';
  const defStatus = document.querySelector('input[name="setupStatus"][value="妊娠中"]');
  if (defStatus) defStatus.checked = true;
  // アバタープレビューリセット
  const preview = document.getElementById('avatarPreview');
  if (preview) preview.src = DEFAULT_AVATAR;
  // バナープレビューリセット
  const bannerPreview = document.getElementById('bannerPreview');
  if (bannerPreview) bannerPreview.style.backgroundImage = '';
  document.getElementById('setupTitle').textContent = 'ようこそ！';
  document.getElementById('setupDesc').textContent = 'まゆみ助産院アプリへようこそ。\nはじめにプロフィールを登録してください。';
  document.getElementById('setupBtn').textContent = '登録してアプリを始める 🌿';
  document.getElementById('setupCancelBtn').style.display = 'none';
  document.getElementById('avatarEditWrap').style.display = 'block';

  // Show passcode field
  const passcodeEl = document.getElementById('setupPasscode');
  if (passcodeEl && passcodeEl.parentElement) {
    passcodeEl.parentElement.style.display = 'block';
  }

  // Show login switch link
  const restoreLinkWrap = document.getElementById('setupRestoreLinkWrap');
  if (restoreLinkWrap) restoreLinkWrap.style.display = 'block';

  document.getElementById('setupOverlay').classList.add('open');
}

function closeSetupModal() {
  document.getElementById('setupOverlay').classList.remove('open');
  // エラー状態リセット
  document.getElementById('setupName').classList.remove('error');
  document.getElementById('setupNameErr').classList.remove('show');
  const passcodeEl = document.getElementById('setupPasscode');
  const passcodeErr = document.getElementById('setupPasscodeErr');
  if (passcodeEl) passcodeEl.classList.remove('error');
  if (passcodeErr) passcodeErr.classList.remove('show');
}

function updateProfileUI() {
  updatePasscodeLoginUI();
  if (!_profile) return;
  const set = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
  set('homeBannerName', _profile.name + '様');
  set('mypageBannerName', _profile.name + '様');
  set('mypageMemberId', '会員ID：' + (_profile.memberId || '---'));
  set('profileName', _profile.name);
  set('profilePhone', _profile.phone || '未登録');
  set('profileBirthday', _profile.birthday ? formatCustomerDateYmd(_profile.birthday) : '未登録');
  set('profileAddress', _profile.address || '未登録');
  set('profileRegDate', _profile.regDate ? formatCustomerDateYmd(_profile.regDate) : '---');
  // アバター画像更新
  const avatarSrc = _avatarData || DEFAULT_AVATAR;
  const profileAvatar = document.getElementById('profileAvatar');
  if (profileAvatar) profileAvatar.src = avatarSrc;

  // バナー背景画像更新
  const banners = [
    document.getElementById('homeBanner'),
    document.getElementById('stampBanner'),
    document.getElementById('mypageBanner')
  ];
  banners.forEach(banner => {
    if (banner) {
      if (_bannerData) {
        banner.style.backgroundImage = `url(${_bannerData})`;
        banner.style.backgroundSize = 'cover';
        banner.style.backgroundPosition = 'center';
        banner.classList.add('has-bg');
      } else {
        banner.style.backgroundImage = '';
        banner.classList.remove('has-bg');
      }
    }
  });
  renderFavoriteList();
  renderUserDevices();
  renderRetryQueueStatus();
  renderAccessibilitySettings();
  renderCurrentDeviceGuidance();
}

function renderFavoriteList() {
  const list = document.getElementById('favoriteList');
  if (!list) return;
  const favorites = getFavoriteEntries();
  if (!favorites.length) {
    list.innerHTML = '<div style="text-align:center;font-size:13px;color:var(--text-light);padding:22px 0">お気に入りはまだありません</div>';
    return;
  }
  list.innerHTML = favorites.map(function (entry) {
    const subtitle = escapeHtml(entry.subtitle || '');
    const title = escapeHtml(entry.title || '');
    const dateLabel = escapeHtml(entry.dateLabel || '');
    const encodedKey = encodeURIComponent(entry.key || '');
    return `<div class="favorite-card"><div class="favorite-main" onclick="openFavoriteEntry('${encodedKey}')"><div class="favorite-meta">${subtitle}${dateLabel ? '<span class="favorite-date">' + dateLabel + '</span>' : ''}</div><div class="favorite-title">${title}</div></div><button class="favorite-remove-btn" type="button" onclick="removeFavoriteEntry('${encodedKey}')">削除</button></div>`;
  }).join('');
}

function openFavoriteEntry(key) {
  const resolvedKey = decodeURIComponent(String(key || ''));
  const entry = getFavoriteEntries().find(function (item) {
    return item && item.key === resolvedKey;
  });
  if (!entry) return;
  switchPage(entry.page || 'home');
  if (entry.kind === 'product') {
    const productIndex = PRODUCTS.findIndex(function (item) { return buildContentItemKey('product', item) === resolvedKey; });
    if (productIndex >= 0) openProductModal(productIndex);
  }
  if (entry.kind === 'blog') {
    const blogItem = blogItems.find(function (item) { return buildContentItemKey('blog', item) === resolvedKey; });
    if (blogItem) openBlogDetail(blogItem);
  }
  if (entry.kind === 'menu') {
    const menuIndex = USER_MENUS.findIndex(function (item) { return buildContentItemKey('menu', item) === resolvedKey; });
    if (menuIndex >= 0) openMenuDetail(menuIndex);
  }
  if (entry.kind === 'calendar') {
    const calendarIndex = calendarData.findIndex(function (item) { return buildContentItemKey('calendar', item) === resolvedKey; });
    if (calendarIndex >= 0) openCalendarEventDetail(calendarIndex);
  }
}

function removeFavoriteEntry(key) {
  const resolvedKey = decodeURIComponent(String(key || ''));
  saveFavoriteEntries(getFavoriteEntries().filter(function (entry) {
    return entry.key !== resolvedKey;
  }));
  showToast('お気に入りから外しました');
  refreshActiveDetailViews();
}

async function syncCurrentDeviceSession(options) {
  if (!_profile || !_profile.memberId) return;
  const opts = options || {};
  const response = await postToGAS(Object.assign({
    type: 'syncUserDeviceSession',
    memberId: _profile.memberId
  }, buildCurrentDeviceSessionPayload()), {
    skipRetryQueue: opts.skipRetryQueue === true,
    silent: opts.silent === true
  });
  if (response && response.status === 'ok' && Array.isArray(response.devices)) {
    cacheDeviceSessions(response.devices);
    renderUserDevices();
  }
}

async function loadUserDevices() {
  if (!_profile || !_profile.memberId) {
    cacheDeviceSessions([]);
    renderUserDevices();
    return;
  }
  const response = await getFromGAS('getUserDevices', { memberId: _profile.memberId });
  if (response && response.status === 'ok' && Array.isArray(response.devices)) {
    cacheDeviceSessions(response.devices);
  }
  renderUserDevices();
}

function renderUserDevices() {
  const list = document.getElementById('deviceSessionList');
  if (!list) return;
  const devices = getCachedDeviceSessions();
  if (!devices.length) {
    list.innerHTML = '<div style="text-align:center;font-size:13px;color:var(--text-light);padding:22px 0">この端末情報は次回同期時に表示されます</div>';
    return;
  }
  const currentDeviceId = getCurrentDeviceSessionId();
  list.innerHTML = devices.map(function (device) {
    const isCurrent = String(device.deviceId || '') === currentDeviceId;
    const flags = [];
    if (device.passcodeEnabled) flags.push('パスコードON');
    if (device.pushEnabled) flags.push('通知ON');
    return `<div class="device-session-card"><div class="device-session-head"><div><div class="device-session-title">${escapeHtml(device.label || 'この端末')}</div><div class="device-session-sub">${escapeHtml(device.platform || '')}${device.appVersion ? ' / ' + escapeHtml(device.appVersion) : ''}</div></div><div class="device-session-chip">${isCurrent ? '使用中' : '登録済み'}</div></div><div class="device-session-meta">最終利用: ${escapeHtml(formatCustomerDateYmdHm(device.lastSeenAt) || formatCustomerDateYmd(device.lastSeenAt) || '---')}</div><div class="device-session-flags">${flags.length ? flags.map(function (flag) { return '<span>' + escapeHtml(flag) + '</span>'; }).join('') : '<span>設定情報なし</span>'}</div>${isCurrent ? '' : '<button class="btn secondary device-session-remove" type="button" onclick="removeUserDeviceSession(\'' + escapeHtml(device.deviceId) + '\')">この端末を一覧から外す</button>'}</div>`;
  }).join('');
}

async function removeUserDeviceSession(deviceId) {
  if (!_profile || !_profile.memberId || !deviceId) return;
  const confirmed = await showAppConfirm('この端末を一覧から外しますか？\n古い端末や使っていない端末だけを外してください。', {
    title: '端末一覧',
    confirmLabel: '外す',
    cancelLabel: '戻る',
    confirmVariant: 'danger'
  });
  if (!confirmed) return;
  const response = await postToGAS({
    type: 'removeUserDeviceSession',
    memberId: _profile.memberId,
    deviceId: deviceId
  });
  if (response && (response.status === 'ok' || response.status === 'queued')) {
    cacheDeviceSessions((response && response.devices) || getCachedDeviceSessions().filter(function (device) {
      return device.deviceId !== deviceId;
    }));
    renderUserDevices();
    showToast(response.status === 'queued' ? '通信回復後に端末一覧へ反映します' : '端末一覧から外しました');
  }
}

async function checkExistingMemberCandidates(name, phone, birthday) {
  if (!name && !phone && !birthday) return [];
  const response = await getFromGAS('getRecoveryCandidates', {
    name: name,
    phone: phone,
    birthday: birthday
  });
  if (response && response.status === 'ok' && Array.isArray(response.candidates)) {
    return response.candidates;
  }
  return [];
}

function prefillRecoveryFormsFromProfile(values) {
  const payload = values || {};
  [
    ['restoreName', payload.name],
    ['restorePhone', payload.phone],
    ['restoreBirthday', payload.birthday],
    ['passcodeRestoreName', payload.name],
    ['passcodeRestorePhone', payload.phone],
    ['passcodeRestoreBirthday', payload.birthday]
  ].forEach(function (pair) {
    const el = document.getElementById(pair[0]);
    if (el && pair[1] !== undefined) el.value = pair[1] || '';
  });
}

async function promptRecoveryCandidates(candidates, formValues) {
  if (!candidates || !candidates.length) return false;
  const summary = candidates.slice(0, 3).map(function (candidate) {
    const reasonText = Array.isArray(candidate.reasons) ? candidate.reasons.join('・') : '一致';
    return '・' + (candidate.name || '会員') + ' / ' + (candidate.memberId || '---') + ' / 一致: ' + reasonText;
  }).join('\n');
  const shouldMove = await showAppConfirm(
    '以前登録した可能性がある会員が見つかりました。\n新規登録ではなく復元を使うと、同じお名前で別の会員IDが増えるのを防げます。\n\n' + summary + '\n\n復元画面へ移動しますか？',
    {
      title: '以前登録した情報が見つかりました',
      confirmLabel: '復元へ進む',
      cancelLabel: '新規登録を続ける'
    }
  );
  if (!shouldMove) return false;
  prefillRecoveryFormsFromProfile(formValues);
  toggleSetupView('restore');
  return true;
}

function checkFirstLaunch() {
  if (!_profile) {
    isPasscodeAuthenticated = false;
    // 初回起動：オンボーディングを即座に表示
    const screen = document.getElementById('onboardingScreen');
    if (screen) screen.classList.add('show');
  } else {
    CUSTOMER_NAME = _profile.name;
    updateProfileUI();
    activatePageSilently(getPreferredStartupPage());

    if (needsRequiredPasscodeSetup()) {
      openMigrationModal();
      return;
    }

    if (consumePasscodeUnlockSkippedOnce()) {
      isPasscodeAuthenticated = true;
      return;
    }

    if (!isPasscodeLoginEnabled()) {
      isPasscodeAuthenticated = true;
      closePasscodeOverlay();
      return;
    }

    isPasscodeAuthenticated = false;
    openPasscodeOverlay();
  }
}

function startAppFromOnboarding() {
  const screen = document.getElementById('onboardingScreen');
  screen.classList.remove('show');
  setTimeout(() => {
    openSetupModal(true);
  }, 600); // フェードアウトを待つ
}

function normalizeOpenPageTarget(value) {
  const raw = String(value || '').trim().toLowerCase();
  const map = {
    home: 'home',
    shop: 'shop',
    news: 'blog',
    blog: 'blog',
    calendar: 'calendar',
    notices: 'notices',
    notice: 'notices',
    menu: 'menu-list',
    'menu-list': 'menu-list',
    mypage: 'mypage',
    profile: 'mypage'
  };
  return map[raw] || '';
}

function openPageFromNotificationTarget(target) {
  const page = normalizeOpenPageTarget(target);
  if (!page) return false;
  if (!_profile && page !== 'home') {
    openSetupModal(true);
    return true;
  }
  switchPage(page);
  return true;
}

// ===== URLパラメータからのスタンプ付与処理 =====
function checkUrlParams() {
  const params = new URLSearchParams(window.location.search);
  const restorePage = normalizeAppPageName(params.get('restorePage'));
  if (restorePage) {
    clearPendingAppReloadTarget();
    clearPendingUpdateRestorePage();
    if (getActivePageName() !== restorePage && _profile) {
      switchPage(restorePage);
    }
    params.delete('restorePage');
    const nextUrl = window.location.pathname + (params.toString() ? '?' + params.toString() : '');
    window.history.replaceState({}, document.title, nextUrl);
    return;
  }
  const openTarget = params.get('open');
  if (openTarget) {
    setTimeout(function () {
      openPageFromNotificationTarget(openTarget);
    }, 400);
    params.delete('open');
    const nextUrl = window.location.pathname + (params.toString() ? '?' + params.toString() : '');
    window.history.replaceState({}, document.title, nextUrl);
    return;
  }
  if (params.get('action') === 'add_stamp') {
    setTimeout(() => {
      // すでにプロフィール登録済みの場合のみスタンプを付与
      if (_profile) {
        addStamp();
        switchPage('home');
      } else {
        // 未登録の場合は一旦登録を促す（次回以降のアクセスに備える）
        showToast('まずはプロフィールを登録してください🌿');
      }
      // URLをクリーンアップしてリロード時の二重付与を防止
      window.history.replaceState({}, document.title, window.location.pathname);
    }, 800);
  }
}

// OneSignalの状態監視（後半の初期化ブロックでまとめて処理されるため、ここはコメントアウト）
/*
window.OneSignalDeferred = window.OneSignalDeferred || [];
OneSignalDeferred.push(function (OneSignal) { ... });
*/

// ネイティブ環境でのプッシュ状態同期
async function syncNativePushStatus() {
  const Push = await waitForCapacitorPushPlugin(2000);
  if (!Push) return;
  try {
    const perm = await Push.checkPermissions();
    if (perm.receive === 'granted') {
      if (getStoredPushPreference() === 'false') {
        updatePushUI('off');
      } else {
        localStorage.setItem(PUSH_ENABLED_STORAGE_KEY, 'true');
        updatePushUI('on');
      }
    } else {
      localStorage.setItem(PUSH_ENABLED_STORAGE_KEY, 'false');
      setStoredNativePushPlayerId('');
      setStoredNativePushToken('');
      updatePushUI('off');
    }
  } catch (e) {
    console.error('Native status sync error:', e);
  }
}
// 起動時とマイページ表示時に同期
syncNativePushStatus();

// 通知の有効/無効を切り替え
async function togglePush() {
  const Push = await waitForCapacitorPushPlugin(4000);

  // ネイティブ環境 (Capacitor) の場合
  if (Push) {
    console.log('togglePush: Native environment detected (v2-debug)');
    try {
      if (isPushEnabled()) {
        const shouldDisable = await showAppConfirm('プッシュ通知をオフにしますか？', {
          title: '通知設定',
          confirmLabel: 'オフにする',
          cancelLabel: '戻る',
          confirmVariant: 'danger'
        });
        if (!shouldDisable) return;

        if (typeof Push.unregister === 'function') {
          try {
            await Push.unregister();
          } catch (unregisterError) {
            console.error('togglePush Native unregister error:', unregisterError);
          }
        }

        const unsubscribeResult = await postToGAS({
          type: 'unsubscribePush',
          playerId: getStoredNativePushPlayerId(),
          memberId: _profile && _profile.memberId ? _profile.memberId : '',
          pushToken: getStoredNativePushToken()
        });
        if (unsubscribeResult && unsubscribeResult.warning) {
          console.warn('unsubscribePush warning:', unsubscribeResult.warning);
        }

        await applyPushEnabledState(false, {
          clearNativePlayerId: true,
          clearNativeToken: true
        });
        showToast('通知をオフにしました');
        return;
      }

      let perm = await Push.checkPermissions();
      console.log('togglePush: current perm:', perm.receive);

      if (perm.receive !== 'granted') {
        perm = await Push.requestPermissions();
      }
      if (perm.receive === 'granted') {
        await Push.register();
        await applyPushEnabledState(true);
        showToast('通知を有効にしました！ ✅');
      } else {
        await applyPushEnabledState(false, { syncProfile: false });
        await showAppAlert('通知が許可されませんでした。\niPhoneの設定またはブラウザ設定で通知を許可してください。', {
          title: '通知設定',
          confirmLabel: '閉じる'
        });
      }
    } catch (e) {
      console.error('togglePush Native error:', e);
      showToast('エラーが発生しました');
    }
    return;
  }

  // Capacitor環境だと推測されるがプラグインが見つからない場合、1回だけ警告を出してWebフローへ。
  if (isLikelyCapacitorRuntime() && !Push) {
    console.warn('Native environment suspected but Push plugin not found. Falling back to Web/OneSignal...');
  }

  // Web環境 (OneSignal SDK) の場合
  console.log('togglePush: Web environment flow (v2-debug)');
  for (let i = 0; i < 30; i++) {
    if (window.OneSignalRef && window.OneSignalRef.Notifications) break;
    await new Promise(r => setTimeout(r, 100));
  }

  const OS = window.OneSignalRef;
  if (!OS || !OS.Notifications) {
    await handleNativeNotificationFallback();
    return;
  }

  try {
    const permissionGranted = await getOneSignalPermissionState(OS);
    const sub = getOneSignalPushSubscription(OS);
    if (!sub) {
      await handleNativeNotificationFallback();
      return;
    }

    console.log('togglePush: OneSignal available. Permission status:', permissionGranted);
    if (permissionGranted) {
      const isOn = !!sub.optedIn;
      console.log('togglePush: Current subscription status:', isOn);
      if (isOn) {
        const shouldOptOut = await showAppConfirm('プッシュ通知をオフにしますか？', {
          title: '通知設定',
          confirmLabel: 'オフにする',
          cancelLabel: '戻る',
          confirmVariant: 'danger'
        });
        if (shouldOptOut) {
          await syncWebPushState({
            oneSignal: OS,
            syncProfile: true,
            targetEnabled: false,
            reconcile: false
          });
          showToast('通知をオフにしました');
        }
      } else {
        console.log('togglePush: Opting in...');
        const enabled = await syncWebPushState({
          oneSignal: OS,
          syncProfile: true,
          targetEnabled: true,
          reconcile: false
        });
        if (enabled) {
          showToast('有効化しました！ ✅');
        } else {
          await showAppAlert('通知が許可されませんでした。\nSafariの設定を確認してください。', {
            title: '通知設定',
            confirmLabel: '閉じる'
          });
        }
      }
      return;
    }

    console.log('togglePush: Requesting initial permission...');
    await OS.Notifications.requestPermission();
    const finalStatus = await syncWebPushState({
      oneSignal: OS,
      syncProfile: true,
      targetEnabled: true,
      reconcile: false
    });
    console.log('togglePush: Permission result status:', finalStatus);
    if (finalStatus) {
      showToast('有効化しました！ ✅');
    } else {
      await showAppAlert('通知が許可されませんでした。\nSafariの設定を確認してください。', {
        title: '通知設定',
        confirmLabel: '閉じる'
      });
    }
  } catch (e) {
    console.error('togglePush Web error:', e);
    showToast('通知設定中にエラーが発生しました');
  }
}

async function handleNativeNotificationFallback() {
  // OneSignalが使えない場合（SDKエラーなど）の保険として、標準ブラウザ通知での切り替えを試みる
  if (isPushEnabled()) {
    await applyPushEnabledState(false, { syncProfile: false });
    showToast('通知設定（ブラウザ版）をオフにしました');
    return;
  }

  if (!('Notification' in window)) {
    showToast('通知が利用できません。ブラウザの設定で通知を許可してください。');
    return;
  }

  try {
    const perm = await Notification.requestPermission();
    if (perm === 'granted') {
      await applyPushEnabledState(true, { syncProfile: false });
      showToast('通知設定（ブラウザ版）をオンにしました！ ✅');
    } else {
      await applyPushEnabledState(false, { syncProfile: false });
      await showAppAlert('通知が許可されませんでした。\nブラウザの設定で通知（プッシュ通知）が許可されているか確認してください。', {
        title: '通知設定',
        confirmLabel: '閉じる'
      });
    }
  } catch (e) {
    console.error('Notification fallback error:', e);
    showToast('通知設定中にエラーが発生しました');
  }
}

// ※ 重複定義のため削除（3180行目付近に集約）

// ===== 初期化 (並列実行で高速化) =====
async function initApp() {
  if (initAppStarted) return;
  initAppStarted = true;
  await initSecureLocalStore();
  applyAccessibilitySettings();
  renderRetryQueueStatus();
  renderCurrentDeviceGuidance();

  // 最初に初回起動チェックを行い、UI（オンボーディング or ホーム）を表示する
  checkFirstLaunch();

  const versionGate = await ensureSupportedAppVersion();
  if (versionGate.blocked) {
    return;
  }
  if (versionGate.needsWebUpdate) {
    await applyPendingAppUpdate();
    return;
  }

  loadStampRewards();
  updateStampUI();
  checkUrlParams();

  // Capacitor環境 (iOS/Androidネイティブアプリ) の場合の通知設定
  (function () {
    let retryCount = 0;
    const maxRetries = 50; // 5 seconds total

    async function attemptInit() {
      console.log('Native Features: Checking for Capacitor... (Retry: ' + retryCount + ')');

      if (window.Capacitor && Capacitor.Plugins && Capacitor.Plugins.PushNotifications) {
        console.log('Native Features: Capacitor found! Initializing...');
        await runNativeInit();
      } else if (retryCount < maxRetries) {
        retryCount++;
        setTimeout(attemptInit, 100);
      } else {
        console.error('Native Features: Capacitor failure (Timeout 5s).');
      }
    }

    async function runNativeInit() {
      const { PushNotifications, Badge } = Capacitor.Plugins;

      PushNotifications.addListener('registration', async (token) => {
        console.log('--- CAPACITOR PUSH REGISTERED (v2-debug) ---');
        console.log('Push token: ' + token.value);
        setStoredNativePushToken(token.value);
        try {
          // OneSignal API への登録
          const payload = {
            app_id: "5f6e01a9-64ac-4cf6-9ea6-438a721d55fb",
            device_type: 0, // iOS
            identifier: token.value,
            language: 'ja',
            test_type: 1 // Sandbox (Xcodeからのデバッグ実行時に必要)
          };
          console.log('OneSignal Registering Payload (v2-debug):', JSON.stringify(payload));

          const response = await fetch('https://onesignal.com/api/v1/players', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
          });
          const resData = await response.json();
          console.log('OneSignal API Response (v2-debug):', resData);
          const playerId = resData && resData.id ? String(resData.id) : '';
          if (playerId) {
            setStoredNativePushPlayerId(playerId);
          }
          await applyPushEnabledState(true, {
            subscriptionValue: playerId || token.value
          });
        } catch (e) {
          console.error('OneSignal API Registration Error (v2-debug):', e);
          await applyPushEnabledState(true, {
            subscriptionValue: token.value
          });
        }
      });

      PushNotifications.addListener('registrationError', (err) => {
        console.error('Push Registration Listener Error:', err);
      });

      try {
        let perm = await PushNotifications.checkPermissions();
        console.log('Current Native Perm:', perm.receive);

        if (perm.receive === 'granted') {
          if (getStoredPushPreference() === 'false') {
            if (typeof PushNotifications.unregister === 'function') {
              try {
                await PushNotifications.unregister();
                console.log('Native push kept disabled by saved preference');
              } catch (unregisterError) {
                console.error('Push unregister on init failed:', unregisterError);
              }
            }
            setStoredNativePushPlayerId('');
            setStoredNativePushToken('');
            updatePushUI('off');
          } else {
            await PushNotifications.register();
            console.log('Native Push Registration Call Sent');
          }
        } else if (perm.receive === 'prompt') {
          localStorage.setItem(PUSH_ENABLED_STORAGE_KEY, 'false');
          setStoredNativePushPlayerId('');
          setStoredNativePushToken('');
          updatePushUI('off');
          console.log('Native push permission prompt deferred until user taps the toggle');
        } else {
          localStorage.setItem(PUSH_ENABLED_STORAGE_KEY, 'false');
          setStoredNativePushPlayerId('');
          setStoredNativePushToken('');
          updatePushUI('off');
          console.warn('Native Push Permission not granted.');
        }
      } catch (e) {
        console.error('Push Initialization Exception:', e);
      }

      PushNotifications.addListener('pushNotificationReceived', async (n) => {
        console.log('Native Push Received:', n);
        showToast('通知を受信しました: ' + (n.title || '新しい通知'));
        refreshNoticeFeed().catch(function (error) {
          console.error('refreshNoticeFeed from native push error:', error);
        });
        try {
          // @capawesome/capacitor-badge スタイル
          const count = await Badge.get();
          await Badge.set({ count: (count.count || 0) + 1 });
        } catch (e) {
          console.error('Badge Update Exception:', e);
        }
      });

      PushNotifications.addListener('pushNotificationActionPerformed', async (action) => {
        console.log('Native Push Action:', action);
        try {
          await Badge.clear();
        } catch (e) { }
        const payload = action && action.notification && action.notification.data ? action.notification.data : {};
        const target = payload.openPage || payload.targetPage || payload.page || '';
        if (target) {
          setTimeout(function () {
            openPageFromNotificationTarget(target);
          }, 250);
        }
      });

      try {
        await Badge.set({ count: 0 });
        console.log('Badge cleared on startup (v2-debug)');
      } catch (e) { }
    }

    // 実行開始
    attemptInit();
  })();

  // データを並列で取得（確実に動作する個別ロード方式）
  fetchLatestManagedContent({
    refreshSupportFaq: true,
    refreshMenus: false,
    refreshOrderHistory: false
  }).then(() => {
    isDataLoaded = true;
    const initializedBaseline = ensureSeenBaselineInitialized();
    if (initializedBaseline) {
      renderBlogList('homeNewsList', 3);
      if (document.getElementById('page-blog').classList.contains('active')) renderDividedBlogList();
      if (document.getElementById('page-notices').classList.contains('active')) renderPushNotices();
      renderMenus();
      renderCalendarEventLists();
      loadProducts().catch(function (error) {
        console.error('loadProducts baseline refresh error:', error);
      });
    }
    updateNavBadges();
    loadUserDevices().catch(function (error) {
      console.error('loadUserDevices error:', error);
    });
    syncCurrentDeviceSession({ silent: true }).catch(function (error) {
      console.error('syncCurrentDeviceSession init error:', error);
    });
    flushRetryQueue().catch(function (error) {
      console.error('flushRetryQueue init error:', error);
    });
  }).catch(e => console.error('Initial load error:', e));
}
initApp().catch(e => console.error('initApp error:', e));

// アプリ再表示時にデータを自動更新（タスクキル後の再起動対応）
let lastFetchTime = Date.now();
document.addEventListener('visibilitychange', function () {
  if (document.visibilityState === 'hidden') {
    appHiddenAt = Date.now();
    return;
  }

  if (document.visibilityState === 'visible') {
    if (needsRequiredPasscodeSetup()) {
      openMigrationModal();
      appHiddenAt = 0;
      return;
    }
    if (shouldRequirePasscodeLock() &&
      isPasscodeAuthenticated &&
      appHiddenAt &&
      (Date.now() - appHiddenAt) >= PASSCODE_RESUME_LOCK_DELAY_MS) {
      isPasscodeAuthenticated = false;
      openPasscodeOverlay();
    }
    appHiddenAt = 0;

    ensureSupportedAppVersion().then(function (versionGate) {
      if (versionGate && versionGate.blocked) {
        showToast('アプリの更新が必要です');
        return;
      }

      const elapsed = Date.now() - lastFetchTime;
      // 前回取得から1分以上経過していたら再取得
      if (elapsed > 60000) {
        lastFetchTime = Date.now();
        const activePageId = document.querySelector('.page.active')?.id || '';
        const fetchTask = activePageId === 'page-notices'
          ? refreshNoticeFeed()
          : fetchLatestManagedContent({
            refreshSupportFaq: true,
            refreshMenus: false,
            refreshOrderHistory: false
          });
        fetchTask.then(() => {
          isDataLoaded = true;
          updateNavBadges();
        }).catch(function () { });
      } else {
        // 1分以内でもバッジの再チェック（未読状態の維持）
        updateNavBadges();
      }

      syncCurrentDeviceSession({ silent: true }).catch(function (error) {
        console.error('syncCurrentDeviceSession resume error:', error);
      });
      flushRetryQueue().catch(function (error) {
        console.error('flushRetryQueue resume error:', error);
      });

      // アプリ復帰時にバッジをクリア (iOS用)
      if (window.Capacitor) {
        import('@capawesome/capacitor-badge').then(({ Badge }) => {
          Badge.clear().catch(e => console.error('Badge clear error on resume:', e));
        }).catch(e => console.error('Badge import error on resume:', e));
      }
    }).catch(function (e) {
      console.error('visibilitychange update check error:', e);
    });
  }
});

setInterval(function () {
  if (document.visibilityState !== 'visible' || !isDataLoaded) return;
  const activePageId = document.querySelector('.page.active')?.id || '';
  if (activePageId !== 'page-notices') return;

  const elapsed = Date.now() - lastFetchTime;
  if (elapsed <= 60000) return;

  lastFetchTime = Date.now();
  refreshNoticeFeed().then(function () {
    isDataLoaded = true;
    updateNavBadges();
  }).catch(function (e) {
    console.error('periodic notice feed refresh error:', e);
  });
}, 90000);

window.addEventListener('online', function () {
  flushRetryQueue().catch(function (error) {
    console.error('flushRetryQueue online error:', error);
  });
});

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.addEventListener('controllerchange', function () {
    if (!appUpdateContext.needsReload) return;
    if (appUpdateContext.reloadTimerId) {
      clearTimeout(appUpdateContext.reloadTimerId);
      appUpdateContext.reloadTimerId = 0;
    }
    const pendingTarget = getPendingAppReloadTarget();
    clearPendingAppReloadTarget();
    if (pendingTarget) {
      window.location.href = pendingTarget;
      return;
    }
    const url = new URL(window.location.href);
    if (!url.searchParams.get('upd')) {
      url.searchParams.set('upd', Date.now());
      window.location.href = url.toString();
    }
  });
}



(async function () {
  const isNative = isLikelyCapacitorRuntime() ||
    String(window.location && window.location.href ? window.location.href : '').includes('localhost');

  if (!isNative) {
    // 1. OneSignal SDK の読み込み
    const script = document.createElement('script');
    script.src = "https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.page.js";
    script.async = true;
    document.head.appendChild(script);

    window.OneSignalDeferred = window.OneSignalDeferred || [];
    OneSignalDeferred.push(async function (OneSignal) {
      window.OneSignalRef = OneSignal;

      // --- 1. OneSignal 初期化 (パス自動検知版) ---
      const appBase = window.location.pathname.substring(0, window.location.pathname.lastIndexOf('/') + 1) || '/';
      console.log('OneSignal: Auto-detected base path:', appBase);

      OneSignal.init({
        appId: "5f6e01a9-64ac-4cf6-9ea6-438a721d55fb",
        serviceWorkerPath: appBase + 'OneSignalSDKWorker.js',
        serviceWorkerParam: { scope: appBase },
        allowLocalhostAsSecureOrigin: true,
      }).then(() => {
        console.log('OneSignal: Init Success!');
      }).catch(e => {
        console.error('OneSignal: Init Error:', e);
      }).finally(() => {
        syncWebPushState({
          oneSignal: OneSignal,
          syncProfile: false
        }).catch(err => {
          console.error('OneSignal: Initial push sync error:', err);
        });

        // 15秒おきに購読状態を同期
        const syncInterval = setInterval(async () => {
          try {
            await syncWebPushState({
              oneSignal: OneSignal,
              syncProfile: false
            });
          } catch (err) {
            console.error('OneSignal: Push sync error:', err);
          }
        }, 15000);

        setTimeout(() => clearInterval(syncInterval), 120000); // 2分間トライ
      });

      // イベントリスナーの集約
      OneSignal.Notifications.addEventListener("permissionChange", function (perm) {
        syncWebPushState({
          oneSignal: OneSignal,
          syncProfile: false,
          reconcile: false
        }).catch(err => {
          console.error('OneSignal: permissionChange sync error:', err, perm);
        });
      });
      OneSignal.User.PushSubscription.addEventListener("change", function (e) {
        syncWebPushState({
          oneSignal: OneSignal,
          syncProfile: false,
          reconcile: false
        }).catch(err => {
          console.error('OneSignal: subscriptionChange sync error:', err, e);
        });
      });
    });
  }
})();

let _qrScanning = false;
let _qrStream = null;
let _qrStarting = false;

async function openScannerModal() {
  if (typeof stampCount !== 'undefined' && stampCount >= 10) {
    showToast('10個達成！新しいカードを取得してください');
    return;
  }
  const permissionState = await getCameraPermissionState();
  if (permissionState === 'denied') {
    await showCameraPermissionRecoveryDialog('denied');
    return;
  }
  const shouldOpen = await showCameraPermissionRecoveryDialog(permissionState);
  if (!shouldOpen) return;
  const modal = document.getElementById('scannerModal');
  if (modal) modal.classList.add('open');
  await startScanner();
}

function closeScannerModal(e) {
  if (e) {
    if (e.currentTarget.id === 'scannerModal' || e.target.tagName.toLowerCase() === 'button') {
      const modal = document.getElementById('scannerModal');
      if (modal) modal.classList.remove('open');
      stopScanner();
    }
  } else {
    const modal = document.getElementById('scannerModal');
    if (modal) modal.classList.remove('open');
    stopScanner();
  }
}

async function startScanner() {
  if (_qrStarting) return;
  const video = document.getElementById('qr-video');
  const canvasElement = document.getElementById('qr-canvas');
  if (!video || !canvasElement) return;
  if (!navigator.mediaDevices || typeof navigator.mediaDevices.getUserMedia !== 'function') {
    showAppAlert('この端末ではカメラを起動できませんでした。\nブラウザまたは端末の設定をご確認ください。', {
      title: 'カメラを利用できません',
      confirmLabel: '閉じる'
    }).then(function () {
      closeScannerModal();
    });
    return;
  }
  const canvas = canvasElement.getContext('2d');
  const scanLine = document.getElementById('qr-scan-line');
  const permissionState = await getCameraPermissionState();

  if (permissionState === 'denied') {
    closeScannerModal();
    await showCameraPermissionRecoveryDialog('denied');
    return;
  }

  _qrStarting = true;
  _qrScanning = true;
  scanLine.style.display = 'block';
  scanLine.classList.add('scanning');

  navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } })
    .then(function (stream) {
      _qrStarting = false;
      _qrStream = stream;
      video.srcObject = stream;
      video.setAttribute('playsinline', true);
      video.play();
      requestAnimationFrame(tick);
    })
    .catch(async function (err) {
      console.error("Camera error:", err);
      _qrStarting = false;
      _qrScanning = false;
      if (scanLine) {
        scanLine.style.display = 'none';
        scanLine.classList.remove('scanning');
      }

      let helpMsg = 'カメラの起動に失敗しました。';
      const errorName = String(err && err.name || '');
      const isPermissionError = errorName === 'NotAllowedError' || errorName === 'PermissionDeniedError' || errorName === 'SecurityError';
      if (errorName === 'NotAllowedError' || errorName === 'PermissionDeniedError') {
        helpMsg = 'カメラの使用が許可されていません。\nこのあと表示される確認で「許可」を選ぶか、iPhone / Android の設定でブラウザまたはアプリのカメラを許可してください。';
      } else if (errorName === 'NotReadableError' || errorName === 'TrackStartError') {
        helpMsg = 'ほかのアプリでカメラを使用している可能性があります。\nほかのアプリを閉じてから、もう一度お試しください。';
      } else if (errorName === 'NotFoundError' || errorName === 'DevicesNotFoundError') {
        helpMsg = 'この端末で利用できるカメラが見つかりませんでした。';
      } else if (location.protocol !== 'https:') {
        helpMsg = '接続が保護されていないため、カメラを利用できません。\nHTTPSで開き直してください。';
      }

      if (isPermissionError) {
        closeScannerModal();
        const latestPermissionState = await getCameraPermissionState();
        const inferredState = latestPermissionState === 'unknown' ? 'denied' : latestPermissionState;
        const shouldRetry = await showCameraPermissionRecoveryDialog(inferredState);
        if (shouldRetry) {
          const modal = document.getElementById('scannerModal');
          if (modal) modal.classList.add('open');
          await startScanner();
        }
        return;
      }

      const shouldRetry = await showAppConfirm(helpMsg + '\n\nもう一度カメラを起動しますか？', {
        title: 'カメラを利用できません',
        confirmLabel: 'もう一度試す',
        cancelLabel: '閉じる'
      });

      if (shouldRetry) {
        await startScanner();
      } else {
        closeScannerModal();
      }
    });

  let frameCount = 0;
  function tick() {
    if (!_qrScanning) return;
    if (video.readyState === video.HAVE_ENOUGH_DATA) {
      frameCount++;
      if (frameCount % 3 === 0) { // Throttle processing rate
        canvasElement.height = video.videoHeight;
        canvasElement.width = video.videoWidth;
        canvas.drawImage(video, 0, 0, canvasElement.width, canvasElement.height);
        var imageData = canvas.getImageData(0, 0, canvasElement.width, canvasElement.height);

        if (typeof jsQR !== 'undefined') {
          var code = jsQR(imageData.data, imageData.width, imageData.height, {
            inversionAttempts: "dontInvert",
          });

          if (code && code.data.includes('action=add_stamp')) {
            _qrScanning = false;
            closeScannerModal();
            showToast('QRコードを読み取りました🎉');
            setTimeout(() => {
              if (typeof addStamp === 'function') addStamp();
            }, 400);
            return;
          }
        }
      }
    }
    requestAnimationFrame(tick);
  }
}

function stopScanner() {
  _qrStarting = false;
  _qrScanning = false;
  const scanLine = document.getElementById('qr-scan-line');
  const video = document.getElementById('qr-video');
  if (scanLine) {
    scanLine.style.display = 'none';
    scanLine.classList.remove('scanning');
  }
  if (_qrStream) {
    _qrStream.getTracks().forEach(track => track.stop());
    _qrStream = null;
  }
  if (video) {
    try { video.pause(); } catch (e) { }
    video.srcObject = null;
  }
}

// メニュー一覧取得
async function loadMenus(options) {
  const container = document.getElementById('menuListContainer');
  const opts = options || {};
  if (!container && !opts.allowMissingContainer) return;

  if (container && !opts.silent) {
    container.innerHTML = '<div class="loading-wrap"><div class="spinner"></div></div>';
  }

  const res = await getFromGAS('getMenus');
  if (res && res.status === 'ok') {
    USER_MENUS = normalizeUserMenus(res.menus || []);

    // Add originalIndex to each menu to ensure detail modal always opens correct item
    USER_MENUS.forEach((m, idx) => {
      m.originalIndex = idx;
    });

    updateMenuCategoryFilter();
    if (container) renderMenus();
    if (document.getElementById('page-notices').classList.contains('active')) {
      renderPushNotices();
    }
  } else {
    if (container && !opts.silent) {
      container.innerHTML = '<div class="empty-state">読み込みに失敗しました</div>';
    }
  }
}

function normalizeUserMenus(items) {
  return (items || []).map(function (item) {
    const imageUrls = normalizeManagedImageList(item && (item.imageUrls || item.imageUrl));
    return {
      rowIdx: item && item.rowIdx,
      date: String(item && item.date || ''),
      updatedAt: String(item && item.updatedAt || item && item.date || ''),
      noticeListedAt: String(item && item.noticeListedAt || ''),
      name: String(item && item.name || ''),
      category: String(item && item.category || ''),
      imageUrl: imageUrls[0] || String(item && item.imageUrl || ''),
      imageUrls: imageUrls,
      description: String(item && item.description || ''),
      reservationStatus: String(item && item.reservationStatus || ''),
      noticeStatus: normalizeNoticeVisibilityStatus(item && item.noticeStatus),
      sortOrder: Number(item && item.sortOrder || 0)
    };
  }).filter(function (item) {
    return item.name;
  }).sort(function (a, b) {
    // 手動並び順が最優先
    if (a.sortOrder !== b.sortOrder) return b.sortOrder - a.sortOrder;
    // 同じ場合は更新日時順
    const timeA = new Date(a.updatedAt || a.date || 0).getTime();
    const timeB = new Date(b.updatedAt || b.date || 0).getTime();
    return timeB - timeA;
  });
}

function renderMenus() {
  const container = document.getElementById('menuListContainer');
  if (!container) return;

  const filterVal = document.getElementById('menuCategoryFilter') ? document.getElementById('menuCategoryFilter').value : '全て';

  let filteredMenus = USER_MENUS;
  if (filterVal !== '全て') {
    filteredMenus = USER_MENUS.filter(function (m) {
      return String(m.category || '').trim() === filterVal;
    });
  }

  if (filteredMenus.length === 0) {
    container.innerHTML = '<div class="empty-state">該当するメニューはありません</div>';
    return;
  }

  let html = '';
  filteredMenus.forEach(m => {
    html += `
          <div class="news-item" onclick="openMenuDetail(${m.originalIndex})" style="padding:15px; border-radius:16px; margin-bottom:15px; box-shadow:0 4px 12px rgba(0,0,0,0.05); background:#fff; border:1px solid #f0ede8; cursor:pointer; display:flex; align-items:center; gap:15px;">
            ${m.imageUrl ?
        `<img src="${getContentDisplayImageUrl(m.imageUrl)}" class="menu-list-thumb" alt="${escapeHtml(m.name || 'メニュー画像')}">` :
        `<div style="width:80px; height:80px; background:var(--bg-gray); border-radius:10px; flex-shrink:0; display:flex; align-items:center; justify-content:center; color:var(--text-light); font-size:24px;">🍴</div>`
      }
            <div style="flex:1; display:flex; align-items:center; min-width:0;">
              <div style="flex:1; min-width:0;">
                ${m.category ? `<div style="font-size:11px; color:var(--sage-dark); font-weight:700; margin-bottom:6px;">${escapeHtml(m.category)}</div>` : ''}
                <h3 style="margin:0; font-size:1.2rem; color:var(--text-main); font-weight:700; line-height:1.4; flex:1; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">${escapeHtml(m.name || '')} ${buildUnreadBadgeHtml('menu', m)} ${isFavoriteKey(buildContentItemKey('menu', m)) ? '<span class="item-favorite-badge inline">★</span>' : ''}</h3>
              </div>
            </div>
            <div style="color:var(--text-light); font-size:18px; margin-left:5px; flex-shrink:0;">›</div>
          </div>
        `;
  });
  container.innerHTML = html;
}

function openMenuDetail(idx) {
  const m = USER_MENUS[idx];
  if (!m) return;
  markContentItemSeen('menu', m);

  const imageHtml = buildDetailImageGalleryHtml(m.imageUrls || m.imageUrl, m.name || 'Menu Image', 'margin-bottom:20px;');

  const formattedDesc = renderManagedTextHtml(m.description || '', {
    inlineImageClass: 'blog-inline-image blog-inline-image--detail',
    inlineImageAlt: (m.name || 'メニュー画像')
  }) || '概要説明はありません。';

  document.getElementById('menuDetailContent').innerHTML = `
        <div style="margin-bottom:12px;">
          <span class="blog-detail-cat">${escapeHtml(m.category || 'メニュー')}</span>
        </div>
        <div class="blog-detail-title" style="margin-bottom:16px;">${escapeHtml(m.name || '')}</div>
        ${imageHtml}
        <div class="blog-detail-body managed-rich-text" style="font-size:15px; line-height:1.8; color:var(--text-main);">${formattedDesc}</div>
        <div style="margin-top:14px;">${buildFavoriteActionMarkup('menu', m)}</div>
      `;
  renderMenus();
  updateNavBadges();
  openModal('menuDetailModal');
}

/**
 * 「天然だし調味粉」専用の価格計算ロジック (アプリ用)
 */
function isDashiProductName(name) {
  const text = String(name || '').trim();
  return text === '天然だし調味粉' || text === '天然だし調理粉';
}

function calculateDashiPricing(qty) {
  const count = Math.max(0, Number(qty) || 0);
  if (count <= 0) return { totalRevenue: 0, avgUnitPrice: 2980 };

  let totalRevenue = 0;
  if (count <= 2) {
    totalRevenue = 2380 * count;
  } else if (count === 3) {
    totalRevenue = (2380 * 2) + 2235;
  } else if (count === 4) {
    totalRevenue = (2380 * 3) + 2086;
  } else if (count === 5) {
    totalRevenue = (2380 * 4) + 1937;
  } else {
    totalRevenue = 2235 * count;
  }

  return {
    totalRevenue: totalRevenue,
    avgUnitPrice: totalRevenue / count
  };
}

if (document.readyState === 'loading') { document.addEventListener('DOMContentLoaded', initApp); } else { initApp(); }
