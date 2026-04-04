// ============================================================
//  まゆみ助産院アプリ - Google Apps Script
//  【最新版】注文管理・ブログ・商品管理・カレンダー対応（LINE Messaging API対応）
// ============================================================

const CONFIG = {
  // LINE Messaging API設定
  CHANNEL_ACCESS_TOKEN: 'SET_IN_APPS_SCRIPT_PROJECT',
  USER_ID: 'SET_IN_APPS_SCRIPT_PROJECT',

  GMAIL_TO: 'josanin.mayumi@gmail.com',
  SHEET_NAME: 'まゆみ助産院_管理',
  ADMIN_NAME: 'まゆみ助産院',
  // OneSignal設定
  ONESIGNAL_APP_ID: 'SET_IN_APPS_SCRIPT_PROJECT',
  ONESIGNAL_REST_API_KEY: 'SET_IN_APPS_SCRIPT_PROJECT',

  // Google Reviews API設定
  GOOGLE_CLIENT_ID: 'SET_IN_APPS_SCRIPT_PROJECT',
  GOOGLE_CLIENT_SECRET: 'SET_IN_APPS_SCRIPT_PROJECT',
};
// ★★★★★★★★★★★★★★★★★★★★

const SHEETS = {
  ORDERS: '注文管理',
  PRODUCTS: '商品マスタ',
  BLOG: 'ブログ・お知らせ',
  MASTER: '管理マスタ',
  CALENDAR: 'カレンダー',
  USERS: '会員データ',

  CATEGORIES: 'カテゴリマスタ',
  PUSH: 'PUSH_NOTICES',
  MENUS: 'MENUS',
  MENU_REVENUE: 'MENU_REVENUE',
  PRODUCT_REVENUE: 'PRODUCT_REVENUE',
  APP_SUPPORT_FAQ: 'APP_SUPPORT_FAQ',
  BACKUP_LOG: 'BACKUP_LOG',
  ADMIN_AUDIT_LOG: 'ADMIN_AUDIT_LOG',
  SUPPORT_CHAT_LOG: 'SUPPORT_CHAT_LOG',
  ADMIN_TEMPLATES: 'ADMIN_TEMPLATES'
};

const USER_HEADERS = [
  'ID',
  '登録日時',
  '氏名',
  'フリガナ',
  '電話番号',
  'アイコンURL',
  'メモ',
  'Push設定',
  'ご状況',
  '生年月日',
  '住所',
  '現在スタンプ数',
  'スタンプカード番号',
  '特典履歴JSON',
  '最終スタンプ取得日',
  'スタンプ達成日時',
  'パスコード',
  '引き継ぎコード',
  '引き継ぎコード発行日時',
  '端末セッションJSON',
  '削除状態',
  '削除日時',
  '統合先会員ID',
  '登録経路',
  '登録経路詳細',
  '登録経路更新日時',
  'スタンプ履歴JSON',
  '最終スタンプ取得日時'
];

const USER_COL = {
  MEMBER_ID: 1,
  TIMESTAMP: 2,
  NAME: 3,
  KANA: 4,
  PHONE: 5,
  AVATAR_URL: 6,
  MEMO: 7,
  PUSH: 8,
  STATUS: 9,
  BIRTHDAY: 10,
  ADDRESS: 11,
  STAMP_COUNT: 12,
  STAMP_CARD_NUM: 13,
  REWARDS: 14,
  LAST_STAMP_DATE: 15,
  STAMP_ACHIEVED_AT: 16,
  PASSCODE: 17,
  TRANSFER_CODE: 18,
  TRANSFER_CODE_ISSUED_AT: 19,
  DEVICE_SESSIONS: 20,
  DELETE_STATUS: 21,
  DELETED_AT: 22,
  MERGED_INTO: 23,
  REGISTRATION_SOURCE: 24,
  REGISTRATION_SOURCE_DETAIL: 25,
  REGISTRATION_SOURCE_UPDATED_AT: 26,
  STAMP_HISTORY_JSON: 27,
  LAST_STAMP_AT: 28
};
const TRANSFER_CODE_LENGTH = 8;
const TRANSFER_CODE_TTL_HOURS = 168; // 1週間 (24*7)
const MAX_DEVICE_SESSIONS = 8;
const SOFT_DELETE_STATUS = '削除済み';
const SOFT_DELETE_STATUS_HEADER = '削除状態';
const SOFT_DELETE_DATE_HEADER = '削除日時';
const PUBLISH_AT_HEADER = '公開開始日時';
const ADMIN_TRASH_SOURCES = {
  BLOG: 'NEWS',
  PRODUCTS: 'ショップ',
  CALENDAR: 'カレンダー',
  MENUS: 'ホーム',
  USERS: '会員',
  ORDERS: '注文',
  PUSH: 'Push通知'
};
const BACKUP_FOLDER_NAME = 'まゆみ助産院_バックアップ';
const LAST_BACKUP_AT_PROPERTY = 'LAST_BACKUP_AT';
const LAST_BACKUP_URL_PROPERTY = 'LAST_BACKUP_URL';
const LAST_BACKUP_FILE_ID_PROPERTY = 'LAST_BACKUP_FILE_ID';
const DAILY_MAINTENANCE_TRIGGER_HANDLER = 'runDailyMaintenance';
const SCHEDULED_PUSH_TRIGGER_HANDLER = 'processScheduledPushQueue';
const PUSH_STATUS_SENT = '送信済み';
const PUSH_STATUS_AUTO_SENT = '自動送信済み';
const PUSH_STATUS_DRAFT = '下書き';
const PUSH_STATUS_SCHEDULED = '予約済み';
const PUSH_STATUS_FAILED = '送信失敗';

const MENU_REVENUE_TYPES = ['母乳外来', 'ビジリス', '教室', 'その他'];
const MENU_REVENUE_HEADERS = ['記録日', 'メニュー種別', '件数', '単価', '原価単価', 'メモ'];
const PRODUCT_REVENUE_HEADERS = ['記録日', '商品名', '個数', '単価', '原価', 'メモ'];
const DASHI_PRODUCT_NAMES = ['天然だし調味粉', '天然だし調理粉'];
const NOTICE_VISIBILITY_HEADER = 'お知らせ一覧公開';
const DASHI_BASE_PRICE = 2980;
const DASHI_COST_PER_ITEM = 1490;
const DASHI_PRICE_TIERS = [
  { key: '20', label: '2,380円(20%OFF)', price: 2380, sortOrder: 1 },
  { key: '25', label: '2,235円(25%OFF)', price: 2235, sortOrder: 2 },
  { key: '30', label: '2,086円(30%OFF)', price: 2086, sortOrder: 3 },
  { key: '35', label: '1,937円(35%OFF)', price: 1937, sortOrder: 4 }
];
const REWARD_GACHA_PRIZE_POOL = [
  { key: 'A', rankLabel: 'A賞', rewardName: 'A賞プレゼント', capsuleColor: '#f5cb6c', accentColor: '#b0791b', message: '受付でその時の特典をお受け取りください。', weight: 5 },
  { key: 'B', rankLabel: 'B賞', rewardName: 'B賞プレゼント', capsuleColor: '#f3b7c9', accentColor: '#b86282', message: 'まゆみ助産院からのうれしいごほうびです。受付でご案内します。', weight: 15 },
  { key: 'C', rankLabel: 'C賞', rewardName: 'C賞プレゼント', capsuleColor: '#b9d8a7', accentColor: '#628f58', message: 'やさしいプレゼントをご用意しています。受付へお声がけください。', weight: 30 },
  { key: 'D', rankLabel: 'D賞', rewardName: 'D賞プレゼント', capsuleColor: '#b9d9f3', accentColor: '#547fa2', message: '当日のおたのしみプレゼントは受付でご確認ください。', weight: 50 }
];
const FALLBACK_SPREADSHEET_ID = '1gIcUGxg2PEuFoU5a_IgQ6lDWgghceJ7v2dgqo9iPe4w';
const DATA_CACHE_VERSION_PROPERTY = 'DATA_CACHE_VERSION';
const DATA_CACHE_TTL_SEC = 120;
const DATA_CACHE_MAX_CHARS = 90000;
const APP_RUNTIME_CONFIG_PROPERTY = 'APP_RUNTIME_CONFIG';
const ADMIN_SECURITY_CONFIG_PROPERTY = 'ADMIN_SECURITY_CONFIG';
const DEFAULT_APP_RUNTIME_CONFIG = {
  latestAppVersion: '1.1.0',
  minimumSupportedVersion: '0.0.0',
  iosStoreUrl: '',
  updateTitle: 'アップデートが必要です',
  updateMessage: 'このアプリを引き続き利用するには、最新版へアップデートしてください。',
  webBundleVersion: '2026.04.04.59'
};
const DEFAULT_ADMIN_SECURITY_CONFIG = {
  roleMode: 'owner',
  actionPin: '',
  operatorName: '管理者'
};
const STALE_PENDING_ORDER_DAYS = 3;
const STALE_UNUSED_REWARD_DAYS = 30;
const STALE_UNPUBLISHED_SCHEDULE_HOURS = 24;

const MUTATING_GET_ACTIONS = {
  order: true,
  cancel: true,
  confirmReceipt: true
};

const NON_INVALIDATING_POST_TYPES = {
  askSupportChat: true,
  uploadImage: true,
  postGoogleReviewReply: true
};

let spreadsheetCache_ = null;

function getDataCacheVersion_() {
  const props = PropertiesService.getScriptProperties();
  let version = props.getProperty(DATA_CACHE_VERSION_PROPERTY);
  if (!version) {
    version = String(Date.now());
    props.setProperty(DATA_CACHE_VERSION_PROPERTY, version);
  }
  return version;
}

function bumpDataCacheVersion_() {
  PropertiesService.getScriptProperties().setProperty(DATA_CACHE_VERSION_PROPERTY, String(Date.now()));
}

function buildDataCacheKey_(scope, payload) {
  const raw = [getDataCacheVersion_(), scope || '', payload || ''].join('::');
  const digest = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw)
  ).replace(/=+$/g, '');
  return 'mayumi-cache:' + digest;
}

function getCachedData_(scope, payload) {
  try {
    const cached = CacheService.getScriptCache().get(buildDataCacheKey_(scope, payload));
    return cached ? JSON.parse(cached) : null;
  } catch (err) {
    Logger.log('getCachedData_ error: ' + err.toString());
    return null;
  }
}

function putCachedData_(scope, payload, data, ttlSec) {
  try {
    const text = JSON.stringify(data);
    if (!text || text.length > DATA_CACHE_MAX_CHARS) return;
    CacheService.getScriptCache().put(buildDataCacheKey_(scope, payload), text, ttlSec || DATA_CACHE_TTL_SEC);
  } catch (err) {
    Logger.log('putCachedData_ error: ' + err.toString());
  }
}

function versionParts_(value) {
  const parts = String(value || '')
    .split(/[^0-9]+/)
    .filter(Boolean)
    .map(function (part) { return parseInt(part, 10) || 0; });
  return parts.length ? parts : [0];
}

function compareVersions_(left, right) {
  const a = versionParts_(left);
  const b = versionParts_(right);
  const maxLength = Math.max(a.length, b.length);
  for (let i = 0; i < maxLength; i++) {
    const leftPart = a[i] || 0;
    const rightPart = b[i] || 0;
    if (leftPart > rightPart) return 1;
    if (leftPart < rightPart) return -1;
  }
  return 0;
}

function sanitizeAppRuntimeConfig_(raw) {
  const input = raw && typeof raw === 'object' ? raw : {};
  const config = {
    latestAppVersion: String(input.latestAppVersion || DEFAULT_APP_RUNTIME_CONFIG.latestAppVersion).trim() || DEFAULT_APP_RUNTIME_CONFIG.latestAppVersion,
    minimumSupportedVersion: String(input.minimumSupportedVersion || DEFAULT_APP_RUNTIME_CONFIG.minimumSupportedVersion).trim() || DEFAULT_APP_RUNTIME_CONFIG.minimumSupportedVersion,
    iosStoreUrl: String(input.iosStoreUrl || DEFAULT_APP_RUNTIME_CONFIG.iosStoreUrl).trim(),
    updateTitle: String(input.updateTitle || DEFAULT_APP_RUNTIME_CONFIG.updateTitle).trim() || DEFAULT_APP_RUNTIME_CONFIG.updateTitle,
    updateMessage: String(input.updateMessage || DEFAULT_APP_RUNTIME_CONFIG.updateMessage).trim() || DEFAULT_APP_RUNTIME_CONFIG.updateMessage,
    // 現在配布中のネイティブ版との整合を優先し、返却値は固定で揃える
    webBundleVersion: DEFAULT_APP_RUNTIME_CONFIG.webBundleVersion
  };

  if (compareVersions_(config.latestAppVersion, config.minimumSupportedVersion) < 0) {
    config.latestAppVersion = config.minimumSupportedVersion;
  }

  if (config.iosStoreUrl && !/^https?:\/\//i.test(config.iosStoreUrl)) {
    config.iosStoreUrl = '';
  }

  return config;
}

function getAppRuntimeConfig() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(APP_RUNTIME_CONFIG_PROPERTY);
    let saved = null;
    if (raw) {
      try {
        saved = JSON.parse(raw);
      } catch (err) {
        Logger.log('getAppRuntimeConfig parse error: ' + err.toString());
      }
    }
    return {
      status: 'ok',
      config: sanitizeAppRuntimeConfig_(saved)
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleSaveAppRuntimeConfig(data) {
  try {
    const nextConfig = sanitizeAppRuntimeConfig_((data && data.config) || data || {});
    if (compareVersions_(nextConfig.minimumSupportedVersion, '0.0.0') > 0 && !nextConfig.iosStoreUrl) {
      return { status: 'error', message: '最低対応バージョンを設定する場合は App Store URL も必要です' };
    }
    PropertiesService.getScriptProperties().setProperty(APP_RUNTIME_CONFIG_PROPERTY, JSON.stringify(nextConfig));
    return {
      status: 'ok',
      config: nextConfig
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function sanitizeAdminSecurityConfig_(raw) {
  const input = raw && typeof raw === 'object' ? raw : {};
  const roleMode = ['viewer', 'editor', 'owner'].indexOf(String(input.roleMode || '').trim()) !== -1
    ? String(input.roleMode || '').trim()
    : DEFAULT_ADMIN_SECURITY_CONFIG.roleMode;
  const actionPin = String(input.actionPin || '').trim().replace(/[^0-9]/g, '').slice(0, 8);
  return {
    roleMode: roleMode,
    actionPin: actionPin,
    operatorName: String(input.operatorName || DEFAULT_ADMIN_SECURITY_CONFIG.operatorName).trim() || DEFAULT_ADMIN_SECURITY_CONFIG.operatorName
  };
}

function getAdminSecurityConfig() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(ADMIN_SECURITY_CONFIG_PROPERTY);
    let saved = null;
    if (raw) {
      try {
        saved = JSON.parse(raw);
      } catch (err) {
        Logger.log('getAdminSecurityConfig parse error: ' + err.toString());
      }
    }
    return {
      status: 'ok',
      config: sanitizeAdminSecurityConfig_(saved)
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleSaveAdminSecurityConfig(data) {
  try {
    const nextConfig = sanitizeAdminSecurityConfig_((data && data.config) || data || {});
    PropertiesService.getScriptProperties().setProperty(ADMIN_SECURITY_CONFIG_PROPERTY, JSON.stringify(nextConfig));
    return { status: 'ok', config: nextConfig };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function ensureAdminAuditLogSheet_() {
  const ss = getOrCreateSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.ADMIN_AUDIT_LOG);
  if (!sheet) sheet = ss.insertSheet(SHEETS.ADMIN_AUDIT_LOG);
  const headers = ['日時', '種別', '結果', '対象', '概要', '操作者', '詳細JSON'];
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet, headers.length, '#455a64');
  [170, 150, 90, 180, 320, 140, 420].forEach(function (width, index) {
    sheet.setColumnWidth(index + 1, width);
  });
  return sheet;
}

function safeJsonStringify_(value) {
  try {
    return JSON.stringify(value);
  } catch (err) {
    return '';
  }
}

function summarizeAuditTarget_(type, data, result) {
  if (!type) return '';
  if (type === 'deleteOrders' || type === 'updateOrder' || type === 'updateAdminOrder' || type === 'createOrder') {
    return String((data && (data.orderId || (Array.isArray(data.orderIds) ? data.orderIds[0] : ''))) || '').trim();
  }
  if (type === 'deleteUser' || type === 'updateUser' || type === 'updateAdminUser' || type === 'mergeUsers') {
    return String((data && (data.memberId || data.targetMemberId || data.name)) || '').trim();
  }
  if (type === 'addProduct' || type === 'updateProduct') return String(data && data.name || '').trim();
  if (type === 'addBlog' || type === 'updateBlog') return String(data && data.title || '').trim();
  if (type === 'addCalendar' || type === 'updateCalendar') return String(data && data.title || '').trim();
  if (type === 'broadcastPush') return String(data && data.title || '').trim();
  if (type === 'saveSupportFaq' || type === 'deleteSupportFaq') return String(data && data.question || data && data.rowIdx || '').trim();
  if (type === 'saveAdminTemplate' || type === 'deleteAdminTemplate') return String(data && data.title || data && data.templateId || '').trim();
  return String(result && result.message || '').trim();
}

function summarizeAuditDescription_(type, data, result) {
  if (type === 'broadcastPush') {
    const mode = String(data && data.mode || 'send');
    const count = Number(result && result.recipientCount || 0);
    return [mode, data && data.targetStatus, count ? ('対象 ' + count + '件') : ''].filter(Boolean).join(' / ');
  }
  if (type === 'mergeUsers') {
    return '統合元 ' + ((data && data.sourceMemberIds && data.sourceMemberIds.length) || 0) + '件';
  }
  if (type === 'deleteOrders') {
    return '削除 ' + ((data && data.orderIds && data.orderIds.length) || 0) + '件';
  }
  return String(result && result.message || '').trim();
}

function appendAdminAuditLog_(type, data, result) {
  try {
    const sheet = ensureAdminAuditLogSheet_();
    const securityConfig = getAdminSecurityConfig();
    const operatorName = securityConfig && securityConfig.status === 'ok'
      ? securityConfig.config.operatorName
      : DEFAULT_ADMIN_SECURITY_CONFIG.operatorName;
    sheet.appendRow([
      formatDateTime_(new Date()),
      String(type || '').trim(),
      result && result.status === 'ok' ? '成功' : '失敗',
      summarizeAuditTarget_(type, data, result),
      summarizeAuditDescription_(type, data, result),
      operatorName,
      safeJsonStringify_({
        request: data || {},
        result: result || {}
      })
    ]);
  } catch (err) {
    Logger.log('appendAdminAuditLog_ error: ' + err.toString());
  }
}

function getAdminAuditLogs() {
  try {
    const sheet = ensureAdminAuditLogSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', logs: [] };
    const values = sheet.getRange(2, 1, lastRow - 1, 7).getDisplayValues();
    const logs = values.map(function (row, index) {
      return {
        rowIdx: index + 2,
        timestamp: row[0] || '',
        type: row[1] || '',
        result: row[2] || '',
        target: row[3] || '',
        summary: row[4] || '',
        operator: row[5] || '',
        detailJson: row[6] || ''
      };
    }).reverse();
    return { status: 'ok', logs: logs };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function ensureSupportChatLogSheet_() {
  const ss = getOrCreateSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.SUPPORT_CHAT_LOG);
  if (!sheet) sheet = ss.insertSheet(SHEETS.SUPPORT_CHAT_LOG);
  const headers = ['日時', '質問', '一致質問', '一致カテゴリ', '一致有無', '会員ID', '回答要約'];
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet, headers.length, '#5d4037');
  [170, 300, 260, 160, 90, 120, 320].forEach(function (width, index) {
    sheet.setColumnWidth(index + 1, width);
  });
  return sheet;
}

function appendSupportChatLog_(payload) {
  try {
    const sheet = ensureSupportChatLogSheet_();
    sheet.appendRow([
      formatDateTime_(new Date()),
      String(payload && payload.message || '').trim(),
      String(payload && payload.matchedQuestion || '').trim(),
      String(payload && payload.category || '').trim(),
      payload && payload.matchedQuestion ? '一致' : '未一致',
      String(payload && payload.memberId || '').trim(),
      String(payload && payload.answer || '').trim().slice(0, 500)
    ]);
  } catch (err) {
    Logger.log('appendSupportChatLog_ error: ' + err.toString());
  }
}

function getSupportChatAnalytics() {
  try {
    const sheet = ensureSupportChatLogSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { status: 'ok', summary: { total: 0, matched: 0, unmatched: 0 }, topQuestions: [], unmatchedExamples: [] };
    }
    const values = sheet.getRange(2, 1, lastRow - 1, 7).getDisplayValues();
    const topMap = {};
    const unmatchedExamples = [];
    let matched = 0;
    values.forEach(function (row) {
      const question = String(row[1] || '').trim();
      const matchedQuestion = String(row[2] || '').trim();
      const isMatched = String(row[4] || '').trim() === '一致';
      if (isMatched) matched++;
      if (question) {
        topMap[question] = (topMap[question] || 0) + 1;
      }
      if (!isMatched && question && unmatchedExamples.length < 10) {
        unmatchedExamples.push({
          askedAt: row[0] || '',
          question: question
        });
      }
    });
    const topQuestions = Object.keys(topMap).map(function (question) {
      return { question: question, count: topMap[question] };
    }).sort(function (a, b) {
      return b.count - a.count;
    }).slice(0, 10);
    return {
      status: 'ok',
      summary: {
        total: values.length,
        matched: matched,
        unmatched: values.length - matched
      },
      topQuestions: topQuestions,
      unmatchedExamples: unmatchedExamples
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function ensureAdminTemplatesSheet_() {
  const ss = getOrCreateSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.ADMIN_TEMPLATES);
  if (!sheet) sheet = ss.insertSheet(SHEETS.ADMIN_TEMPLATES);
  const headers = ['テンプレートID', '種別', 'タイトル', '本文', '更新日時'];
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet, headers.length, '#8d6e63');
  [140, 120, 220, 360, 170].forEach(function (width, index) {
    sheet.setColumnWidth(index + 1, width);
  });
  return sheet;
}

function getAdminTemplates() {
  try {
    const sheet = ensureAdminTemplatesSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', templates: [] };
    const values = sheet.getRange(2, 1, lastRow - 1, 5).getDisplayValues();
    const templates = values.map(function (row, index) {
      return {
        rowIdx: index + 2,
        templateId: row[0] || '',
        kind: row[1] || '',
        title: row[2] || '',
        body: row[3] || '',
        updatedAt: row[4] || ''
      };
    }).filter(function (item) { return item.templateId; }).reverse();
    return { status: 'ok', templates: templates };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleSaveAdminTemplate(data) {
  try {
    const kind = String(data && data.kind || '').trim() || '汎用';
    const title = String(data && data.title || '').trim();
    const body = String(data && data.body || '').trim();
    if (!title) return { status: 'error', message: 'テンプレート名を入力してください' };
    if (!body) return { status: 'error', message: '本文を入力してください' };
    const sheet = ensureAdminTemplatesSheet_();
    const rowIdx = Number(data && data.rowIdx || 0);
    const templateId = rowIdx > 1
      ? String(sheet.getRange(rowIdx, 1).getDisplayValue() || '').trim()
      : 'TPL-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const row = [templateId, kind, title, body, formatDateTime_(new Date())];
    if (rowIdx > 1) {
      sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return { status: 'ok', templateId: templateId };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleDeleteAdminTemplate(data) {
  try {
    const rowIdx = Number(data && data.rowIdx || 0);
    if (rowIdx <= 1) return { status: 'error', message: '削除対象が見つかりません' };
    const sheet = ensureAdminTemplatesSheet_();
    sheet.deleteRow(rowIdx);
    return { status: 'ok' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function ensureTimeTriggerExists_(handlerName, builderFn) {
  const triggers = getProjectTriggersSafe_();
  if (!triggers) return false;
  const exists = triggers.some(function (trigger) {
    return trigger && trigger.getHandlerFunction && trigger.getHandlerFunction() === handlerName;
  });
  if (!exists && typeof builderFn === 'function') {
    builderFn();
  }
  return true;
}

function ensureMaintenanceTriggers_() {
  const dailyReady = ensureTimeTriggerExists_(DAILY_MAINTENANCE_TRIGGER_HANDLER, function () {
    ScriptApp.newTrigger(DAILY_MAINTENANCE_TRIGGER_HANDLER).timeBased().everyDays(1).atHour(3).create();
  });
  const pushReady = ensureTimeTriggerExists_(SCHEDULED_PUSH_TRIGGER_HANDLER, function () {
    ScriptApp.newTrigger(SCHEDULED_PUSH_TRIGGER_HANDLER).timeBased().everyMinutes(5).create();
  });
  return { dailyReady: !!dailyReady, pushReady: !!pushReady };
}

function getProjectTriggersSafe_() {
  try {
    return ScriptApp.getProjectTriggers();
  } catch (err) {
    Logger.log('getProjectTriggersSafe_ skipped: ' + err.toString());
    return null;
  }
}

function getBackupFolder_() {
  const folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

function createSpreadsheetBackup_(reason) {
  const ss = getOrCreateSpreadsheet();
  const folder = getBackupFolder_();
  const file = DriveApp.getFileById(ss.getId());
  const label = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const copy = file.makeCopy(CONFIG.SHEET_NAME + '_backup_' + label, folder);
  const url = copy.getUrl();
  const props = PropertiesService.getScriptProperties();
  props.setProperty(LAST_BACKUP_AT_PROPERTY, formatDateTime_(new Date()));
  props.setProperty(LAST_BACKUP_URL_PROPERTY, url);
  props.setProperty(LAST_BACKUP_FILE_ID_PROPERTY, copy.getId());

  const logSheet = getOrCreateBackupLogSheet_(ss);
  logSheet.appendRow([
    formatDateTime_(new Date()),
    String(reason || 'auto').trim() || 'auto',
    copy.getName(),
    copy.getId(),
    url
  ]);
  return {
    createdAt: formatDateTime_(new Date()),
    fileId: copy.getId(),
    url: url,
    name: copy.getName()
  };
}

function getOrCreateBackupLogSheet_(ss) {
  let sheet = ss.getSheetByName(SHEETS.BACKUP_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.BACKUP_LOG);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['作成日時', '種別', 'ファイル名', 'ファイルID', 'URL']);
    styleHeader(sheet, 5, '#546e7a');
  }
  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 280);
  sheet.setColumnWidth(4, 190);
  sheet.setColumnWidth(5, 320);
  return sheet;
}

function getRecentBackupLogEntries_() {
  const ss = getOrCreateSpreadsheet();
  const sheet = getOrCreateBackupLogSheet_(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 5).getDisplayValues()
    .map(function (row) {
      return {
        createdAt: row[0] || '',
        reason: row[1] || '',
        fileName: row[2] || '',
        fileId: row[3] || '',
        url: row[4] || ''
      };
    })
    .reverse()
    .slice(0, 10);
}

function getBackupStatus() {
  try {
    const triggerState = ensureMaintenanceTriggers_();
    const props = PropertiesService.getScriptProperties();
    const triggers = getProjectTriggersSafe_();
    const recentLogs = getRecentBackupLogEntries_();
    const lastBackupAt = String(props.getProperty(LAST_BACKUP_AT_PROPERTY) || '');
    const staleThreshold = Date.now() - (1000 * 60 * 60 * 36);
    return {
      status: 'ok',
      lastBackupAt: lastBackupAt,
      lastBackupUrl: String(props.getProperty(LAST_BACKUP_URL_PROPERTY) || ''),
      hasDailyTrigger: !!(triggers && triggers.some(function (trigger) {
        return trigger.getHandlerFunction() === DAILY_MAINTENANCE_TRIGGER_HANDLER;
      })),
      hasPushSchedulerTrigger: !!(triggers && triggers.some(function (trigger) {
        return trigger.getHandlerFunction() === SCHEDULED_PUSH_TRIGGER_HANDLER;
      })),
      triggerPermissionAvailable: !!triggers,
      triggerSetupAttempted: !!(triggerState.dailyReady || triggerState.pushReady),
      isBackupStale: !lastBackupAt || parseLooseDateToTimestamp_(lastBackupAt) < staleThreshold,
      recentLogs: recentLogs
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleRunManualBackup(data) {
  try {
    return {
      status: 'ok',
      backup: createSpreadsheetBackup_(data && data.reason ? data.reason : 'manual')
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function runDailyMaintenance() {
  ensureMaintenanceTriggers_();
  createSpreadsheetBackup_('daily');
  processScheduledPushQueue();
}

function getPendingOrderSummary_() {
  const orders = getAdminOrders({ showAll: true });
  if (!orders || orders.status !== 'ok') return { total: 0, pending: 0, received: 0 };
  const list = Array.isArray(orders.orders) ? orders.orders : [];
  return {
    total: list.length,
    pending: list.filter(function (order) { return normalizeOrderStatus_(order.status) === '受付中'; }).length,
    received: list.filter(function (order) { return normalizeOrderStatus_(order.status) === '受取済'; }).length
  };
}

function getProductInventoryAlerts_() {
  const products = getAdminProducts();
  if (!products || products.status !== 'ok') return [];
  return (products.products || []).filter(function (product) {
    const stockQty = Number(product.stockQty || 0);
    const lowStockThreshold = Number(product.lowStockThreshold || 0);
    return stockQty <= 0 || (lowStockThreshold > 0 && stockQty <= lowStockThreshold);
  }).slice(0, 10);
}

function getPublishSchedule() {
  try {
    const nowTs = Date.now();
    const items = [];
    const collect = function (source, records, getLabel) {
      (records || []).forEach(function (record) {
        const publishAt = String(record.publishAt || '').trim();
        if (!publishAt) return;
        const ts = parseLooseDateToTimestamp_(publishAt);
        if (!ts || ts <= nowTs) return;
        items.push({
          source: source,
          title: getLabel(record),
          publishAt: publishAt,
          rowIdx: Number(record.rowIdx || 0)
        });
      });
    };
    const blogs = getAdminBlogs();
    const products = getAdminProducts();
    const calendar = getAdminCalendar();
    const menus = getAdminMenus();
    collect('NEWS', blogs && blogs.blogs, function (item) { return item.title || 'NEWS'; });
    collect('ショップ', products && products.products, function (item) { return item.name || '商品'; });
    collect('カレンダー', calendar && calendar.events, function (item) { return item.title || 'イベント'; });
    collect('ホーム', menus && menus.menus, function (item) { return item.name || 'メニュー'; });
    items.sort(function (a, b) {
      return parseLooseDateToTimestamp_(a.publishAt) - parseLooseDateToTimestamp_(b.publishAt);
    });
    return { status: 'ok', items: items };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getStalePendingOrders_(days) {
  const thresholdDays = Math.max(1, Number(days) || STALE_PENDING_ORDER_DAYS);
  const cutoff = Date.now() - thresholdDays * 24 * 60 * 60 * 1000;
  const orders = getAdminOrders({ showAll: true });
  if (!orders || orders.status !== 'ok') return [];
  return (orders.orders || []).filter(function (order) {
    return normalizeOrderStatus_(order.status) === '受付中' && parseLooseDateToTimestamp_(order.date) > 0 && parseLooseDateToTimestamp_(order.date) <= cutoff;
  }).sort(function (a, b) {
    return parseLooseDateToTimestamp_(a.date) - parseLooseDateToTimestamp_(b.date);
  });
}

function getStaleUnusedRewards_(days) {
  const thresholdDays = Math.max(1, Number(days) || STALE_UNUSED_REWARD_DAYS);
  const cutoff = Date.now() - thresholdDays * 24 * 60 * 60 * 1000;
  const users = getAdminUsers();
  if (!users || users.status !== 'ok') return [];
  const items = [];
  (users.users || []).forEach(function (user) {
    (user.rewards || []).forEach(function (reward) {
      if (reward && reward.used) return;
      const earnedTs = parseLooseDateToTimestamp_(reward && reward.earnedDate);
      if (!earnedTs || earnedTs > cutoff) return;
      items.push({
        memberId: user.memberId,
        name: user.name,
        rewardName: String(reward && reward.rewardName || '特典'),
        earnedDate: formatMaybeDateTime_(reward && reward.earnedDate)
      });
    });
  });
  return items.sort(function (a, b) {
    return parseLooseDateToTimestamp_(a.earnedDate) - parseLooseDateToTimestamp_(b.earnedDate);
  });
}

function collectOverduePublishItemsFromSheet_(sheet, sourceLabel, titleCol, statusCol, overdueBeforeTs) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const publishAtCol = ensurePublishAtColumn_(sheet);
  const deleteCols = ensureSoftDeleteColumns_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return values.map(function (row, index) {
    if (isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol)) return null;
    const publishAt = formatMaybeDateTime_(row[publishAtCol - 1]);
    const publishAtTs = parseLooseDateToTimestamp_(publishAt);
    if (!publishAtTs || publishAtTs > overdueBeforeTs) return null;
    if (String(row[(statusCol || 0) - 1] || '').trim() === '公開') return null;
    return {
      source: sourceLabel,
      title: String(row[(titleCol || 0) - 1] || '').trim() || sourceLabel,
      publishAt: publishAt,
      rowIdx: index + 2
    };
  }).filter(Boolean).sort(function (a, b) {
    return parseLooseDateToTimestamp_(a.publishAt) - parseLooseDateToTimestamp_(b.publishAt);
  });
}

function getOverdueScheduledPublishes_(hours) {
  const thresholdHours = Math.max(1, Number(hours) || STALE_UNPUBLISHED_SCHEDULE_HOURS);
  const overdueBeforeTs = Date.now() - thresholdHours * 60 * 60 * 1000;
  const ss = getOrCreateSpreadsheet();
  return []
    .concat(collectOverduePublishItemsFromSheet_(ss.getSheetByName(SHEETS.BLOG), 'NEWS', 2, 6, overdueBeforeTs))
    .concat(collectOverduePublishItemsFromSheet_(ss.getSheetByName(SHEETS.PRODUCTS), 'ショップ', 2, 6, overdueBeforeTs))
    .concat(collectOverduePublishItemsFromSheet_(ss.getSheetByName(SHEETS.CALENDAR), 'カレンダー', 2, 5, overdueBeforeTs))
    .concat(collectOverduePublishItemsFromSheet_(ensureMenusSheetStructure_(ss.getSheetByName(SHEETS.MENUS)), 'ホーム', 2, 6, overdueBeforeTs));
}

function getAdminDashboardData() {
  try {
    const ordersRes = getAdminOrders({ showAll: true });
    const orderList = ordersRes && ordersRes.status === 'ok' ? (ordersRes.orders || []) : [];
    const users = getAdminUsers();
    const userList = users && users.status === 'ok' ? (users.users || []) : [];
    const orders = {
      total: orderList.length,
      pending: orderList.filter(function (order) { return normalizeOrderStatus_(order.status) === '受付中'; }).length,
      received: orderList.filter(function (order) { return normalizeOrderStatus_(order.status) === '受取済'; }).length
    };
    const stalePendingCutoff = Date.now() - STALE_PENDING_ORDER_DAYS * 24 * 60 * 60 * 1000;
    const stalePendingOrders = orderList.filter(function (order) {
      return normalizeOrderStatus_(order.status) === '受付中'
        && parseLooseDateToTimestamp_(order.date) > 0
        && parseLooseDateToTimestamp_(order.date) <= stalePendingCutoff;
    }).sort(function (a, b) {
      return parseLooseDateToTimestamp_(a.date) - parseLooseDateToTimestamp_(b.date);
    });
    const staleRewardCutoff = Date.now() - STALE_UNUSED_REWARD_DAYS * 24 * 60 * 60 * 1000;
    const staleUnusedRewards = [];
    userList.forEach(function (user) {
      (user.rewards || []).forEach(function (reward) {
        if (reward && reward.used) return;
        const earnedTs = parseLooseDateToTimestamp_(reward && reward.earnedDate);
        if (!earnedTs || earnedTs > staleRewardCutoff) return;
        staleUnusedRewards.push({
          memberId: user.memberId,
          name: user.name,
          rewardName: String(reward && reward.rewardName || '特典'),
          earnedDate: formatMaybeDateTime_(reward && reward.earnedDate)
        });
      });
    });
    staleUnusedRewards.sort(function (a, b) {
      return parseLooseDateToTimestamp_(a.earnedDate) - parseLooseDateToTimestamp_(b.earnedDate);
    });
    const duplicateGroups = buildDuplicateUsersFromRows_(userList.map(function (user) {
      return {
        rowIdx: user.rowIdx,
        memberId: user.memberId,
        name: user.name,
        phone: user.phone,
        birthday: user.birthday,
        updatedAt: user.timestamp,
        stampCount: user.stampCount,
        orderCount: user.orderCount,
        pendingOrderCount: user.pendingOrderCount,
        lastOrderAt: user.lastOrderAt,
        orderTotal: user.orderTotal,
        deviceCount: user.deviceCount
      };
    }));
    const backup = getBackupStatus();
    const schedule = getPublishSchedule();
    const push = getPushNotices();
    const products = getAdminProducts();
    const overduePublishes = getOverdueScheduledPublishes_(STALE_UNPUBLISHED_SCHEDULE_HOURS);
    const alerts = [];
    if (backup && backup.status === 'ok' && backup.isBackupStale) {
      alerts.push({ level: 'warning', label: 'バックアップ未更新', detail: '36時間以上バックアップが更新されていません' });
    }
    if (backup && backup.status === 'ok' && !backup.hasDailyTrigger) {
      alerts.push({ level: 'warning', label: '日次バックアップ', detail: '日次トリガーが未設定です' });
    }
    if (orders.pending > 0) {
      alerts.push({ level: 'info', label: '未対応注文', detail: orders.pending + '件の受付中注文があります' });
    }
    const failedPush = push && push.status === 'ok'
      ? (push.notices || []).filter(function (notice) { return String(notice.status || '').trim() === PUSH_STATUS_FAILED; }).length
      : 0;
    if (failedPush > 0) {
      alerts.push({ level: 'warning', label: 'Push失敗', detail: failedPush + '件の送信失敗があります' });
    }
    const lowStockProducts = getProductInventoryAlerts_();
    if (lowStockProducts.length) {
      alerts.push({ level: 'warning', label: '在庫警告', detail: lowStockProducts.length + '件の商品で在庫警告があります' });
    }
    const duplicateCount = duplicateGroups.length;
    if (duplicateCount > 0) {
      alerts.push({ level: 'warning', label: '重複会員候補', detail: duplicateCount + '組の重複候補があります' });
    }
    if (stalePendingOrders.length > 0) {
      alerts.push({
        level: 'warning',
        label: '長期間未対応注文',
        detail: stalePendingOrders.length + '件の受付中注文が' + STALE_PENDING_ORDER_DAYS + '日以上経過しています'
      });
    }
    if (staleUnusedRewards.length > 0) {
      alerts.push({
        level: 'warning',
        label: '未受取特典',
        detail: staleUnusedRewards.length + '件の特典が' + STALE_UNUSED_REWARD_DAYS + '日以上未受取です'
      });
    }
    if (overduePublishes.length > 0) {
      alerts.push({
        level: 'warning',
        label: '未公開予約',
        detail: overduePublishes.length + '件の公開予約が予定時刻を過ぎても非公開のままです'
      });
    }
    return {
      status: 'ok',
      summary: {
        users: userList.length,
        ordersPending: orders.pending,
        products: products && products.status === 'ok' ? (products.products || []).length : 0,
        scheduledPublishes: schedule && schedule.status === 'ok' ? (schedule.items || []).length : 0,
        failedPushes: failedPush,
        duplicateGroups: duplicateCount
      },
      alerts: alerts,
      lowStockProducts: lowStockProducts,
      stalePendingOrders: stalePendingOrders.slice(0, 10),
      staleUnusedRewards: staleUnusedRewards.slice(0, 10),
      overduePublishes: overduePublishes.slice(0, 10),
      recentBackups: backup && backup.status === 'ok' ? (backup.recentLogs || []) : [],
      upcomingPublishes: schedule && schedule.status === 'ok' ? (schedule.items || []).slice(0, 12) : []
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function buildSheetHeaderMap_(sheet) {
  const header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  const map = {};
  header.forEach(function (label, index) {
    map[String(label || '').trim()] = index + 1;
  });
  return map;
}

function ensureSoftDeleteColumns_(sheet) {
  if (!sheet) return { statusCol: 0, deletedAtCol: 0 };
  const statusCol = ensureNamedColumn_(sheet, SOFT_DELETE_STATUS_HEADER, 110);
  const deletedAtCol = ensureNamedColumn_(sheet, SOFT_DELETE_DATE_HEADER, 170);
  return { statusCol: statusCol, deletedAtCol: deletedAtCol };
}

function ensurePublishAtColumn_(sheet) {
  if (!sheet) return 0;
  return ensureNamedColumn_(sheet, PUBLISH_AT_HEADER, 170);
}

function isSoftDeletedByColumns_(row, statusCol, deletedAtCol) {
  const status = String(row[(statusCol || 0) - 1] || '').trim();
  const deletedAt = String(row[(deletedAtCol || 0) - 1] || '').trim();
  return status === SOFT_DELETE_STATUS || !!deletedAt;
}

function markRowSoftDeleted_(sheet, rowIdx, reason) {
  if (!sheet || rowIdx <= 1) return false;
  const cols = ensureSoftDeleteColumns_(sheet);
  sheet.getRange(rowIdx, cols.statusCol).setValue(SOFT_DELETE_STATUS);
  sheet.getRange(rowIdx, cols.deletedAtCol).setValue(formatDateTime_(new Date()));
  if (reason) {
    const reasonCol = ensureNamedColumn_(sheet, '削除理由', 180);
    sheet.getRange(rowIdx, reasonCol).setValue(String(reason));
  }
  return true;
}

function clearRowSoftDeleted_(sheet, rowIdx) {
  if (!sheet || rowIdx <= 1) return false;
  const cols = ensureSoftDeleteColumns_(sheet);
  sheet.getRange(rowIdx, cols.statusCol).setValue('');
  sheet.getRange(rowIdx, cols.deletedAtCol).setValue('');
  const headerMap = buildSheetHeaderMap_(sheet);
  const mergedIntoCol = headerMap['統合先会員ID'] || 0;
  if (mergedIntoCol) sheet.getRange(rowIdx, mergedIntoCol).setValue('');
  return true;
}

function normalizePublishAtValue_(value) {
  return normalizeDateTimeValue_(value);
}

function isPublishAtAvailable_(publishAtValue) {
  const normalized = normalizePublishAtValue_(publishAtValue);
  if (!normalized) return true;
  const publishAt = new Date(normalized);
  if (isNaN(publishAt.getTime())) return true;
  return publishAt.getTime() <= Date.now();
}

function getLogicalSheetName_(sheetKey) {
  if (sheetKey === 'BLOG') return SHEETS.BLOG;
  if (sheetKey === 'PRODUCTS') return SHEETS.PRODUCTS;
  if (sheetKey === 'CALENDAR') return SHEETS.CALENDAR;
  if (sheetKey === 'PUSH') return SHEETS.PUSH;
  if (sheetKey === 'MENUS') return SHEETS.MENUS;
  if (sheetKey === 'USERS') return SHEETS.USERS;
  if (sheetKey === 'ORDERS') return SHEETS.ORDERS;
  return '';
}

function getSheetByLogicalKey_(ss, sheetKey) {
  const name = getLogicalSheetName_(sheetKey);
  return name ? ss.getSheetByName(name) : null;
}

function getTrashRowLabel_(sheetKey, row, headerMap) {
  if (sheetKey === 'USERS') return String(row[USER_COL.NAME - 1] || row[USER_COL.MEMBER_ID - 1] || '会員');
  if (sheetKey === 'ORDERS') return String(row[0] || '注文');
  if (sheetKey === 'BLOG') return String(row[1] || 'NEWS');
  if (sheetKey === 'PRODUCTS') return String(row[1] || '商品');
  if (sheetKey === 'CALENDAR') return String(row[1] || 'イベント');
  if (sheetKey === 'MENUS') return String(row[1] || 'メニュー');
  if (sheetKey === 'PUSH') return String(row[1] || '通知');
  return '';
}

function collectDeletedRowsFromSheet_(sheetKey) {
  const ss = getOrCreateSpreadsheet();
  const sheet = getSheetByLogicalKey_(ss, sheetKey);
  if (!sheet || sheet.getLastRow() < 2) return [];
  if (sheetKey === 'USERS') ensureUsersSheetStructure_(sheet);
  const cols = ensureSoftDeleteColumns_(sheet);
  const headerMap = buildSheetHeaderMap_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return values.map(function (row, index) {
    if (!isSoftDeletedByColumns_(row, cols.statusCol, cols.deletedAtCol)) return null;
    return {
      sheet: sheetKey,
      source: ADMIN_TRASH_SOURCES[sheetKey] || sheetKey,
      rowIdx: index + 2,
      title: getTrashRowLabel_(sheetKey, row, headerMap),
      deletedAt: formatMaybeDateTime_(row[cols.deletedAtCol - 1]),
      reason: headerMap['削除理由'] ? String(row[headerMap['削除理由'] - 1] || '') : ''
    };
  }).filter(Boolean);
}

function getAdminTrashItems() {
  try {
    const items = []
      .concat(collectDeletedRowsFromSheet_('USERS'))
      .concat(collectDeletedRowsFromSheet_('ORDERS'))
      .concat(collectDeletedRowsFromSheet_('BLOG'))
      .concat(collectDeletedRowsFromSheet_('PRODUCTS'))
      .concat(collectDeletedRowsFromSheet_('CALENDAR'))
      .concat(collectDeletedRowsFromSheet_('MENUS'))
      .concat(collectDeletedRowsFromSheet_('PUSH'))
      .sort(function (a, b) {
        return parseLooseDateToTimestamp_(b.deletedAt) - parseLooseDateToTimestamp_(a.deletedAt);
      });
    return { status: 'ok', items: items };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleRestoreDeletedRecord(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = getSheetByLogicalKey_(ss, data.sheet);
    const rowIdx = Number(data.rowIdx || 0);
    if (!sheet || rowIdx <= 1) return { status: 'error', message: '復元対象が見つかりません' };
    if (data.sheet === 'USERS') ensureUsersSheetStructure_(sheet);
    clearRowSoftDeleted_(sheet, rowIdx);
    return { status: 'ok' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleHardDeleteRecord(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = getSheetByLogicalKey_(ss, data.sheet);
    const rowIdx = Number(data.rowIdx || 0);
    if (!sheet || rowIdx <= 1) return { status: 'error', message: '削除対象が見つかりません' };
    sheet.deleteRow(rowIdx);
    return { status: 'ok' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function parseLooseDateToTimestamp_(value) {
  const normalized = normalizeDateTimeValue_(value);
  if (!normalized) return 0;
  const parsed = new Date(normalized);
  return isNaN(parsed.getTime()) ? 0 : parsed.getTime();
}

function ensurePushNoticeSheetStructure_() {
  const ss = getOrCreateSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.PUSH);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.PUSH);
  }
  const headers = ['日時', 'タイトル', '本文', '送信対象', '送信対象詳細', '送信件数', 'ステータス', '配信予定日時', '対象ページ', 'プレビュー本文', '通知ID', '送信結果', '更新日時'];
  const maxCols = sheet.getMaxColumns();
  if (maxCols < headers.length) {
    sheet.insertColumnsAfter(maxCols, headers.length - maxCols);
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet, headers.length, '#6a5acd');
  [170, 220, 320, 120, 260, 110, 120, 170, 120, 320, 160, 220, 170].forEach(function (width, index) {
    sheet.setColumnWidth(index + 1, width);
  });
  return {
    sheet: sheet,
    columns: {
      sentAt: 1,
      title: 2,
      body: 3,
      targetStatus: 4,
      targetDetail: 5,
      recipientCount: 6,
      status: 7,
      scheduledAt: 8,
      targetPage: 9,
      previewBody: 10,
      notificationId: 11,
      result: 12,
      updatedAt: 13
    }
  };
}

function normalizeTargetPage_(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (['home', 'shop', 'calendar', 'news', 'notices', 'mypage', 'menu-list'].indexOf(normalized) !== -1) {
    return normalized;
  }
  return 'home';
}

function buildAppNotificationUrl_(targetPage) {
  const page = normalizeTargetPage_(targetPage);
  return 'https://mayumi-josanin.github.io/mayumi-app/?open=' + encodeURIComponent(page);
}

function buildPushNoticeRowFromInput_(data, status) {
  const targetUsers = Array.isArray(data && data.targetUsers) ? data.targetUsers : [];
  return {
    sentAt: normalizeDateTimeValue_(data && data.sentAt),
    title: String(data && data.title || '').trim(),
    body: String(data && data.body || '').trim(),
    targetStatus: String(data && data.targetStatus || 'all').trim() || 'all',
    targetDetail: String(data && data.targetDetail || (targetUsers.length ? safeJsonStringify_(targetUsers) : '')).trim(),
    recipientCount: Number(data && data.recipientCount || targetUsers.length || 0),
    status: String(status || data && data.status || PUSH_STATUS_DRAFT).trim() || PUSH_STATUS_DRAFT,
    scheduledAt: normalizeDateTimeValue_(data && data.scheduledAt),
    targetPage: normalizeTargetPage_(data && data.targetPage),
    previewBody: String(data && data.previewBody || data && data.body || '').trim(),
    notificationId: String(data && data.notificationId || '').trim(),
    result: String(data && data.result || '').trim(),
    updatedAt: formatDateTime_(new Date())
  };
}

function writePushNoticeRow_(sheet, rowIdx, columns, rowData) {
  const payload = new Array(Math.max(sheet.getLastColumn(), columns.updatedAt)).fill('');
  payload[columns.sentAt - 1] = rowData.sentAt || '';
  payload[columns.title - 1] = rowData.title || '';
  payload[columns.body - 1] = rowData.body || '';
  payload[columns.targetStatus - 1] = rowData.targetStatus || 'all';
  payload[columns.targetDetail - 1] = rowData.targetDetail || '';
  payload[columns.recipientCount - 1] = Number(rowData.recipientCount || 0);
  payload[columns.status - 1] = rowData.status || PUSH_STATUS_DRAFT;
  payload[columns.scheduledAt - 1] = rowData.scheduledAt || '';
  payload[columns.targetPage - 1] = rowData.targetPage || 'home';
  payload[columns.previewBody - 1] = rowData.previewBody || '';
  payload[columns.notificationId - 1] = rowData.notificationId || '';
  payload[columns.result - 1] = rowData.result || '';
  payload[columns.updatedAt - 1] = rowData.updatedAt || formatDateTime_(new Date());
  if (rowIdx && rowIdx > 1) {
    sheet.getRange(rowIdx, 1, 1, payload.length).setValues([payload]);
    return rowIdx;
  }
  sheet.appendRow(payload);
  return sheet.getLastRow();
}

function parsePushTargetUsers_(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value;
  try {
    const parsed = JSON.parse(String(value));
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}

function buildOneSignalPayload_(record) {
  const payload = {
    app_id: CONFIG.ONESIGNAL_APP_ID,
    contents: { en: record.body, ja: record.body },
    headings: { en: record.title, ja: record.title },
    url: buildAppNotificationUrl_(record.targetPage),
    data: {
      openPage: normalizeTargetPage_(record.targetPage)
    }
  };
  const targetUsers = parsePushTargetUsers_(record.targetUsers || record.targetDetail);
  const subscriptionIds = targetUsers
    .map(function (user) { return String(user && (user.subscription || user.subscriptionId) || '').trim(); })
    .filter(Boolean);
  if (subscriptionIds.length) {
    payload.include_subscription_ids = subscriptionIds;
  } else if (record.targetStatus && record.targetStatus !== 'all') {
    payload.filters = [
      { field: 'tag', key: 'status', relation: '=', value: record.targetStatus }
    ];
  } else {
    payload.included_segments = ['All'];
  }
  return payload;
}

function sendPushNoticePayload_(record) {
  const response = UrlFetchApp.fetch('https://onesignal.com/api/v1/notifications', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Basic ' + CONFIG.ONESIGNAL_REST_API_KEY
    },
    payload: JSON.stringify(buildOneSignalPayload_(record)),
    muteHttpExceptions: true
  });
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  if (responseCode !== 200) {
    throw new Error('OneSignal送信失敗: ' + responseText);
  }
  let parsed = {};
  try {
    parsed = JSON.parse(responseText || '{}');
  } catch (err) { }
  return {
    id: String(parsed.id || ''),
    raw: responseText
  };
}

function processScheduledPushQueue() {
  try {
    const ensured = ensurePushNoticeSheetStructure_();
    const sheet = ensured.sheet;
    const columns = ensured.columns;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', processed: 0 };

    const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    let processed = 0;
    rows.forEach(function (row, index) {
      const rowIdx = index + 2;
      const status = String(row[columns.status - 1] || '').trim();
      if (status !== PUSH_STATUS_SCHEDULED) return;
      const scheduledAt = row[columns.scheduledAt - 1];
      if (!isPublishAtAvailable_(scheduledAt)) return;
      const record = buildPushNoticeRowFromInput_({
        title: row[columns.title - 1],
        body: row[columns.body - 1],
        targetStatus: row[columns.targetStatus - 1],
        targetDetail: row[columns.targetDetail - 1],
        recipientCount: row[columns.recipientCount - 1],
        targetPage: row[columns.targetPage - 1],
        previewBody: row[columns.previewBody - 1]
      }, PUSH_STATUS_SCHEDULED);
      try {
        const result = sendPushNoticePayload_(record);
        record.sentAt = formatDateTime_(new Date());
        record.status = PUSH_STATUS_SENT;
        record.notificationId = result.id;
        record.result = '予約送信済み';
      } catch (err) {
        record.sentAt = formatDateTime_(new Date());
        record.status = PUSH_STATUS_FAILED;
        record.result = err.toString();
      }
      writePushNoticeRow_(sheet, rowIdx, columns, record);
      processed++;
    });
    return { status: 'ok', processed: processed };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// ========== GET：アプリからのデータ取得 ==========

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  const cachePayload = JSON.stringify((e && e.parameter) || {});
  const isMutatingAction = !!MUTATING_GET_ACTIONS[action];
  const canUseCache = !!action && !isMutatingAction;

  try {
    if (canUseCache) {
      const cached = getCachedData_(action, cachePayload);
      if (cached) {
        return createJsonResponse(cached);
      }
    }

    let result;
    switch (action) {
      case 'getNews':
        result = getBlogNews();
        break;
      case 'getProducts':
        result = getProducts();
        break;
      case 'getCalendar':
        result = getCalendarEvents();
        break;
      case 'getAdminOrders':
        result = getAdminOrders(e.parameter.data ? JSON.parse(decodeURIComponent(e.parameter.data)) : null);
        break;
      case 'getAdminUserOrders':
        result = getAdminUserOrders(e.parameter.data ? JSON.parse(decodeURIComponent(e.parameter.data)) : null);
        break;
      case 'getAdminBlogs':
        result = getAdminBlogs();
        break;
      case 'getAdminProducts':
        result = getAdminProducts();
        break;
      case 'getAdminCalendar':
        result = getAdminCalendar();
        break;
      case 'getAdminUsers':
        result = getAdminUsers();
        break;
      case 'getPushUsers':
        result = getPushUsers();
        break;
      case 'getPushNotices':
        result = getPushNotices();
        break;
      case 'getInitialData':
        result = getInitialData();
        break;
      case 'getAppRuntimeConfig':
        result = getAppRuntimeConfig();
        break;
      case 'getAdminSecurityConfig':
        result = getAdminSecurityConfig();
        break;
      case 'getCustomerOrders':
        result = getCustomerOrders(e.parameter.data ? JSON.parse(decodeURIComponent(e.parameter.data)) : null);
        break;
      case 'getAnalytics':
        result = getAnalyticsData();
        break;
      case 'getAdminDashboardData':
        result = getAdminDashboardData();
        break;
      case 'getPublishSchedule':
        result = getPublishSchedule();
        break;
      case 'getMenuRevenueRecords':
        result = getMenuRevenueRecords();
        break;
      case 'getProductRevenueRecords':
        result = getProductRevenueRecords();
        break;
      case 'getUserRewardStatus':
        result = getUserRewardStatus(e.parameter.data ? JSON.parse(decodeURIComponent(e.parameter.data)) : null);
        break;

      case 'getMenus':
        result = getMenus();
        break;
      case 'getAdminMenus':
        result = getAdminMenus();
        break;
      case 'getCategories':
        result = getCategories();
        break;
      case 'getSupportFaq':
        result = getSupportFaq();
        break;
      case 'getAdminSupportFaq':
        result = getAdminSupportFaq();
        break;
      case 'getRecoveryCandidates':
        result = getRecoveryCandidates(e.parameter.data ? JSON.parse(decodeURIComponent(e.parameter.data)) : null);
        break;
      case 'getDuplicateUsers':
        result = getDuplicateUsers();
        break;
      case 'getAdminTrashItems':
        result = getAdminTrashItems();
        break;
      case 'getAdminAuditLogs':
        result = getAdminAuditLogs();
        break;
      case 'getBackupStatus':
        result = getBackupStatus();
        break;
      case 'getSupportChatAnalytics':
        result = getSupportChatAnalytics();
        break;
      case 'getAdminTemplates':
        result = getAdminTemplates();
        break;
      case 'getUserDevices':
        result = getUserDevices(e.parameter.data ? JSON.parse(decodeURIComponent(e.parameter.data)) : null);
        break;
      case 'order':
        result = handleOrder(JSON.parse(decodeURIComponent(e.parameter.data)));
        break;
      case 'cancel':
        result = handleCancel(JSON.parse(decodeURIComponent(e.parameter.data)));
        break;
      case 'confirmReceipt':
        result = handleConfirmReceipt(JSON.parse(decodeURIComponent(e.parameter.data)));
        break;

      default:
        result = { status: 'error', message: '未定義のGETアクションです: ' + action };
        break;
    }

    if (result && result.status === 'ok') {
      if (isMutatingAction) {
        bumpDataCacheVersion_();
      } else if (canUseCache) {
        putCachedData_(action, cachePayload, result, DATA_CACHE_TTL_SEC);
      }
    }

    return createJsonResponse(result);
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========== POST：注文受信 ==========

// ========== POST：注文受信 ==========

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const type = data.type;
    let result = {};

    switch (type) {
      case 'order':
        result = handleOrder(data);
        break;
      case 'cancel':
        result = handleCancel(data);
        break;
      case 'updateProduct':
        result = handleUpdateProduct(data);
        break;
      case 'updateBlog':
        result = handleUpdateBlog(data);
        break;
      case 'updateCalendar':
        result = handleUpdateCalendar(data);
        break;
      case 'updateRecordStatus':
        result = handleUpdateRecordStatus(data);
        break;
      case 'updateNoticeVisibility':
        result = handleUpdateNoticeVisibility(data);
        break;
      case 'addBlog':
        result = handleAddBlog(data);
        break;
      case 'addCalendar':
        result = handleAddCalendar(data);
        break;
      case 'addProduct':
        result = handleAddProduct(data);
        break;
      case 'addCategory':
        result = handleAddCategory(data);
        break;
      case 'updateCategory':
        result = handleUpdateCategory(data);
        break;
      case 'deleteCategory':
        result = handleDeleteCategory(data);
        break;
      case 'saveSupportFaq':
        result = handleSaveSupportFaq(data);
        break;
      case 'deleteSupportFaq':
        result = handleDeleteSupportFaq(data);
        break;
      case 'askSupportChat':
        result = askSupportChat(data);
        break;
      case 'uploadImage':
        result = handleUploadImage(data);
        break;
      case 'deleteRow':
        result = handleDeleteRow(data);
        break;
      case 'deleteRows':
        result = handleDeleteRows(data);
        break;
      case 'deleteOrders':
        result = handleDeleteOrders(data);
        break;
      case 'updateOrder':
        result = handleUpdateOrder(data);
        break;
      case 'updateItemOrder':
        result = handleUpdateItemOrder(data);
        break;
      case 'confirmReceipt':
        result = handleConfirmReceipt(data);
        break;
      case 'recoverAccount':
        result = handleRecoverAccount(data);
        break;
      case 'issueTransferCode':
        result = handleIssueTransferCode(data);
        break;
      case 'resetForgottenPasscode':
        result = handleResetForgottenPasscode(data);
        break;
      case 'updateUser':
        result = handleUpdateUser(data);
        break;
      case 'syncUserDeviceSession':
        result = handleSyncUserDeviceSession(data);
        break;
      case 'removeUserDeviceSession':
        result = handleRemoveUserDeviceSession(data);
        break;
      case 'unsubscribePush':
        result = handleUnsubscribePush(data);
        break;
      case 'syncUserRewardStatus':
        result = handleSyncUserRewardStatus(data);
        break;
      case 'drawRewardGacha':
        result = handleDrawRewardGacha(data);
        break;
      case 'broadcastPush':
        result = broadcastPush(data);
        break;
      case 'postGoogleReviewReply':
        result = postGoogleReviewReply(data);
        break;

      case 'resetUsers':
        const ssReset = getOrCreateSpreadsheet();
        const sheetReset = ssReset.getSheetByName(SHEETS.USERS);
        if (sheetReset) {
          const lastRow = sheetReset.getLastRow();
          if (lastRow > 1) sheetReset.deleteRows(2, lastRow - 1);
        }
        result = { status: 'ok', message: '会員情報をリセットしました' };
        break;
      case 'deleteUser':
        result = handleDeleteUser(data);
        break;
      case 'mergeUsers':
        result = handleMergeUsers(data);
        break;
      case 'restoreDeletedRecord':
        result = handleRestoreDeletedRecord(data);
        break;
      case 'hardDeleteRecord':
        result = handleHardDeleteRecord(data);
        break;
      case 'updateAdminUser':
        result = handleUpdateAdminUser(data);
        break;
      case 'updateAdminRewardStatus':
        result = handleUpdateAdminRewardStatus(data);
        break;
      case 'saveAppRuntimeConfig':
        result = handleSaveAppRuntimeConfig(data);
        break;
      case 'saveAdminSecurityConfig':
        result = handleSaveAdminSecurityConfig(data);
        break;
      case 'runManualBackup':
        result = handleRunManualBackup(data);
        break;
      case 'saveMenuRevenueRecord':
        result = handleSaveMenuRevenueRecord(data);
        break;
      case 'deleteMenuRevenueRecord':
        result = handleDeleteMenuRevenueRecord(data);
        break;
      case 'saveProductRevenueRecord':
        result = handleSaveProductRevenueRecord(data);
        break;
      case 'deleteProductRevenueRecord':
        result = handleDeleteProductRevenueRecord(data);
        break;
      case 'saveAdminTemplate':
        result = handleSaveAdminTemplate(data);
        break;
      case 'deleteAdminTemplate':
        result = handleDeleteAdminTemplate(data);
        break;
      case 'addMenu':
        result = handleAddMenu(data);
        break;
      case 'updateMenu':
        result = handleUpdateMenu(data);
        break;
      case 'deleteMenu':
        result = handleDeleteMenu(data);
        break;
      case 'moveMenu':
        result = handleMoveMenu(data);
        break;
      case 'createOrder':
        result = createOrder(data);
        break;
      case 'updateAdminOrder':
        result = updateAdminOrder(data);
        break;

      default:
        result = { status: 'error', message: '未定義のPOSTアクションです: ' + type };
    }
    if (result && result.status === 'ok' && !NON_INVALIDATING_POST_TYPES[type]) {
      bumpDataCacheVersion_();
    }
    if (type !== 'askSupportChat' && type !== 'uploadImage' && type !== 'postGoogleReviewReply') {
      appendAdminAuditLog_(type, data, result);
    }
    return createJsonResponse(result);
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

// ========== 注文キャンセル処理 ==========
// 同じ注文IDの全行のステータスを「キャンセル済」に更新

function handleCancel(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.ORDERS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', updated: 0 };

    const values = sheet.getRange(2, 1, lastRow - 1, 13).getValues(); // A~M列対応
    let updated = 0;
    values.forEach(function (row, i) {
      if (String(row[0]).trim() === String(data.orderId).trim()) {
        sheet.getRange(i + 2, 12).setValue('キャンセル済'); // L列：ステータス
        sheet.getRange(i + 2, 5).setValue(0); // E列：個数を0にしてカウントから削除
        updated++;
      }
    });

    Logger.log('キャンセル反映: ' + updated + '行 / orderId: ' + data.orderId);
    return { status: 'ok', updated: updated };
  } catch (err) {
    Logger.log('handleCancel error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function normalizeOrderStatus_(status) {
  const value = String(status || '').trim();
  if (!value || value === '未受取' || value === '受付中' || value === 'pending') return '受付中';
  if (value === '受取済' || value.indexOf('受取済') !== -1 || value === 'done') return '受取済';
  if (value === 'キャンセル済' || value.indexOf('キャンセル') !== -1 || value === 'cancelled') return 'キャンセル済';
  return value;
}



// ========== 注文受付処理 ==========
// 商品ごとに1行ずつ記録します

function handleOrder(data) {
  const ss = getOrCreateSpreadsheet();
  const sheet = ensureOrdersSheetStructure_(ss.getSheetByName(SHEETS.ORDERS));
  const usersSheet = getOrCreateUsersSheet_(ss);

  const now = data.time || getCurrentTime();

  // 商品ごとに1行ずつスプレッドシートに追記
  const lastRow = sheet.getLastRow();
  const orderId = String(data.orderId || '').trim();
  if (!orderId) return { status: 'error', message: 'orderId is required' };

  ensureUserRowFromActivity_(usersSheet, {
    memberId: data.memberId,
    name: data.name || data.customerName,
    phone: data.phone,
    birthday: data.birthday,
    address: data.address
  });

  if (lastRow >= 2) {
    const existingOrderIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const alreadyExists = existingOrderIds.some(function (row) {
      return String(row[0] || '').trim() === orderId;
    });
    if (alreadyExists) {
      return { status: 'ok', orderId: orderId, duplicate: true };
    }
  }

  const startRow = lastRow + 1;

  data.items.forEach(function (item, index) {
    const r = startRow + index;

    let recordedPrice = item.price;
    // 天然だしは個数に応じた価格帯で記録
    if (isDashiProductName_(item.name)) {
       const dashiPrice = calculateDashiPricing_(item.qty);
       recordedPrice = dashiPrice.avgUnitPrice;
    }

    sheet.appendRow([
      orderId,               // A: 注文ID
      now,                   // B: 注文日時
      data.customerName,     // C: 注文者名
      item.name,             // D: 商品名
      item.qty,              // E: 個数
      recordedPrice,         // F: 単価
      `=IFERROR(VLOOKUP(D${r}, '${SHEETS.MASTER}'!A:B, 2, FALSE), 0)`, // G: 仕入値（数式で管理マスタから取得）
      `=E${r}*F${r}`,        // H: 小計（数式）
      `=E${r}*(F${r}-G${r})`,// I: 純利益（数式）
      index === 0 ? '¥' + Number(data.total).toLocaleString() : '', // J: 合計金額
      index === 0 ? data.payment : '',  // K: 支払方法
      index === 0 ? '受付中' : '',      // L: ステータス
      '',                    // M: 受取確認
      data.memberId || '',   // N: 会員ID
      index === 0 ? String(data.internalNote || '') : '' // O: 管理メモ
    ]);
    sheet.getRange(r, 13).insertCheckboxes(); // M列にチェックボックスを挿入
  });

  // 注文内容のテキスト（通知用）
  const itemLines = data.items.map(function (i) {
    const lineTotal = isDashiProductName_(i.name)
      ? calculateDashiPricing_(i.qty).totalRevenue
      : Number(i.price * i.qty);
    return '  ・' + i.name + '　×' + i.qty + '個　¥' + Number(lineTotal).toLocaleString();
  }).join('\n');

  // LINE通知メッセージ作成
  const lineMsg = '🛍 新しいご注文が届きました！\n'
    + '\n👤 注文者：' + data.customerName
    + '\n━━━━━━━━━━━━━━━'
    + '\n📦 注文商品：\n' + itemLines
    + '\n━━━━━━━━━━━━━━━'
    + '\n💴 合計：¥' + Number(data.total).toLocaleString()
    + '\n💳 支払：' + data.payment
    + '\n🕐 時刻：' + now;

  // LINE Messaging APIで送信
  sendLineMessage(lineMsg);

  // Gmail通知
  const gmailBody =
    'まゆみ助産院アプリ - 新しいご注文が届きました\n\n'
    + '━━━━━━━━━━━━━━━━━━━━━━\n'
    + '注文ID　：' + orderId + '\n'
    + '注文者　：' + data.customerName + '\n'
    + '注文日時：' + now + '\n'
    + '支払方法：' + data.payment + '\n'
    + '合計金額：¥' + Number(data.total).toLocaleString() + '\n'
    + '━━━━━━━━━━━━━━━━━━━━━━\n\n'
    + '【注文商品】\n' + itemLines + '\n\n'
    + '━━━━━━━━━━━━━━━━━━━━━━\n'
    + 'スプレッドシートでご確認ください。';
  sendGmail(
    '🛍 新しいご注文 - ' + data.customerName + '様（¥' + Number(data.total).toLocaleString() + '）',
    gmailBody
  );

  return { status: 'ok', orderId: orderId };
}

// ========== 管理者用：新規注文作成 ==========
function createOrder(data) {
  const ss = getOrCreateSpreadsheet();
  const sheet = ensureOrdersSheetStructure_(ss.getSheetByName(SHEETS.ORDERS));
  if (!sheet) return { status: 'error', message: 'ORDERS sheet not found' };
  const usersSheet = getOrCreateUsersSheet_(ss);

  const now = data.date || getCurrentTime();
  const lastRow = sheet.getLastRow();
  const startRow = lastRow + 1;
  const nextOrderId = 'ADM-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '-' + startRow;

  if (!data.items || data.items.length === 0) {
    return { status: 'error', message: 'No items provided' };
  }

  ensureUserRowFromActivity_(usersSheet, {
    memberId: data.memberId,
    name: data.name || data.customerName,
    phone: data.phone,
    birthday: data.birthday,
    address: data.address
  });

  data.items.forEach(function (item, index) {
    const r = startRow + index;

    let recordedPrice = item.price || 0;
    if (isDashiProductName_(item.name)) {
       const dashiPrice = calculateDashiPricing_(item.qty);
       recordedPrice = dashiPrice.avgUnitPrice;
    }

    sheet.appendRow([
      nextOrderId,             // A: 注文ID
      now,                     // B: 日時
      data.customerName || '', // C: 名前
      item.name || '',         // D: 商品名
      item.qty || 1,           // E: 個数
      recordedPrice,           // F: 単価
      `=IFERROR(VLOOKUP(D${r}, '${SHEETS.MASTER}'!A:B, 2, FALSE), 0)`, // G: 仕入値
      `=E${r}*F${r}`,        // H: 小計
      `=E${r}*(F${r}-G${r})`,// I: 純利益
      index === 0 ? '¥' + Number(data.total).toLocaleString() : '', // J: 合計金額
      '手動入力',              // K: 支払方法
      normalizeOrderStatus_(data.status || '受付中'),   // L: ステータス
      data.checked || false,   // M: 受取確認
      data.memberId || '',     // N: 会員ID
      index === 0 ? String(data.internalNote || '') : '' // O: 管理メモ
    ]);
    sheet.getRange(r, 13).insertCheckboxes();
    if (data.checked) sheet.getRange(r, 13).setValue(true);
  });

  return { status: 'ok', orderId: nextOrderId };
}

// ========== 管理者用：注文詳細更新 ==========
function updateAdminOrder(data) {
  const ss = getOrCreateSpreadsheet();
  const sheet = ensureOrdersSheetStructure_(ss.getSheetByName(SHEETS.ORDERS));
  if (!sheet) return { status: 'error', message: 'ORDERS sheet not found' };

  const targetId = String(data.orderId).trim();
  if (!targetId) return { status: 'error', message: 'Order ID is required' };

  // 既存の注文行をすべて削除してから再登録する（複数商品対応のため）
  const allData = sheet.getDataRange().getValues();
  // 削除対象の行番号を後ろから特定
  const rowsToDelete = [];
  for (let i = allData.length - 1; i >= 1; i--) {
    if (String(allData[i][0]).trim() === targetId) {
      rowsToDelete.push(i + 1);
    }
  }

  if (rowsToDelete.length === 0) return { status: 'error', message: 'Order not found' };

  // 1行目の情報を元に基本データを保持（もしdataにない場合）
  const baseRow = allData[rowsToDelete[0] - 1]; 
  const finalDate = data.date || baseRow[1];
  const finalCustomer = data.customerName || baseRow[2];
  const finalStatus = normalizeOrderStatus_(data.status || baseRow[11]);
  const finalChecked = data.checked !== undefined ? data.checked : baseRow[12];
  const finalMemberId = data.memberId !== undefined ? data.memberId : baseRow[13];
  const finalInternalNote = data.internalNote !== undefined ? data.internalNote : baseRow[14];

  // 削除
  rowsToDelete.forEach(row => sheet.deleteRow(row));

  // 「受取済」チェックがあれば再追加せず終了（実質的な削除）
  if (finalChecked === true || String(finalChecked) === 'true') {
    return { status: 'ok', deleted: true };
  }

  // 再作成 (createOrderとほぼ同じロジックだがIDは固定)
  const lastRow = sheet.getLastRow();
  const startRow = lastRow + 1;

  data.items.forEach(function (item, index) {
    const r = startRow + index;
    sheet.appendRow([
      targetId,
      finalDate,
      finalCustomer,
      item.name,
      item.qty,
      item.price,
      `=IFERROR(VLOOKUP(D${r}, '${SHEETS.MASTER}'!A:B, 2, FALSE), 0)`,
      `=E${r}*F${r}`,
      `=E${r}*(F${r}-G${r})`,
      index === 0 ? '¥' + Number(data.total).toLocaleString() : '',
      '手動修正', // 支払方法は固定
      finalStatus,
      finalChecked,
      finalMemberId,
      index === 0 ? String(finalInternalNote || '') : ''
    ]);
    sheet.getRange(r, 13).insertCheckboxes();
    if (finalChecked) sheet.getRange(r, 13).setValue(true);
  });

  return { status: 'ok' };
}

// 行ごとの簡易更新（ステータス・チェックのみ）
function handleUpdateOrder(data) {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ORDERS);
  if (!sheet) return { status: 'error', message: 'ORDERS sheet not found' };

  const targetId = String(data.orderId).trim();
  if (!targetId) return { status: 'error', message: 'Order ID is required' };

  const allData = sheet.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = allData.length - 1; i >= 1; i--) {
    if (String(allData[i][0]).trim() === targetId) {
      rowsToDelete.push(i + 1);
    }
  }

  if (rowsToDelete.length === 0) return { status: 'error', message: 'Order not found' };

  // チェックが入っている場合は削除
  if (data.checked === true || String(data.checked) === 'true') {
    rowsToDelete.forEach(row => sheet.deleteRow(row));
    return { status: 'ok', deleted: true };
  }

  // チェックがない場合はステータスとチェック列を更新
  rowsToDelete.forEach(row => {
    if (data.status) sheet.getRange(row, 12).setValue(normalizeOrderStatus_(data.status)); // L列: ステータス
    sheet.getRange(row, 13).setValue(false); // M列: チェック
  });

  return { status: 'ok' };
}

// ========== ブログ・お知らせの取得 ==========

function getBlogImageUrlFromRow_(row, imageCol, body) {
  const storedImageUrl = imageCol ? String(row[imageCol - 1] || '').trim() : '';
  if (storedImageUrl) return storedImageUrl;
  const embeddedImageMatch = String(body || '').match(/📷\s*(https?:\/\/\S+)/);
  return embeddedImageMatch ? embeddedImageMatch[1] : '';
}

function getBlogNews() {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.BLOG);
  if (!sheet) return { status: 'ok', news: [] };
  ensureUpdatedAtColumn_(sheet, '更新日時');
  const noticeCol = ensureNoticeVisibilityColumn_(sheet, 6, '公開');
  const deleteCols = ensureSoftDeleteColumns_(sheet);
  const publishAtCol = ensurePublishAtColumn_(sheet);
  const imageCol = ensureNamedColumn_(sheet, '画像URL', 220);

  const data = sheet.getDataRange().getValues();
  const news = [];
  const catsRes = getCategories();
  const categoryTypeMap = {};
  (catsRes.categories || []).forEach(function (item) {
    const name = String(item && item.name || '').trim();
    if (!name) return;
    categoryTypeMap[name] = item.type === 'お知らせ' ? 'お知らせ' : 'ブログ';
  });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;
    if (isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol)) continue;
    if (normalizePublishVisibilityStatus_(row[noticeCol - 1] || row[5] || '公開') === '非公開') continue;
    if (!isPublishAtAvailable_(row[publishAtCol - 1])) continue;

    const dateVal = row[0];
    let dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy.MM.dd');
    } else {
      dateStr = String(dateVal).trim();
    }

    const category = String(row[2] || 'お知らせ');
    const body = String(row[4] || '');
    const imageUrl = getBlogImageUrlFromRow_(row, imageCol, body);
    news.push({
      date: dateStr,
      title: String(row[1]),
      category: category,
      type: categoryTypeMap[category] || (category === 'お知らせ' || category === '休診情報' ? 'お知らせ' : 'ブログ'),
      icon: String(row[3] || '📢'),
      body: body,
      image: imageUrl,
      imageUrl: imageUrl,
      updatedAt: formatMaybeDateTime_(row[6]),
    });
  }

  // 日付降順
  news.sort(function (a, b) { return b.date.localeCompare(a.date); });

  return {
    status: 'ok',
    news: news,
    categories: catsRes.categories || []
  };
}

// ========== カレンダーの取得 ==========

function getCalendarEvents() {
  const ss = getOrCreateSpreadsheet();
  let sheet = ss.getSheetByName('カレンダー');
  if (!sheet) {
    sheet = ss.insertSheet('カレンダー');
    sheet.getRange(1, 1, 1, 7).setValues([[
      '日付（例：2025-06-15）', 'イベント名', '詳細', 'カラー（色名またはコード）', '公開設定', '画像URL', '更新日時'
    ]]);
    sheet.getRange(1, 1, 1, 7).setBackground('#e57373').setFontColor('#ffffff').setFontWeight('bold');

    // 現在の月に合わせたサンプルデータを登録
    const today = new Date();
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, '0');
    const updatedAt = formatDateTime_(new Date());
    const calEvents = [
      [`${y}-${m}-15`, '📝 産後ヨガ教室', '10:00〜11:30 定員5名様', '#f48fb1', '公開', '', updatedAt],
      [`${y}-${m}-22`, '休診日', '臨時休診となります', '#e57373', '公開', '', updatedAt],
      [`${y}-${m}-25`, '🌿 お灸イベント', 'ご自宅でできるお灸のやり方', '#81c784', '公開', '', updatedAt]
    ];
    sheet.getRange(2, 1, calEvents.length, 7).setValues(calEvents);
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 300);
    sheet.setColumnWidth(4, 150);
    sheet.setColumnWidth(5, 90);

    sheet.getRange('G1').setValue('【カラーの例】').setFontColor('#888888').setFontSize(10);
    sheet.getRange('G2').setValue('#e57373 (休診などに)').setFontColor('#888888').setFontSize(10);
    sheet.getRange('G3').setValue('#f48fb1 (ヨガなどに)').setFontColor('#888888').setFontSize(10);
    sheet.getRange('G4').setValue('#81c784 (お灸などに)').setFontColor('#888888').setFontSize(10);
    sheet.getRange('G5').setValue('#ffb74d (イベントなどに)').setFontColor('#888888').setFontSize(10);
    sheet.setColumnWidth(7, 250);
  }
  ensureUpdatedAtColumn_(sheet, '更新日時');

  const data = sheet.getDataRange().getValues();
  const events = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || !row[1] || String(row[4]).trim() === '非公開') continue;

    const dateVal = row[0];
    let dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy-MM-dd');
    } else {
      // 文字列フォーマット(例：2026/2/1など)のゼロ埋め対応
      const strVal = String(dateVal).trim().replace(/\//g, '-');
      const parts = strVal.split('-');
      if (parts.length === 3) {
        const py = parts[0];
        const pm = String(parts[1]).padStart(2, '0');
        const pd = String(parts[2]).padStart(2, '0');
        dateStr = `${py}-${pm}-${pd}`;
      } else {
        dateStr = strVal;
      }
    }

    events.push({
      date: dateStr,
      title: String(row[1]),
      desc: String(row[2] || ''),
      color: String(row[3] || '#e57373'),
      image: String(row[5] || ''),
      updatedAt: formatMaybeDateTime_(row[6]),
    });
  }

  return { status: 'ok', events: events };
}

// ========== 商品マスタの取得（将来拡張用）==========

function getProducts() {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PRODUCTS);
  if (!sheet) return { status: 'ok', products: [] };
  ensureProductSheetStructure_(sheet);
  const noticeCol = ensureNoticeVisibilityColumn_(sheet, 6, '公開');
  const deleteCols = ensureSoftDeleteColumns_(sheet);
  const publishAtCol = ensurePublishAtColumn_(sheet);

  const data = sheet.getDataRange().getValues();
  const products = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1] || String(row[5]).trim() === '非公開') continue;
    if (isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol)) continue;
    if (!isPublishAtAvailable_(row[publishAtCol - 1])) continue;
    const stockQty = Number(row[9] || 0);
    const lowStockThreshold = Number(row[10] || 0);
    products.push({
      category: String(row[0]),
      name: String(row[1]),
      price: Number(row[2]),
      icon: String(row[3] || '🌿'),
      bg: String(row[4] || 'c1'),
      description: String(row[6] || ''),
      descriptionImage: String(row[7] || ''),
      updatedAt: formatMaybeDateTime_(row[8]),
      stockQty: stockQty,
      lowStockThreshold: lowStockThreshold,
      isSoldOut: stockQty <= 0,
      isLowStock: lowStockThreshold > 0 && stockQty > 0 && stockQty <= lowStockThreshold,
      publishAt: formatMaybeDateTime_(row[publishAtCol - 1]),
      noticeStatus: normalizePublishVisibilityStatus_(row[noticeCol - 1] || row[5] || '公開')
    });
  }
  return { status: 'ok', products: products };
}

// ========== LINE Messaging API 送信 ==========

function sendLineMessage(message) {
  if (!CONFIG.CHANNEL_ACCESS_TOKEN || !CONFIG.USER_ID) {
    Logger.log('LINE設定が不足しています');
    return;
  }

  const url = 'https://api.line.me/v2/bot/message/push';

  const payload = {
    to: CONFIG.USER_ID,
    messages: [
      { type: 'text', text: message }
    ]
  };

  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CONFIG.CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      Logger.log('LINE送信エラー: ' + response.getContentText());
    }
  } catch (e) {
    Logger.log('LINE送信例外: ' + e.message);
  }
}

// ========== Gmail通知 ==========

function sendGmail(subject, body) {
  if (!CONFIG.GMAIL_TO || CONFIG.GMAIL_TO.indexOf('ここに') !== -1) return;
  try {
    GmailApp.sendEmail(CONFIG.GMAIL_TO, subject, body, { name: CONFIG.ADMIN_NAME });
  } catch (e) {
    Logger.log('Gmail送信失敗: ' + e.toString());
  }
}

// ========== 現在時刻 ==========

function getCurrentTime() {
  const now = new Date();
  return Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/M/d H:mm');
}

// ========== スプレッドシートの取得・自動作成 ==========

function getOrCreateSpreadsheet() {
  if (spreadsheetCache_) return spreadsheetCache_;

  // 1. スクリプトプロパティから取得
  const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (ssId) {
    try {
      spreadsheetCache_ = SpreadsheetApp.openById(ssId);
      return spreadsheetCache_;
    } catch (e) {
      Logger.log('Property SS ID failed: ' + e.message);
    }
  }

  // 2. アクティブなSS（バインドされている場合）
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', active.getId());
      spreadsheetCache_ = active;
      return spreadsheetCache_;
    }
  } catch (e) { }

  // 3. 既知のIDを優先
  try {
    spreadsheetCache_ = SpreadsheetApp.openById(FALLBACK_SPREADSHEET_ID);
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetCache_.getId());
    return spreadsheetCache_;
  } catch (e) {
    Logger.log('Fallback SS ID failed: ' + e.message);
  }

  // 4. 名前で検索
  const files = DriveApp.getFilesByName(CONFIG.SHEET_NAME);
  if (files.hasNext()) {
    const ss = SpreadsheetApp.open(files.next());
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
    spreadsheetCache_ = ss;
    return spreadsheetCache_;
  }

  // 5. 最後は新規作成
  spreadsheetCache_ = createSpreadsheet();
  return spreadsheetCache_;
}

// ========== 初回セットアップ（最初に1回だけ実行）==========

function setupSpreadsheet() {
  const ss = createSpreadsheet();
  const url = ss.getUrl();
  Logger.log('スプレッドシート作成完了: ' + url);
  SpreadsheetApp.getUi().alert(
    '✅ セットアップ完了！\n\n'
    + 'スプレッドシートのURLをコピーして保存してください：\n'
    + url + '\n\n'
    + '次の手順：\n'
    + '1. 上記URLを開いて3つのシートを確認\n'
    + '2. 商品マスタシートに商品情報が入っているか確認\n'
    + '3. ブログ・お知らせシートに記事を追加\n'
    + '4. CONFIG のGMAIL_TOを入力して再デプロイ'
  );
}

// ========== スプレッドシート作成処理 ==========

function createSpreadsheet() {
  const ss = SpreadsheetApp.create(CONFIG.SHEET_NAME);

  // ===== シート1：注文管理 =====
  const orderSheet = ss.getActiveSheet();
  orderSheet.setName(SHEETS.ORDERS);
  orderSheet.getRange(1, 1, 1, 13).setValues([[
    '注文ID', '注文日時', '注文者名', '商品名', '個数', '単価（円）', '仕入値（円）', '小計（円）', '純利益（円）', '合計金額', '支払方法', 'ステータス', '受取確認'
  ]]);
  styleHeader(orderSheet, 13, '#1a4a73');
  // 列幅設定
  orderSheet.setColumnWidth(1, 160);
  orderSheet.setColumnWidth(2, 110);
  orderSheet.setColumnWidth(3, 100);
  orderSheet.setColumnWidth(4, 220);
  orderSheet.setColumnWidth(5, 60);
  orderSheet.setColumnWidth(6, 90);
  orderSheet.setColumnWidth(7, 90);
  orderSheet.setColumnWidth(8, 90);
  orderSheet.setColumnWidth(9, 90);
  orderSheet.setColumnWidth(10, 100);
  orderSheet.setColumnWidth(11, 80);
  orderSheet.setColumnWidth(12, 90);
  orderSheet.setColumnWidth(13, 80);
  // 数値フォーマット
  orderSheet.getRange('E2:I1000').setNumberFormat('#,##0');
  orderSheet.getRange('J2:J1000').setNumberFormat('"¥"#,##0');
  // チェックボックス
  orderSheet.getRange('M2:M1000').insertCheckboxes();

  // 条件付きフォーマット：受取確認済み、またはキャンセル済み
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$M2=TRUE')
    .setBackground('#eeeeee')
    .setStrikethrough(true)
    .setFontColor('#9e9e9e')
    .setRanges([orderSheet.getRange('A2:M1000')])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$L2="キャンセル済"')
    .setBackground('#eeeeee')
    .setStrikethrough(true)
    .setFontColor('#9e9e9e')
    .setRanges([orderSheet.getRange('A2:M1000')])
    .build());
  orderSheet.setConditionalFormatRules(rules);

  // ===== シート2：商品マスタ =====
  const productSheet = ss.insertSheet(SHEETS.PRODUCTS);
  productSheet.getRange(1, 1, 1, 9).setValues([[
    'カテゴリ', '商品名', '価格（円）', 'アイコン', '背景色コード', '公開設定', '商品説明', '商品説明画像', '更新日時'
  ]]);
  styleHeader(productSheet, 9, '#2e7d32');

  // 商品サンプルデータ
  const productUpdatedAt = getCurrentTime();
  const products = [
    ['よもぎ茶', 'よもぎ茶（50パック）', 2100, '🍵', 'c1', '公開', '', '', productUpdatedAt],
    ['よもぎ茶', 'よもぎ茶（30パック）', 1575, '🍵', 'c1', '公開', '', '', productUpdatedAt],
    ['よもぎ入浴剤', 'よもぎ入浴剤（10パック）', 1540, '🛁', 'c4', '公開', '', '', productUpdatedAt],
    ['よもぎ入浴剤', 'よもぎ入浴剤（2パック）', 350, '🛁', 'c4', '公開', '', '', productUpdatedAt],
    ['石鹸', '石鹸（あこがれのきよら）', 540, '🧼', 'c2', '公開', '', '', productUpdatedAt],
    ['冷え取りパット', '冷え取りパット（3枚セット）', 2400, '🌸🌸🌸', 'c3', '公開', '', '', productUpdatedAt],
    ['冷え取りパット', '冷え取りパット（2枚）', 1760, '🌸🌸', 'c3', '公開', '', '', productUpdatedAt],
    ['冷え取りパット', '冷え取りパット（1枚）', 880, '🌸', 'c3', '公開', '', '', productUpdatedAt],
  ];
  productSheet.getRange(2, 1, products.length, 9).setValues(products);
  productSheet.setColumnWidth(1, 120);
  productSheet.setColumnWidth(2, 240);
  productSheet.setColumnWidth(3, 100);
  productSheet.setColumnWidth(4, 80);
  productSheet.setColumnWidth(5, 100);
  productSheet.setColumnWidth(6, 90);
  productSheet.getRange('C2:C1000').setNumberFormat('#,##0');

  // 注意書き（H列）
  productSheet.getRange('H1').setValue('【価格変更・追加の方法】').setFontColor('#888888').setFontSize(10);
  productSheet.getRange('H2').setValue('C列の数字を変えると価格が変わります').setFontColor('#888888').setFontSize(10);
  productSheet.getRange('H3').setValue('新商品は新しい行に追加してください').setFontColor('#888888').setFontSize(10);
  productSheet.getRange('H4').setValue('F列を「非公開」にすると注文管理に記録されなくなります').setFontColor('#888888').setFontSize(10);
  productSheet.setColumnWidth(8, 340);

  // ===== シート3：ブログ・お知らせ =====
  const blogSheet = ss.insertSheet(SHEETS.BLOG);
  blogSheet.getRange(1, 1, 1, 7).setValues([[
    '投稿日（例：2025-06-15）', 'タイトル', 'カテゴリ', 'アイコン絵文字', '本文', '公開設定', '更新日時'
  ]]);
  styleHeader(blogSheet, 7, '#7b1fa2');

  // サンプルデータ
  const blogs = [
    ['2025-06-15', '夏の産後ヨガ体験会のご案内', 'お知らせ', '🌷', '今年の夏も産後ヨガ体験会を開催いたします。初めての方も大歓迎です。お気軽にご参加ください。', '公開', formatDateTime_(new Date())],
    ['2025-06-10', '授乳期にうれしい！おすすめハーブティー', 'ブログ', '🍵', '授乳中のママにぴったりのよもぎ茶をご紹介します。ノンカフェインで体を温める効果があります。', '公開', formatDateTime_(new Date())],
    ['2025-06-05', '6月22日（日）は臨時休診となります', '休診情報', '📅', '誠に恐れ入りますが6月22日（日）は臨時休診とさせていただきます。ご不便をおかけして申し訳ございません。', '公開', formatDateTime_(new Date())],
  ];
  blogSheet.getRange(2, 1, blogs.length, 7).setValues(blogs);
  blogSheet.setColumnWidth(1, 180);
  blogSheet.setColumnWidth(2, 260);
  blogSheet.setColumnWidth(3, 100);
  blogSheet.setColumnWidth(4, 80);
  blogSheet.setColumnWidth(5, 340);
  blogSheet.setColumnWidth(6, 90);

  // 注意書き
  blogSheet.getRange('H1').setValue('【カテゴリの例】').setFontColor('#888888').setFontSize(10);
  blogSheet.getRange('H2').setValue('お知らせ / ブログ / 休診情報 / イベント / 商品情報').setFontColor('#888888').setFontSize(10);
  blogSheet.getRange('H3').setValue('').setFontColor('#888888').setFontSize(10);
  blogSheet.getRange('H4').setValue('【公開設定】').setFontColor('#888888').setFontSize(10);
  blogSheet.getRange('H5').setValue('公開　→ アプリに表示されます').setFontColor('#888888').setFontSize(10);
  blogSheet.getRange('H6').setValue('非公開 → アプリに表示されません').setFontColor('#888888').setFontSize(10);
  blogSheet.setColumnWidth(8, 280);

  // ===== シート4：カレンダー =====
  const calSheet = ss.insertSheet('カレンダー');
  calSheet.getRange(1, 1, 1, 5).setValues([[
    '日付（例：2025-06-15）', 'イベント名', '詳細', 'カラー（色名またはコード）', '公開設定'
  ]]);
  styleHeader(calSheet, 5, '#e57373');

  // 現在の月に合わせたサンプルデータを登録
  const todayDate = new Date();
  const calY = todayDate.getFullYear();
  const calM = String(todayDate.getMonth() + 1).padStart(2, '0');
  const calEvents = [
    [`${calY}-${calM}-15`, '📝 産後ヨガ教室', '10:00〜11:30 定員5名様', '#f48fb1', '公開'],
    [`${calY}-${calM}-22`, '休診日', '臨時休診となります', '#e57373', '公開'],
    [`${calY}-${calM}-25`, '🌿 お灸イベント', 'ご自宅でできるお灸のやり方', '#81c784', '公開']
  ];
  calSheet.getRange(2, 1, calEvents.length, 5).setValues(calEvents);
  calSheet.setColumnWidth(1, 180);
  calSheet.setColumnWidth(2, 200);
  calSheet.setColumnWidth(3, 300);
  calSheet.setColumnWidth(4, 150);
  calSheet.setColumnWidth(5, 90);

  calSheet.getRange('G1').setValue('【カラーの例】').setFontColor('#888888').setFontSize(10);
  calSheet.getRange('G2').setValue('#e57373 (休診などに)').setFontColor('#888888').setFontSize(10);
  calSheet.getRange('G3').setValue('#f48fb1 (ヨガなどに)').setFontColor('#888888').setFontSize(10);
  calSheet.getRange('G4').setValue('#81c784 (お灸などに)').setFontColor('#888888').setFontSize(10);
  calSheet.getRange('G5').setValue('#ffb74d (イベントなどに)').setFontColor('#888888').setFontSize(10);
  calSheet.setColumnWidth(7, 250);

  // ===== シート5：管理マスタ =====
  const masterSheet = ss.insertSheet(SHEETS.MASTER);
  masterSheet.getRange(1, 1, 1, 3).setValues([[
    '商品名（完全一致）', '仕入値（円）', '備考'
  ]]);
  styleHeader(masterSheet, 3, '#607d8b');
  masterSheet.setColumnWidth(1, 260);
  masterSheet.setColumnWidth(2, 120);
  masterSheet.setColumnWidth(3, 300);
  masterSheet.getRange('B2:B1000').setNumberFormat('#,##0');

  const masterData = [
    ['よもぎ茶（50パック）', 1200, ''],
    ['よもぎ茶（30パック）', 800, ''],
    ['よもぎ入浴剤（10パック）', 800, ''],
    ['よもぎ入浴剤（2パック）', 150, ''],
    ['石鹸（あこがれのきよら）', 300, ''],
    ['冷え取りパット（3枚セット）', 1200, ''],
    ['冷え取りパット（2枚）', 800, ''],
    ['冷え取りパット（1枚）', 400, ''],
  ];
  masterSheet.getRange(2, 1, masterData.length, 3).setValues(masterData);

  masterSheet.getRange('E1').setValue('【仕入値の自動入力について】').setFontColor('#888888').setFontSize(10);
  masterSheet.getRange('E2').setValue('注文管理シートでは、D列の「商品名」とここの「商品名」を照合して').setFontColor('#888888').setFontSize(10);
  masterSheet.getRange('E3').setValue('「仕入値」を自動で書き出す数式（VLOOKUP）が入ります。').setFontColor('#888888').setFontSize(10);
  masterSheet.getRange('E4').setValue('※新商品を追加する際は、必ずこちらにも商品名と仕入値を記入してください。').setFontColor('#888888').setFontSize(10);

  // ===== シート6：会員データ =====
  const usersSheet = ss.insertSheet(SHEETS.USERS);
  ensureUsersSheetStructure_(usersSheet);

  ss.setActiveSheet(orderSheet);
  ss.moveActiveSheet(1);

  return ss;
}

// ========== ヘッダー行スタイル ==========

function styleHeader(sheet, numCols, color) {
  const range = sheet.getRange(1, 1, 1, numCols);
  range.setBackground(color)
    .setFontColor('white')
    .setFontWeight('bold')
    .setFontSize(10);
  sheet.setFrozenRows(1);
}

function configureUsersSheet_(sheet) {
  if (!sheet) return;
  sheet.setColumnWidth(USER_COL.MEMBER_ID, 100);
  sheet.setColumnWidth(USER_COL.TIMESTAMP, 120);
  sheet.setColumnWidth(USER_COL.NAME, 120);
  sheet.setColumnWidth(USER_COL.KANA, 120);
  sheet.setColumnWidth(USER_COL.PHONE, 120);
  sheet.setColumnWidth(USER_COL.AVATAR_URL, 120);
  sheet.setColumnWidth(USER_COL.MEMO, 300);
  sheet.setColumnWidth(USER_COL.STATUS, 100);
  sheet.setColumnWidth(USER_COL.BIRTHDAY, 110);
  sheet.setColumnWidth(USER_COL.ADDRESS, 220);
  sheet.setColumnWidth(USER_COL.STAMP_COUNT, 100);
  sheet.setColumnWidth(USER_COL.STAMP_CARD_NUM, 120);
  sheet.setColumnWidth(USER_COL.REWARDS, 260);
  sheet.setColumnWidth(USER_COL.LAST_STAMP_DATE, 120);
  sheet.setColumnWidth(USER_COL.STAMP_ACHIEVED_AT, 150);
  sheet.setColumnWidth(USER_COL.PASSCODE, 100);
  sheet.setColumnWidth(USER_COL.TRANSFER_CODE, 120);
  sheet.setColumnWidth(USER_COL.TRANSFER_CODE_ISSUED_AT, 150);
  sheet.setColumnWidth(USER_COL.DEVICE_SESSIONS, 260);
  sheet.setColumnWidth(USER_COL.DELETE_STATUS, 110);
  sheet.setColumnWidth(USER_COL.DELETED_AT, 160);
  sheet.setColumnWidth(USER_COL.MERGED_INTO, 120);
  sheet.setColumnWidth(USER_COL.STAMP_HISTORY_JSON, 260);
  sheet.setColumnWidth(USER_COL.LAST_STAMP_AT, 160);
}

function ensureUsersSheetStructure_(sheet) {
  if (!sheet) return null;
  const maxCols = sheet.getMaxColumns();
  let needsFormatting = false;

  if (maxCols < USER_HEADERS.length) {
    sheet.insertColumnsAfter(maxCols, USER_HEADERS.length - maxCols);
    needsFormatting = true;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, USER_HEADERS.length).getDisplayValues()[0];
  const headersNeedSync = USER_HEADERS.some(function (header, index) {
    return String(currentHeaders[index] || '') !== header;
  });

  if (headersNeedSync) {
    sheet.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]);
    needsFormatting = true;
  }

  if (needsFormatting || sheet.getFrozenRows() !== 1) {
    styleHeader(sheet, USER_HEADERS.length, '#0288d1');
    configureUsersSheet_(sheet);
  }
  return sheet;
}

function getOrCreateUsersSheet_(ss) {
  let sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.USERS);
  }
  return ensureUsersSheetStructure_(sheet);
}

function ensureOrdersSheetStructure_(sheet) {
  if (!sheet) return null;
  const headers = ['注文ID', '注文日時', '注文者名', '商品名', '個数', '単価', '仕入値', '小計', '純利益', '合計金額', '支払方法', 'ステータス', '受取確認', '会員ID', '管理メモ'];
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  const current = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  let needsHeader = false;
  headers.forEach(function (label, index) {
    if (String(current[index] || '').trim() !== label) needsHeader = true;
  });
  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.setColumnWidth(15, 260);
  return sheet;
}

function padDatePart_(value) {
  return String(value).padStart(2, '0');
}

function formatDateOnly_(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function formatDateTime_(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function ensureUpdatedAtColumn_(sheet, headerLabel) {
  if (!sheet) return;
  const label = headerLabel || '更新日時';
  const maxCols = sheet.getMaxColumns();
  if (maxCols < 7) {
    sheet.insertColumnAfter(maxCols);
  }
  if (String(sheet.getRange(1, 7).getValue() || '').trim() !== label) {
    sheet.getRange(1, 7).setValue(label);
  }
  sheet.setColumnWidth(7, 180);
}

function ensureProductSheetStructure_(sheet) {
  if (!sheet) return;
  const headers = ['カテゴリ', '商品名', '価格（円）', 'アイコン', '背景色コード', '公開設定', '商品説明', '商品説明画像', '更新日時', '在庫数', '在庫警告閾値'];
  const currentMaxCols = sheet.getMaxColumns();
  if (currentMaxCols < headers.length) {
    sheet.insertColumnsAfter(currentMaxCols, headers.length - currentMaxCols);
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setColumnWidth(7, 260);
  sheet.setColumnWidth(8, 220);
  sheet.setColumnWidth(9, 180);
  sheet.setColumnWidth(10, 90);
  sheet.setColumnWidth(11, 120);
  ensureNoticeVisibilityColumn_(sheet, 6, '公開');
}

function ensureSortOrderColumn_(sheet) {
  if (!sheet) return;
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let colIdx = header.indexOf('表示順') + 1;
  if (colIdx === 0) {
    colIdx = sheet.getLastColumn() + 1;
    sheet.getRange(1, colIdx).setValue('表示順');
    sheet.setColumnWidth(colIdx, 100);
  }
  return colIdx;
}

function ensureNamedColumn_(sheet, label, width) {
  if (!sheet) return 0;
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let colIdx = header.indexOf(label) + 1;
  if (colIdx === 0) {
    colIdx = lastCol + 1;
    if (sheet.getMaxColumns() < colIdx) {
      sheet.insertColumnAfter(sheet.getMaxColumns());
    }
    sheet.getRange(1, colIdx).setValue(label);
  }
  if (width) sheet.setColumnWidth(colIdx, width);
  return colIdx;
}

function normalizePublishVisibilityStatus_(status) {
  return String(status || '').trim() === '非公開' ? '非公開' : '公開';
}

function ensureNoticeVisibilityColumn_(sheet, sourceStatusCol, defaultStatus) {
  if (!sheet) return 0;
  const noticeCol = ensureNamedColumn_(sheet, NOTICE_VISIBILITY_HEADER, 120);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return noticeCol;

  const currentValues = sheet.getRange(2, noticeCol, lastRow - 1, 1).getDisplayValues();
  const sourceValues = sourceStatusCol
    ? sheet.getRange(2, sourceStatusCol, lastRow - 1, 1).getDisplayValues()
    : null;
  const fallbackStatus = normalizePublishVisibilityStatus_(defaultStatus || '公開');
  const nextValues = [];
  let hasBlank = false;

  for (let i = 0; i < currentValues.length; i++) {
    const current = String(currentValues[i][0] || '').trim();
    const sourceStatus = sourceValues
      ? normalizePublishVisibilityStatus_(sourceValues[i][0])
      : fallbackStatus;
    if (!current) hasBlank = true;
    nextValues.push([current || sourceStatus]);
  }

  if (hasBlank) {
    sheet.getRange(2, noticeCol, nextValues.length, 1).setValues(nextValues);
  }

  return noticeCol;
}

function formatMaybeDateTime_(value) {
  if (value instanceof Date) return formatDateTime_(value);
  return String(value || '').trim();
}

function getProductCostMap_() {
  const ss = getOrCreateSpreadsheet();
  const masterSheet = ss.getSheetByName(SHEETS.MASTER);
  const costMap = {};
  if (!masterSheet) return costMap;

  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) return costMap;

  const values = masterSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  values.forEach(function (row) {
    const name = String(row[0] || '').trim();
    if (!name) return;
    costMap[name] = Number(row[1] || 0);
  });
  return costMap;
}

function normalizeDateOnlyValue_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return formatDateOnly_(value);
  }
  const text = String(value).trim();
  if (!text) return '';
  const simpleMatch = text.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (simpleMatch) {
    return [
      simpleMatch[1],
      padDatePart_(simpleMatch[2]),
      padDatePart_(simpleMatch[3])
    ].join('-');
  }
  const parsed = new Date(text);
  if (isNaN(parsed.getTime())) return '';
  return formatDateOnly_(parsed);
}

function normalizeDateTimeValue_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return formatDateTime_(value);
  }
  const text = String(value).trim();
  if (!text) return '';
  const parsed = new Date(text);
  if (isNaN(parsed.getTime())) return '';
  return formatDateTime_(parsed);
}

function sanitizeRewardEntry_(entry, fallbackId) {
  const base = entry || {};
  const earnedDate = normalizeDateTimeValue_(base.earnedDate) || formatDateTime_(new Date());
  const expiryBase = normalizeDateTimeValue_(base.expiryDate);
  let expiryDate = expiryBase;
  if (!expiryDate) {
    const expiry = new Date(earnedDate);
    expiry.setMonth(expiry.getMonth() + 1);
    expiryDate = formatDateTime_(expiry);
  }
  return {
    id: String(base.id || fallbackId || new Date().getTime()),
    cardNum: Math.max(1, Number(base.cardNum || 1) || 1),
    rewardName: String(base.rewardName || 'スタンプ達成特典'),
    earnedDate: earnedDate,
    expiryDate: expiryDate,
    used: base.used === true || String(base.used).toLowerCase() === 'true',
    usedAt: normalizeDateTimeValue_(base.usedAt)
  };
}

function sanitizeStampHistoryEntry_(entry, fallbackId) {
  const base = entry || {};
  const acquiredDate = normalizeDateTimeValue_(base.acquiredDate || base.earnedDate || base.date) || formatDateTime_(new Date());
  return {
    id: String(base.id || fallbackId || new Date().getTime()),
    cardNum: Math.max(1, Number(base.cardNum || 1) || 1),
    stampNumber: Math.max(1, Math.min(10, Number(base.stampNumber !== undefined ? base.stampNumber : base.stampCount) || 1)),
    acquiredDate: acquiredDate,
    dateKey: normalizeDateOnlyValue_(base.dateKey || acquiredDate)
  };
}

function sanitizeStampHistoryList_(entries) {
  const list = Array.isArray(entries) ? entries : [];
  const map = {};
  list.forEach(function (entry, index) {
    const normalized = sanitizeStampHistoryEntry_(entry, 'stamp-' + (index + 1) + '-' + new Date().getTime());
    const key = [normalized.cardNum, normalized.stampNumber].join('|');
    const current = map[key];
    if (!current || parseLooseDateToTimestamp_(normalized.acquiredDate) >= parseLooseDateToTimestamp_(current.acquiredDate)) {
      map[key] = normalized;
    }
  });
  return Object.keys(map).map(function (key) {
    return map[key];
  }).sort(function (a, b) {
    return parseLooseDateToTimestamp_(b.acquiredDate) - parseLooseDateToTimestamp_(a.acquiredDate);
  });
}

function deriveLastStampAt_(data, stampHistory) {
  return normalizeDateTimeValue_(data && data.lastStampAt)
    || (Array.isArray(stampHistory) && stampHistory[0] && stampHistory[0].acquiredDate)
    || normalizeDateTimeValue_(data && data.lastStampDate)
    || '';
}

function sanitizeRewardStatus_(data) {
  const rewardsInput = Array.isArray(data && data.rewards) ? data.rewards : [];
  const rewards = rewardsInput.map(function (reward, index) {
    return sanitizeRewardEntry_(reward, 'reward-' + (index + 1) + '-' + new Date().getTime());
  });
  const stampHistory = sanitizeStampHistoryList_(data && data.stampHistory);
  const lastStampAt = deriveLastStampAt_(data, stampHistory);
  return {
    stampCount: Math.max(0, Math.min(10, Number(data && data.stampCount) || 0)),
    stampCardNum: Math.max(1, Number(data && data.stampCardNum) || 1),
    rewards: rewards,
    stampHistory: stampHistory,
    lastStampDate: normalizeDateOnlyValue_(data && data.lastStampDate || lastStampAt),
    lastStampAt: lastStampAt,
    stampAchievedDate: normalizeDateTimeValue_(data && data.stampAchievedDate)
  };
}

function parseRewardsJson_(value) {
  if (!value) return [];
  try {
    const parsed = JSON.parse(String(value));
    if (!Array.isArray(parsed)) return [];
    return parsed.map(function (reward, index) {
      return sanitizeRewardEntry_(reward, 'reward-' + (index + 1));
    });
  } catch (err) {
    Logger.log('parseRewardsJson_ error: ' + err.toString());
    return [];
  }
}

function serializeRewardsJson_(rewards) {
  return JSON.stringify((rewards || []).map(function (reward, index) {
    return sanitizeRewardEntry_(reward, 'reward-' + (index + 1));
  }));
}

function parseStampHistoryJson_(value) {
  if (!value) return [];
  try {
    const parsed = JSON.parse(String(value));
    if (!Array.isArray(parsed)) return [];
    return sanitizeStampHistoryList_(parsed);
  } catch (err) {
    Logger.log('parseStampHistoryJson_ error: ' + err.toString());
    return [];
  }
}

function serializeStampHistoryJson_(history) {
  return JSON.stringify(sanitizeStampHistoryList_(history));
}

function getDefaultRewardStatus_() {
  return {
    stampCount: 0,
    stampCardNum: 1,
    rewards: [],
    stampHistory: [],
    lastStampDate: '',
    lastStampAt: '',
    stampAchievedDate: ''
  };
}

function getRewardStatusFromRow_(row) {
  if (!row || !row.length) return getDefaultRewardStatus_();
  return sanitizeRewardStatus_({
    stampCount: row[USER_COL.STAMP_COUNT - 1],
    stampCardNum: row[USER_COL.STAMP_CARD_NUM - 1],
    rewards: parseRewardsJson_(row[USER_COL.REWARDS - 1]),
    stampHistory: parseStampHistoryJson_(row[USER_COL.STAMP_HISTORY_JSON - 1]),
    lastStampDate: row[USER_COL.LAST_STAMP_DATE - 1],
    lastStampAt: row[USER_COL.LAST_STAMP_AT - 1],
    stampAchievedDate: row[USER_COL.STAMP_ACHIEVED_AT - 1]
  });
}

function applyRewardStatusToRow_(row, rewardStatus) {
  const next = row.slice();
  next[USER_COL.STAMP_COUNT - 1] = rewardStatus.stampCount;
  next[USER_COL.STAMP_CARD_NUM - 1] = rewardStatus.stampCardNum;
  next[USER_COL.REWARDS - 1] = serializeRewardsJson_(rewardStatus.rewards);
  next[USER_COL.LAST_STAMP_DATE - 1] = rewardStatus.lastStampDate || '';
  next[USER_COL.STAMP_ACHIEVED_AT - 1] = rewardStatus.stampAchievedDate || '';
  next[USER_COL.STAMP_HISTORY_JSON - 1] = serializeStampHistoryJson_(rewardStatus.stampHistory);
  next[USER_COL.LAST_STAMP_AT - 1] = rewardStatus.lastStampAt || '';
  return next;
}

function normalizeDeviceSessionList_(input) {
  let raw = input;
  if (!raw) return [];
  if (typeof raw === 'string') {
    try {
      raw = JSON.parse(raw);
    } catch (err) {
      return [];
    }
  }
  if (!Array.isArray(raw)) return [];
  return raw.map(function (session) {
    return {
      deviceId: String(session && session.deviceId || '').trim(),
      label: String(session && session.label || '').trim(),
      platform: String(session && session.platform || '').trim(),
      appVersion: String(session && session.appVersion || '').trim(),
      lastSeenAt: normalizeDateTimeValue_(session && session.lastSeenAt) || formatDateTime_(new Date()),
      passcodeEnabled: session && session.passcodeEnabled === true,
      pushEnabled: session && session.pushEnabled === true,
      current: session && session.current === true
    };
  }).filter(function (session) {
    return !!session.deviceId;
  }).sort(function (a, b) {
    return parseLooseDateToTimestamp_(b.lastSeenAt) - parseLooseDateToTimestamp_(a.lastSeenAt);
  }).slice(0, MAX_DEVICE_SESSIONS);
}

function serializeDeviceSessionList_(sessions) {
  return JSON.stringify(normalizeDeviceSessionList_(sessions));
}

function mergeDeviceSessionLists_(left, right) {
  const mergedMap = {};
  normalizeDeviceSessionList_(left).concat(normalizeDeviceSessionList_(right)).forEach(function (session) {
    if (!session.deviceId) return;
    const current = mergedMap[session.deviceId];
    if (!current || parseLooseDateToTimestamp_(session.lastSeenAt) > parseLooseDateToTimestamp_(current.lastSeenAt)) {
      mergedMap[session.deviceId] = Object.assign({}, current || {}, session);
    } else {
      mergedMap[session.deviceId] = Object.assign({}, session, current);
    }
  });
  return Object.keys(mergedMap).map(function (key) {
    return mergedMap[key];
  }).sort(function (a, b) {
    return parseLooseDateToTimestamp_(b.lastSeenAt) - parseLooseDateToTimestamp_(a.lastSeenAt);
  }).slice(0, MAX_DEVICE_SESSIONS);
}

function getUserDeviceSessionsFromRow_(row) {
  if (!row || !row.length) return [];
  return normalizeDeviceSessionList_(row[USER_COL.DEVICE_SESSIONS - 1]);
}

function applyUserDeviceSessionsToRow_(row, sessions) {
  const next = row.slice();
  next[USER_COL.DEVICE_SESSIONS - 1] = serializeDeviceSessionList_(sessions);
  return next;
}

function upsertUserDeviceSession_(sheet, rowIdx, deviceData) {
  if (!sheet || rowIdx <= 1) return [];
  const range = sheet.getRange(rowIdx, 1, 1, USER_HEADERS.length);
  const row = range.getValues()[0];
  const sessions = getUserDeviceSessionsFromRow_(row);
  const next = normalizeDeviceSessionList_(sessions).filter(function (session) {
    return session.deviceId !== String(deviceData && deviceData.deviceId || '').trim();
  });
  next.unshift({
    deviceId: String(deviceData && deviceData.deviceId || '').trim(),
    label: String(deviceData && deviceData.label || '').trim(),
    platform: String(deviceData && deviceData.platform || '').trim(),
    appVersion: String(deviceData && deviceData.appVersion || '').trim(),
    lastSeenAt: formatDateTime_(new Date()),
    passcodeEnabled: deviceData && deviceData.passcodeEnabled === true,
    pushEnabled: deviceData && deviceData.pushEnabled === true,
    current: true
  });
  const normalized = normalizeDeviceSessionList_(next).map(function (session) {
    session.current = session.deviceId === String(deviceData && deviceData.deviceId || '').trim();
    return session;
  });
  range.setValues([applyUserDeviceSessionsToRow_(row, normalized)]);
  return normalized;
}

function getRewardGachaPrizeMeta_(rewardName) {
  const normalized = String(rewardName || '').trim();
  const matched = REWARD_GACHA_PRIZE_POOL.find(function (prize) {
    return prize.rewardName === normalized ||
      prize.rankLabel === normalized ||
      (normalized && normalized.indexOf(prize.rankLabel) === 0);
  });
  if (matched) {
    return matched;
  }
  return {
    key: 'SPECIAL',
    rankLabel: 'ごほうび獲得',
    rewardName: normalized || '特典プレゼント',
    capsuleColor: '#d9c5a2',
    accentColor: '#8d6c46',
    message: '受付でその時の特典をお受け取りください。',
    weight: 0
  };
}

function pickWeightedRewardGachaPrize_() {
  const totalWeight = REWARD_GACHA_PRIZE_POOL.reduce(function (sum, prize) {
    return sum + Math.max(1, Number(prize.weight) || 0);
  }, 0);
  let roll = Math.random() * totalWeight;
  for (let i = 0; i < REWARD_GACHA_PRIZE_POOL.length; i++) {
    roll -= Math.max(1, Number(REWARD_GACHA_PRIZE_POOL[i].weight) || 0);
    if (roll < 0) {
      return REWARD_GACHA_PRIZE_POOL[i];
    }
  }
  return REWARD_GACHA_PRIZE_POOL[REWARD_GACHA_PRIZE_POOL.length - 1];
}

function buildRewardGachaEntry_(rewardStatus, prizeMeta) {
  const earnedDate = normalizeDateTimeValue_(rewardStatus && rewardStatus.stampAchievedDate) || formatDateTime_(new Date());
  const expiryDate = new Date(earnedDate);
  expiryDate.setMonth(expiryDate.getMonth() + 1);
  return sanitizeRewardEntry_({
    id: 'reward-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000),
    cardNum: Math.max(1, Number(rewardStatus && rewardStatus.stampCardNum) || 1),
    rewardName: prizeMeta.rewardName,
    earnedDate: earnedDate,
    expiryDate: formatDateTime_(expiryDate),
    used: false
  });
}

function buildRewardGachaResult_(reward, alreadyDrawn) {
  const prizeMeta = getRewardGachaPrizeMeta_(reward && reward.rewardName);
  return {
    key: prizeMeta.key,
    rankLabel: prizeMeta.rankLabel,
    rewardName: String((reward && reward.rewardName) || prizeMeta.rewardName || '特典プレゼント'),
    capsuleColor: prizeMeta.capsuleColor,
    accentColor: prizeMeta.accentColor,
    message: prizeMeta.message,
    alreadyDrawn: !!alreadyDrawn
  };
}

function findUserRowByMemberId_(sheet, memberId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, USER_COL.MEMBER_ID, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === String(memberId || '').trim()) {
      return i + 2;
    }
  }
  return -1;
}

function clearUserPushSubscription_(memberId) {
  const normalizedMemberId = String(memberId || '').trim();
  if (!normalizedMemberId) return false;

  const ss = getOrCreateSpreadsheet();
  const sheet = getOrCreateUsersSheet_(ss);
  const rowIdx = findUserRowByMemberId_(sheet, normalizedMemberId);
  if (rowIdx === -1) return false;

  const range = sheet.getRange(rowIdx, 1, 1, USER_HEADERS.length);
  const row = range.getValues()[0];
  row[USER_COL.PUSH - 1] = '';
  row[USER_COL.TIMESTAMP - 1] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/M/d H:mm');
  range.setValues([row]);
  return true;
}

// ============================================================
//  管理者用機能
// ============================================================

// ========== 管理者用：注文一覧取得 ==========

function getAdminOrders(params) {
  const showAll = params && params.showAll === true;
  const ss = getOrCreateSpreadsheet();
  const sheet = ensureOrdersSheetStructure_(ss.getSheetByName(SHEETS.ORDERS));
  if (!sheet) return { status: 'ok', orders: [] };
  const deleteCols = ensureSoftDeleteColumns_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'ok', orders: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const orderMap = {};

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol)) continue;
    const orderId = String(row[0]);
    if (!orderId) continue;

    if (!orderMap[orderId]) {
      orderMap[orderId] = {
        orderId: orderId,
        date: (row[1] instanceof Date) ? Utilities.formatDate(row[1], 'Asia/Tokyo', 'yyyy/M/d H:mm') : String(row[1] || ''),
        customerName: String(row[2]),
        items: [],
        total: String(row[9]),
        payment: String(row[10]),
        status: normalizeOrderStatus_(row[11] || '受付中'),
        checked: row[12] === true,
        memberId: String(row[13] || ''),
        internalNote: String(row[14] || ''),
        rowIndices: []
      };
    }

    orderMap[orderId].items.push({
      name: String(row[3]),
      qty: Number(row[4]),
      price: Number(row[5])
    });
    orderMap[orderId].rowIndices.push(i + 2);
  }

  // 配列に変換（rowIdxは代表として最初の行を指定）※キャンセル済みは除外
  const orders = Object.keys(orderMap).map(id => {
    const o = orderMap[id];
    o.rowIdx = o.rowIndices[0]; // 後方互換性のため
    o.itemName = o.items.map(it => it.name + ' (' + it.qty + ')').join('\n');
    return o;
  }).filter(o => {
    if (showAll) return o.status !== 'キャンセル済';
    return o.status !== 'キャンセル済' && o.status !== '受取済' && !o.checked;
  });

  // 新しい順
  orders.reverse();
  return { status: 'ok', orders: orders };
}

function getAdminUserOrders(params) {
  const memberId = String(params && params.memberId || '').trim();
  if (!memberId) return { status: 'error', message: '会員IDが必要です' };

  const ss = getOrCreateSpreadsheet();
  const sheet = ensureOrdersSheetStructure_(ss.getSheetByName(SHEETS.ORDERS));
  if (!sheet) return { status: 'ok', orders: [] };
  const deleteCols = ensureSoftDeleteColumns_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'ok', orders: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const orderMap = {};

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol)) continue;
    if (String(row[13] || '').trim() !== memberId) continue;

    const orderId = String(row[0] || '').trim();
    if (!orderId) continue;

    if (!orderMap[orderId]) {
      orderMap[orderId] = {
        orderId: orderId,
        date: (row[1] instanceof Date) ? Utilities.formatDate(row[1], 'Asia/Tokyo', 'yyyy/M/d H:mm') : String(row[1] || ''),
        customerName: String(row[2] || ''),
        items: [],
        total: String(row[9] || ''),
        payment: String(row[10] || ''),
        status: normalizeOrderStatus_(row[11] || '受付中'),
        checked: row[12] === true,
        memberId: memberId,
        internalNote: String(row[14] || ''),
        rowIndices: []
      };
    }

    orderMap[orderId].items.push({
      name: String(row[3] || ''),
      qty: Number(row[4]) || 0,
      price: Number(row[5]) || 0
    });
    orderMap[orderId].rowIndices.push(i + 2);
  }

  const orders = Object.keys(orderMap).map(function (id) {
    const order = orderMap[id];
    order.rowIdx = order.rowIndices[0] || 0;
    order.itemName = order.items.map(function (item) {
      return item.name + ' (' + item.qty + ')';
    }).join('\n');
    return order;
  }).sort(function (a, b) {
    return parseLooseDateToTimestamp_(b.date) - parseLooseDateToTimestamp_(a.date);
  });

  return { status: 'ok', orders: orders };
}

// ========== 管理者用：注文ステータス更新 ==========

function handleUpdateOrder(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.ORDERS);
    const lastRow = sheet.getLastRow();

    // 注文IDが指定されている場合は、そのIDを持つすべての行を更新
    if (data.orderId) {
      const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < values.length; i++) {
        if (String(values[i][0]) === String(data.orderId)) {
          let row = i + 2;
          if (data.status !== undefined) {
            sheet.getRange(row, 12).setValue(normalizeOrderStatus_(data.status));
          }
          if (data.checked !== undefined) {
            sheet.getRange(row, 13).setValue(data.checked === true);
          }
        }
      }
    } else if (data.rowIdx) {
      // 単一行の更新（従来互換）
      const rowIdx = Number(data.rowIdx);
      if (data.status !== undefined) {
        sheet.getRange(rowIdx, 12).setValue(normalizeOrderStatus_(data.status));
      }
      if (data.checked !== undefined) {
        sheet.getRange(rowIdx, 13).setValue(data.checked === true);
      }
    }

    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleUpdateOrder error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

// ========== 会員用：自分の注文履歴取得 ==========

function getCustomerOrders(params) {
  if (!params || !params.memberId) {
    return { status: 'error', message: '会員IDが必要です' };
  }
  const memberId = String(params.memberId).trim();
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ORDERS);
  if (!sheet) return { status: 'ok', orders: [] };
  const deleteCols = ensureSoftDeleteColumns_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'ok', orders: [] };

  // N列(14列)まで取得
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const orderMap = {};

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol)) continue;
    const rowMemberId = String(row[13]).trim(); // N列: 会員ID
    if (rowMemberId !== memberId) continue;

    // 受取済み・キャンセル済みの注文は履歴から除外する
    const rowStatus = normalizeOrderStatus_(row[11] || '受付中');
    const isReceived = rowStatus === '受取済' || rowStatus.indexOf('受取済') !== -1;
    const isCancelled = rowStatus === 'キャンセル済' || rowStatus.indexOf('キャンセル') !== -1;
    if (isReceived || isCancelled) continue;

    const orderId = String(row[0]).trim();
    if (!orderId) continue;

    if (!orderMap[orderId]) {
      orderMap[orderId] = {
        id: orderId,
        time: row[1],
        items: [],
        total: row[9],
        payment: String(row[10]),
        status: normalizeOrderStatus_(row[11] || '受付中'),
        checked: row[12] === true
      };
    }

    orderMap[orderId].items.push({
      name: String(row[3]),
      qty: Number(row[4]),
      price: Number(row[5])
    });
  }

  const orders = Object.keys(orderMap).map(id => {
    const o = orderMap[id];
    if (o.status === '受付中') o.status = 'pending';
    else if (o.status === 'キャンセル済') o.status = 'cancelled';
    else o.status = 'done';

    if (typeof o.total === 'string') {
      o.total = Number(o.total.replace(/[^0-9.-]+/g, ""));
    }
    return o;
  });

  orders.sort((a, b) => new Date(b.time) - new Date(a.time));
  return { status: 'ok', orders: orders, debugCount: orders.length };
}

// ========== 管理者用：ブログ一覧取得 ==========

function getAdminBlogs() {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.BLOG);
  if (!sheet) return { status: 'ok', blogs: [] };
  ensureUpdatedAtColumn_(sheet, '更新日時');
  const noticeCol = ensureNoticeVisibilityColumn_(sheet, 6, '公開');
  const deleteCols = ensureSoftDeleteColumns_(sheet);
  const imageCol = ensureNamedColumn_(sheet, '画像URL', 220);
  const publishAtCol = ensurePublishAtColumn_(sheet);

  const data = sheet.getDataRange().getValues();
  const blogs = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;
    if (isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol)) continue;

    const dateVal = row[0];
    let dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy-MM-dd');
    } else {
      dateStr = String(dateVal).trim();
    }

    const body = String(row[4] || '');
    const imageUrl = getBlogImageUrlFromRow_(row, imageCol, body);
    blogs.push({
      rowIdx: i + 1,
      date: dateStr,
      title: String(row[1]),
      category: String(row[2] || 'お知らせ'),
      icon: String(row[3] || '📢'),
      body: body,
      status: String(row[5] || '公開'),
      imageUrl: imageUrl,
      publishAt: formatMaybeDateTime_(row[publishAtCol - 1]),
      updatedAt: formatMaybeDateTime_(row[6]),
      noticeStatus: normalizePublishVisibilityStatus_(row[noticeCol - 1] || row[5] || '公開')
    });
  }

  blogs.reverse();
  return { status: 'ok', blogs: blogs };
}

// ========== 管理者用：ブログ追加 ==========

function handleAddBlog(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.BLOG);
    ensureUpdatedAtColumn_(sheet, '更新日時');
    const noticeCol = ensureNoticeVisibilityColumn_(sheet, 6, '公開');
    const imageCol = ensureNamedColumn_(sheet, '画像URL', 220);
    const publishAtCol = ensurePublishAtColumn_(sheet);

    // 本文に画像URLを追記（画像がある場合）
    const body = data.body || '';
    const updatedAt = formatDateTime_(new Date());

    sheet.appendRow([
      data.date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/M/d'),
      data.title,
      data.category || 'お知らせ',
      data.icon || '📢',
      body,
      data.status || '公開',
      updatedAt
    ]);
    const rowIdx = sheet.getLastRow();
    sheet.getRange(rowIdx, noticeCol).setValue(
      normalizePublishVisibilityStatus_(data.noticeStatus || data.status || '公開')
    );
    if (imageCol) sheet.getRange(rowIdx, imageCol).setValue(String(data.image || ''));
    if (publishAtCol) sheet.getRange(rowIdx, publishAtCol).setValue(normalizePublishAtValue_(data.publishAt));

    // 自動プッシュ通知
    if (String(data.status || '公開') !== '非公開' && isPublishAtAvailable_(data.publishAt)) {
      sendAutoPush('📝 ' + (data.title || '新しい投稿'), 'NEWSが更新されました', {
        targetPage: 'news'
      });
    }

    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleAddBlog error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

/**
 * 管理者用：ブログ更新
 */
function handleUpdateBlog(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.BLOG);
    ensureUpdatedAtColumn_(sheet, '更新日時');
    const noticeCol = ensureNoticeVisibilityColumn_(sheet, 6, '公開');
    const imageCol = ensureNamedColumn_(sheet, '画像URL', 220);
    const publishAtCol = ensurePublishAtColumn_(sheet);
    const rowIdx = Number(data.rowIdx);

    // B:タイトル, C:カテゴリ, D:アイコン, E:本文, F:公開設定 (Aは日付)
    if (data.date) sheet.getRange(rowIdx, 1).setValue(data.date);
    if (data.title) sheet.getRange(rowIdx, 2).setValue(data.title);
    if (data.category) sheet.getRange(rowIdx, 3).setValue(data.category);
    if (data.icon) sheet.getRange(rowIdx, 4).setValue(data.icon);

    if (data.body !== undefined) sheet.getRange(rowIdx, 5).setValue(data.body || '');

    if (data.status) sheet.getRange(rowIdx, 6).setValue(data.status);
    sheet.getRange(rowIdx, 7).setValue(formatDateTime_(new Date()));
    if (data.noticeStatus) {
      sheet.getRange(rowIdx, noticeCol).setValue(normalizePublishVisibilityStatus_(data.noticeStatus));
    }
    if (imageCol && data.image !== undefined) {
      sheet.getRange(rowIdx, imageCol).setValue(String(data.image || ''));
    }
    if (publishAtCol && data.publishAt !== undefined) {
      sheet.getRange(rowIdx, publishAtCol).setValue(normalizePublishAtValue_(data.publishAt));
    }

    // 自動プッシュ通知
    const effectiveStatus = data.status || sheet.getRange(rowIdx, 6).getValue();
    const effectivePublishAt = publishAtCol ? sheet.getRange(rowIdx, publishAtCol).getValue() : '';
    if (String(effectiveStatus || '公開') !== '非公開' && isPublishAtAvailable_(effectivePublishAt)) {
      sendAutoPush('📝 ' + (data.title || 'ブログ更新'), 'NEWSが更新されました', {
        targetPage: 'news'
      });
    }

    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleUpdateBlog error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

// ========== 管理者用：レコードの公開/非公開切替 ==========

function handleUpdateRecordStatus(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    let sheetName = '';
    let statusCol = 6; // デフォルトF列

    if (data.sheet === 'BLOG') {
      sheetName = SHEETS.BLOG;
      statusCol = 6; // F列：公開設定
    } else if (data.sheet === 'PRODUCTS') {
      sheetName = SHEETS.PRODUCTS;
      statusCol = 6; // F列：公開設定
    } else if (data.sheet === 'CALENDAR') {
      sheetName = SHEETS.CALENDAR;
      statusCol = 5; // E列：公開設定
    } else if (data.sheet === 'SURVEYS') {
      sheetName = SHEETS.SURVEY_MASTER;
      statusCol = 5; // E列：ステータス
    }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { status: 'error', message: 'シートが見つかりません' };

    const rowIdx = Number(data.rowIdx);
    sheet.getRange(rowIdx, statusCol).setValue(data.status);

    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleUpdateRecordStatus error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function handleUpdateNoticeVisibility(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    let sheetName = '';
    let sourceStatusCol = 0;

    if (data.sheet === 'BLOG') {
      sheetName = SHEETS.BLOG;
      sourceStatusCol = 6;
    } else if (data.sheet === 'PRODUCTS') {
      sheetName = SHEETS.PRODUCTS;
      sourceStatusCol = 6;
    } else if (data.sheet === 'CALENDAR') {
      sheetName = SHEETS.CALENDAR;
      sourceStatusCol = 5;
    } else if (data.sheet === 'MENUS') {
      sheetName = SHEETS.MENUS;
      sourceStatusCol = 6;
    } else {
      return { status: 'error', message: '不明なシート名です' };
    }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { status: 'error', message: 'シートが見つかりません' };

    const rowIdx = Number(data.rowIdx || 0);
    if (rowIdx <= 1) return { status: 'error', message: '更新対象が見つかりません' };

    const noticeCol = ensureNoticeVisibilityColumn_(sheet, sourceStatusCol, '公開');
    sheet.getRange(rowIdx, noticeCol).setValue(normalizePublishVisibilityStatus_(data.status));
    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleUpdateNoticeVisibility error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

// ========== 管理者用：商品一覧取得 ==========

function getAdminProducts() {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PRODUCTS);
  if (!sheet) return { status: 'ok', products: [] };
  ensureProductSheetStructure_(sheet);
  const noticeCol = ensureNoticeVisibilityColumn_(sheet, 6, '公開');
  const deleteCols = ensureSoftDeleteColumns_(sheet);
  const publishAtCol = ensurePublishAtColumn_(sheet);
  const costMap = getProductCostMap_();

  const data = sheet.getDataRange().getValues();
  const products = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;
    if (isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol)) continue;

    products.push({
      rowIdx: i + 1,
      category: String(row[0]),
      name: String(row[1]),
      price: Number(row[2]),
      costPrice: Number(costMap[String(row[1]).trim()] || 0),
      icon: String(row[3] || '🌿'),
      bg: String(row[4] || 'c1'),
      status: String(row[5] || '公開'),
      description: String(row[6] || ''),
      descriptionImage: String(row[7] || ''),
      updatedAt: formatMaybeDateTime_(row[8]),
      stockQty: Number(row[9] || 0),
      lowStockThreshold: Number(row[10] || 0),
      publishAt: formatMaybeDateTime_(row[publishAtCol - 1]),
      noticeStatus: normalizePublishVisibilityStatus_(row[noticeCol - 1] || row[5] || '公開')
    });
  }

  return { status: 'ok', products: products };
}

// ========== 管理者用：商品追加 ==========

function handleAddProduct(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const prodSheet = ss.getSheetByName(SHEETS.PRODUCTS);
    ensureProductSheetStructure_(prodSheet);
    const noticeCol = ensureNoticeVisibilityColumn_(prodSheet, 6, '公開');
    const publishAtCol = ensurePublishAtColumn_(prodSheet);
    const updatedAt = getCurrentTime();

    prodSheet.appendRow([
      data.category || 'その他',
      data.name,
      Number(data.price),
      data.icon || '🌿',
      data.bg || 'c1',
      data.status || '公開',
      data.description || '',
      data.descriptionImage || '',
      updatedAt,
      Number(data.stockQty || 0),
      Number(data.lowStockThreshold || 0)
    ]);
    if (publishAtCol) {
      prodSheet.getRange(prodSheet.getLastRow(), publishAtCol).setValue(normalizePublishAtValue_(data.publishAt));
    }
    prodSheet.getRange(prodSheet.getLastRow(), noticeCol).setValue(
      normalizePublishVisibilityStatus_(data.noticeStatus || data.status || '公開')
    );

    // 管理マスタにも仕入値を追加
    if (data.costPrice) {
      const masterSheet = ss.getSheetByName(SHEETS.MASTER);
      if (masterSheet) {
        masterSheet.appendRow([
          data.name,
          Number(data.costPrice),
          '管理アプリから追加'
        ]);
      }
    }

    if (String(data.status || '公開') !== '非公開' && isPublishAtAvailable_(data.publishAt)) {
      sendAutoPush('🛍️ ' + (data.name || '新商品'), '新しい商品が追加されました', {
        targetPage: 'shop'
      });
    }

    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleAddProduct error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

// ========== 管理者用：商品更新（価格・公開設定） ==========

function handleUpdateProduct(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.PRODUCTS);
    ensureProductSheetStructure_(sheet);
    const noticeCol = ensureNoticeVisibilityColumn_(sheet, 6, '公開');
    const publishAtCol = ensurePublishAtColumn_(sheet);
    const rowIdx = Number(data.rowIdx);

    // A:カテゴリ, B:商品名, C:価格, D:アイコン, E:背景色, F:公開設定
    if (data.category) sheet.getRange(rowIdx, 1).setValue(data.category);
    if (data.name) {
      const oldName = sheet.getRange(rowIdx, 2).getValue();
      sheet.getRange(rowIdx, 2).setValue(data.name);

      // 商品名が変わった場合、マスタの名称も更新を試みる
      if (oldName !== data.name) {
        const masterSheet = ss.getSheetByName(SHEETS.MASTER);
        if (masterSheet) {
          const mData = masterSheet.getDataRange().getValues();
          for (let i = 0; i < mData.length; i++) {
            if (String(mData[i][0]) === String(oldName)) {
              masterSheet.getRange(i + 1, 1).setValue(data.name);
            }
          }
        }
      }
    }
    if (data.price !== undefined) sheet.getRange(rowIdx, 3).setValue(Number(data.price));
    if (data.icon) sheet.getRange(rowIdx, 4).setValue(data.icon);
    if (data.bg) sheet.getRange(rowIdx, 5).setValue(data.bg);
    if (data.status) sheet.getRange(rowIdx, 6).setValue(data.status);
    if (data.description !== undefined) sheet.getRange(rowIdx, 7).setValue(data.description);
    if (data.stockQty !== undefined) sheet.getRange(rowIdx, 10).setValue(Number(data.stockQty || 0));
    if (data.lowStockThreshold !== undefined) sheet.getRange(rowIdx, 11).setValue(Number(data.lowStockThreshold || 0));
    if (data.descriptionImage !== undefined) sheet.getRange(rowIdx, 8).setValue(data.descriptionImage);
    sheet.getRange(rowIdx, 9).setValue(getCurrentTime());
    if (publishAtCol && data.publishAt !== undefined) {
      sheet.getRange(rowIdx, publishAtCol).setValue(normalizePublishAtValue_(data.publishAt));
    }
    if (data.noticeStatus) {
      sheet.getRange(rowIdx, noticeCol).setValue(normalizePublishVisibilityStatus_(data.noticeStatus));
    }

    // 仕入値の更新（マスタ）
    if (data.costPrice !== undefined) {
      const masterSheet = ss.getSheetByName(SHEETS.MASTER);
      if (masterSheet) {
        const pName = data.name || sheet.getRange(rowIdx, 2).getValue();
        const mData = masterSheet.getDataRange().getValues();
        let updated = false;
        for (let i = 0; i < mData.length; i++) {
          if (String(mData[i][0]) === String(pName)) {
            masterSheet.getRange(i + 1, 2).setValue(Number(data.costPrice));
            updated = true;
            break;
          }
        }
        if (!updated) {
          masterSheet.appendRow([pName, Number(data.costPrice), '編集時に追加']);
        }
      }
    }

    const effectiveStatus = data.status !== undefined ? data.status : String(sheet.getRange(rowIdx, 6).getValue() || '公開');
    const effectivePublishAt = publishAtCol ? sheet.getRange(rowIdx, publishAtCol).getValue() : '';
    if (String(effectiveStatus || '公開') !== '非公開' && isPublishAtAvailable_(effectivePublishAt)) {
      sendAutoPush('🛍️ ' + (data.name || '商品情報更新'), 'ショップの商品情報が更新されました', {
        targetPage: 'shop'
      });
    }

    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleUpdateProduct error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

// ========== 管理者用：画像をGoogle Driveにアップロード ==========

function handleUploadImage(data) {
  try {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.base64),
      data.mimeType,
      data.filename
    );

    // 「まゆみ助産院_画像」フォルダを取得or作成
    const folderName = 'まゆみ助産院_画像';
    let folder;
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // 直接表示可能なURLを生成
    const fileId = file.getId();
    const url = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1000';

    return { status: 'ok', url: url, fileId: fileId };
  } catch (err) {
    Logger.log('handleUploadImage error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

// ========== 管理者用：行削除 ==========

function handleDeleteRow(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    let sheetName = '';
    if (data.sheet === 'BLOG') {
      sheetName = SHEETS.BLOG;
    } else if (data.sheet === 'PRODUCTS') {
      sheetName = SHEETS.PRODUCTS;
    } else if (data.sheet === 'CALENDAR') {
      sheetName = SHEETS.CALENDAR;
    } else if (data.sheet === 'PUSH') {
      sheetName = SHEETS.PUSH;
    } else {
      return { status: 'error', message: '不明なシート名: ' + (data.sheet || '空') };
    }
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { status: 'error', message: 'シートが見つかりません' };
    const rowIdx = Number(data.rowIdx);
    if (rowIdx < 2) return { status: 'error', message: '削除できない行です' };
    markRowSoftDeleted_(sheet, rowIdx, '管理画面から削除');
    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleDeleteRow error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function handleDeleteRows(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    let sheetName = '';
    if (data.sheet === 'BLOG') {
      sheetName = SHEETS.BLOG;
    } else if (data.sheet === 'PRODUCTS') {
      sheetName = SHEETS.PRODUCTS;
    } else if (data.sheet === 'CALENDAR') {
      sheetName = SHEETS.CALENDAR;
    } else if (data.sheet === 'PUSH') {
      sheetName = SHEETS.PUSH;
    } else {
      return { status: 'error', message: '不明なシート名: ' + (data.sheet || '空') };
    }
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { status: 'error', message: 'シートが見つかりません' };

    // 行番を数値の配列にし、降順にソート（下の行から消さないとズレるため）
    const rowIndices = data.rowIndices.map(Number).sort((a, b) => b - a);

    for (const rowIdx of rowIndices) {
      if (rowIdx >= 2) {
        markRowSoftDeleted_(sheet, rowIdx, '一括削除');
      }
    }
    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleDeleteRows error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

/**
 * 注文の削除処理
 * data.orderIds: 削除対象の注文ID配列
 */
function handleDeleteOrders(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.ORDERS);
    if (!sheet) return { status: 'error', message: '注文シートが見つかりません' };

    const orderIds = data.orderIds;
    if (!orderIds || !Array.isArray(orderIds)) {
      return { status: 'error', message: '注文IDが指定されていません' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok' };

    const deleteCols = ensureSoftDeleteColumns_(sheet);
    const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const rowsToDelete = [];

    // 削除対象の行番号を特定
    for (let i = 0; i < values.length; i++) {
      const cellVal = String(values[i][0]).trim();
      if (isSoftDeletedByColumns_(values[i], deleteCols.statusCol, deleteCols.deletedAtCol)) continue;
      if (orderIds.some(id => String(id).trim() === cellVal)) {
        rowsToDelete.push(i + 2);
      }
    }

    for (const rowIdx of rowsToDelete) {
      markRowSoftDeleted_(sheet, rowIdx, '一括削除');
    }

    return { status: 'ok', count: orderIds.length };
  } catch (err) {
    Logger.log('handleDeleteOrders error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

/**
 * お客様による受取確認処理
 * data.orderId: 更新対象の注文ID
 */
function handleConfirmReceipt(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.ORDERS);
    if (!sheet) return { status: 'error', message: '注文シートが見つかりません' };

    const orderId = String(data.orderId).trim();
    if (!orderId) return { status: 'error', message: '注文IDが指定されていません' };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'error', message: 'データがありません' };

    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let updated = false;

    Logger.log('Target orderId: ' + orderId);

    // 該当するすべての行を更新
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === orderId) {
        let row = i + 2;
        sheet.getRange(row, 12).setValue('受取済'); // L列: ステータス
        sheet.getRange(row, 13).setValue(true);      // M列: 受取確認チェックボックス
        updated = true;
      }
    }

    if (!updated) return { status: 'error', message: '対象の注文が見つかりません' };

    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleConfirmReceipt error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

/**
 * データ分析用：月別×商品別の注文個数集計
 */
function normalizeMenuRevenueType_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  if (text.indexOf('母乳') !== -1) return '母乳外来';
  if (text.indexOf('ビジ') !== -1) return 'ビジリス';
  if (text.indexOf('教室') !== -1) return '教室';
  return text; // 既知のタイプ以外（その他で入力されたものなど）もそのまま許可する
}

function isDashiProductName_(value) {
  const text = String(value || '').trim();
  return DASHI_PRODUCT_NAMES.indexOf(text) !== -1;
}

function buildDashiTierBreakdown_(qty) {
  const count = Math.max(0, Number(qty) || 0);
  const breakdown = {
    '20': 0,
    '25': 0,
    '30': 0,
    '35': 0
  };

  if (count <= 0) return breakdown;

  if (count <= 2) {
    breakdown['20'] = count;
    return breakdown;
  }
  if (count === 3) {
    breakdown['20'] = 2;
    breakdown['25'] = 1;
    return breakdown;
  }
  if (count === 4) {
    breakdown['20'] = 3;
    breakdown['30'] = 1;
    return breakdown;
  }
  if (count === 5) {
    breakdown['20'] = 4;
    breakdown['35'] = 1;
    return breakdown;
  }

  breakdown['25'] = count;
  return breakdown;
}

function getDashiDetailKey_(productName, tierKey) {
  return String(productName || '').trim() + '|dashi|' + tierKey;
}

function ensureDashiDetailRows_(productDetails, productName) {
  DASHI_PRICE_TIERS.forEach(function (tier) {
    const detailKey = getDashiDetailKey_(productName, tier.key);
    if (!productDetails[detailKey]) {
      productDetails[detailKey] = {
        name: String(productName || '').trim(),
        price: tier.price,
        priceLabel: tier.label,
        sortOrder: tier.sortOrder,
        qty: 0,
        sales: 0,
        cost: 0,
        profit: 0
      };
    }
  });
}

function formatDashiPriceBreakdownText_(breakdown) {
  const items = [];
  DASHI_PRICE_TIERS.forEach(function (tier) {
    const qty = Number(breakdown && breakdown[tier.key] || 0);
    if (!qty) return;
    items.push(tier.label + ' × ' + qty + '個');
  });
  return items.join(' / ');
}

function getOrCreateMenuRevenueSheet_(ss) {
  let sheet = ss.getSheetByName(SHEETS.MENU_REVENUE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.MENU_REVENUE);
  }
  const currentHeaderValues = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  const currentHeaders = currentHeaderValues.slice(0, MENU_REVENUE_HEADERS.length);
  if (sheet.getMaxColumns() < MENU_REVENUE_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), MENU_REVENUE_HEADERS.length - sheet.getMaxColumns());
  }
  if (String(currentHeaders[3] || '').trim() === '売上金額') {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 5)).getValues();
      const migrated = values.map(function (row) {
        const count = Math.max(1, Number(row[2]) || 1);
        const legacyTotal = Math.max(0, Number(row[3]) || 0);
        const unitPrice = count > 0 ? Math.round(legacyTotal / count) : legacyTotal;
        return [row[0], row[1], count, unitPrice, 0, row[4]];
      });
      sheet.getRange(2, 1, migrated.length, MENU_REVENUE_HEADERS.length).setValues(migrated);
    }
  } else if (String(currentHeaders[4] || '').trim() === 'メモ' && String(currentHeaders[5] || '').trim() !== 'メモ') {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 5)).getValues();
      const migrated = values.map(function (row) {
        return [row[0], row[1], row[2], row[3], 0, row[4]];
      });
      sheet.getRange(2, 1, migrated.length, MENU_REVENUE_HEADERS.length).setValues(migrated);
    }
  }
  const headerRange = sheet.getRange(1, 1, 1, MENU_REVENUE_HEADERS.length);
  headerRange.setValues([MENU_REVENUE_HEADERS]);
  styleHeader(sheet, MENU_REVENUE_HEADERS.length, '#8d6e63');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 260);
  return sheet;
}

function sanitizeMenuRevenueRecord_(data) {
  const count = Math.max(1, Number(data && data.count) || 1);
  const unitPrice = Math.max(0, Number(data && (data.unitPrice !== undefined ? data.unitPrice : data.amount)) || 0);
  const unitCost = Math.max(0, Number(data && data.unitCost) || 0);
  const totalAmount = count * unitPrice;
  const totalCost = count * unitCost;
  return {
    date: normalizeDateOnlyValue_(data && data.date),
    menuType: normalizeMenuRevenueType_(data && data.menuType),
    count: count,
    unitPrice: unitPrice,
    unitCost: unitCost,
    totalAmount: totalAmount,
    totalCost: totalCost,
    profit: totalAmount - totalCost,
    note: String((data && data.note) || '').trim()
  };
}


function getMenuRevenueRecords() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.MENU_REVENUE);
    if (!sheet) return { status: 'ok', records: [] };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', records: [] };

    const data = sheet.getRange(2, 1, lastRow - 1, MENU_REVENUE_HEADERS.length).getValues();
    const records = data.map(function (row, index) {
      const record = sanitizeMenuRevenueRecord_({
        date: row[0],
        menuType: row[1],
        count: row[2],
        unitPrice: row[3],
        unitCost: row[4],
        note: row[5]
      });
      return {
        rowIdx: index + 2,
        date: record.date,
        menuType: record.menuType,
        count: record.count,
        unitPrice: record.unitPrice,
        unitCost: record.unitCost,
        totalAmount: record.totalAmount,
        totalCost: record.totalCost,
        profit: record.profit,
        note: record.note
      };
    }).filter(function (record) {
      return record.date && record.menuType;
    }).sort(function (a, b) {
      return new Date(b.date).getTime() - new Date(a.date).getTime();
    });

    return { status: 'ok', records: records };
  } catch (err) {
    Logger.log('getMenuRevenueRecords error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function handleSaveMenuRevenueRecord(data) {
  try {
    const recordsInput = Array.isArray(data && data.records) && data.records.length ? data.records : [data];
    const records = recordsInput.map(function (item) {
      const rawRowIdx = item && item.rowIdx !== undefined ? item.rowIdx : (data && data.rowIdx);
      const hasRowIdx = rawRowIdx !== undefined && rawRowIdx !== null && rawRowIdx !== '';
      return {
        hasRowIdx: hasRowIdx,
        rowIdx: hasRowIdx ? Number(rawRowIdx) : 0,
        record: sanitizeMenuRevenueRecord_({
          date: item && item.date ? item.date : (data && data.date),
          menuType: item && item.menuType,
          count: item && item.count,
          unitPrice: item && item.unitPrice,
          unitCost: item && item.unitCost,
          amount: item && item.amount,
          note: item && item.note
        })
      };
    });
    if (!records.length) return { status: 'error', message: '保存対象がありません' };
    for (let i = 0; i < records.length; i++) {
      const payload = records[i];
      const record = payload.record;
      if (!record.date) return { status: 'error', message: (i + 1) + '件目の記録日が必要です' };
      if (!record.menuType) return { status: 'error', message: (i + 1) + '件目のメニュー種別が不正です' };
      if (record.unitPrice <= 0) return { status: 'error', message: (i + 1) + '件目の単価を入力してください' };
      if (payload.hasRowIdx && (!payload.rowIdx || payload.rowIdx < 2)) {
        return { status: 'error', message: (i + 1) + '件目の更新対象が不正です' };
      }
    }

    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateMenuRevenueSheet_(ss);
    const lastRow = sheet.getLastRow();
    const appendValues = [];
    let updatedCount = 0;
    let createdCount = 0;

    records.forEach(function (payload, index) {
      const record = payload.record;
      const values = [[record.date, record.menuType, record.count, record.unitPrice, record.unitCost, record.note]];
      if (payload.hasRowIdx) {
        if (payload.rowIdx > lastRow) {
          throw new Error((index + 1) + '件目の更新対象が見つかりません');
        }
        sheet.getRange(payload.rowIdx, 1, 1, MENU_REVENUE_HEADERS.length).setValues(values);
        updatedCount += 1;
      } else {
        appendValues.push(values[0]);
        createdCount += 1;
      }
    });

    if (appendValues.length) {
      sheet.getRange(lastRow + 1, 1, appendValues.length, MENU_REVENUE_HEADERS.length).setValues(appendValues);
    }
    return { status: 'ok', savedCount: records.length, updatedCount: updatedCount, createdCount: createdCount };
  } catch (err) {
    Logger.log('handleSaveMenuRevenueRecord error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function handleDeleteMenuRevenueRecord(data) {
  try {
    const rowIdx = Number(data && data.rowIdx);
    if (!rowIdx || rowIdx < 2) return { status: 'error', message: '削除対象が不正です' };
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.MENU_REVENUE);
    if (!sheet) return { status: 'error', message: 'メニュー収益シートが見つかりません' };
    sheet.deleteRow(rowIdx);
    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleDeleteMenuRevenueRecord error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function getProductRevenueMasterMap_() {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PRODUCTS);
  const costMap = getProductCostMap_();
  const productMap = {};
  if (!sheet) return productMap;

  ensureProductSheetStructure_(sheet);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const name = String(row[1] || '').trim();
    if (!name) continue;
    productMap[name] = {
      name: name,
      price: Number(row[2] || 0),
      costPrice: Number(costMap[name] || 0),
      status: String(row[5] || '公開')
    };
  }
  return productMap;
}

function getOrCreateProductRevenueSheet_(ss) {
  let sheet = ss.getSheetByName(SHEETS.PRODUCT_REVENUE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.PRODUCT_REVENUE);
  }
  const headerRange = sheet.getRange(1, 1, 1, PRODUCT_REVENUE_HEADERS.length);
  if (sheet.getMaxColumns() < PRODUCT_REVENUE_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), PRODUCT_REVENUE_HEADERS.length - sheet.getMaxColumns());
  }
  headerRange.setValues([PRODUCT_REVENUE_HEADERS]);
  styleHeader(sheet, PRODUCT_REVENUE_HEADERS.length, '#5d8a73');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 260);
  return sheet;
}

function sanitizeProductRevenueRecord_(data, productMap) {
  const lookup = productMap || {};
  const productName = String(data && (data.productName || data.name) || '').trim();
  const master = lookup[productName] || null;
  const qty = Math.max(1, Number(data && data.qty) || 1);
  let unitPrice = Math.max(0, Number(
    data && data.unitPrice !== undefined ? data.unitPrice : (master ? master.price : 0)
  ) || 0);
  let unitCost = Math.max(0, Number(
    data && data.unitCost !== undefined ? data.unitCost : (master ? master.costPrice : 0)
  ) || 0);
  if (isDashiProductName_(productName) && unitCost <= 0) {
    unitCost = DASHI_COST_PER_ITEM;
  }

  let totalAmount = qty * unitPrice;
  let totalCost = qty * unitCost;
  let profit = totalAmount - totalCost;
  let priceBreakdown = null;
  let priceBreakdownText = '';
  if (isDashiProductName_(productName)) {
    const pricing = calculateDashiPricing_(qty);
    unitPrice = pricing.avgUnitPrice;
    totalAmount = pricing.totalRevenue;
    totalCost = qty * unitCost;
    profit = totalAmount - totalCost;
    priceBreakdown = pricing.priceBreakdown;
    priceBreakdownText = formatDashiPriceBreakdownText_(priceBreakdown);
  }

  return {
    date: normalizeDateOnlyValue_(data && data.date),
    productName: productName,
    qty: qty,
    unitPrice: unitPrice,
    unitCost: unitCost,
    totalAmount: totalAmount,
    totalCost: totalCost,
    profit: profit,
    priceBreakdown: priceBreakdown,
    priceBreakdownText: priceBreakdownText,
    note: String((data && data.note) || '').trim()
  };
}

function getProductRevenueRecords() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.PRODUCT_REVENUE);
    if (!sheet) return { status: 'ok', records: [] };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', records: [] };

    const data = sheet.getRange(2, 1, lastRow - 1, PRODUCT_REVENUE_HEADERS.length).getValues();
    const records = data.map(function (row, index) {
      const record = sanitizeProductRevenueRecord_({
        date: row[0],
        productName: row[1],
        qty: row[2],
        unitPrice: row[3],
        unitCost: row[4],
        note: row[5]
      });
      return {
        rowIdx: index + 2,
        date: record.date,
        productName: record.productName,
        qty: record.qty,
        unitPrice: record.unitPrice,
        unitCost: record.unitCost,
        totalAmount: record.totalAmount,
        totalCost: record.totalCost,
        profit: record.profit,
        priceBreakdownText: record.priceBreakdownText,
        note: record.note
      };
    }).filter(function (record) {
      return record.date && record.productName;
    }).sort(function (a, b) {
      return new Date(b.date).getTime() - new Date(a.date).getTime();
    });

    return { status: 'ok', records: records };
  } catch (err) {
    Logger.log('getProductRevenueRecords error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function handleSaveProductRevenueRecord(data) {
  try {
    const recordsInput = Array.isArray(data && data.records) && data.records.length ? data.records : [data];
    const productMap = getProductRevenueMasterMap_();
    const records = recordsInput.map(function (item) {
      const rawRowIdx = item && item.rowIdx !== undefined ? item.rowIdx : (data && data.rowIdx);
      const hasRowIdx = rawRowIdx !== undefined && rawRowIdx !== null && rawRowIdx !== '';
      return {
        hasRowIdx: hasRowIdx,
        rowIdx: hasRowIdx ? Number(rawRowIdx) : 0,
        record: sanitizeProductRevenueRecord_({
          date: item && item.date ? item.date : (data && data.date),
          productName: item && (item.productName || item.name),
          qty: item && item.qty,
          unitPrice: item && item.unitPrice,
          unitCost: item && item.unitCost,
          note: item && item.note
        }, productMap)
      };
    });
    if (!records.length) return { status: 'error', message: '保存対象がありません' };

    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateProductRevenueSheet_(ss);
    const lastRow = sheet.getLastRow();

    for (let i = 0; i < records.length; i++) {
      const payload = records[i];
      const record = payload.record;
      if (!record.date) return { status: 'error', message: (i + 1) + '件目の記録日が必要です' };
      if (!record.productName) return { status: 'error', message: (i + 1) + '件目の商品名を選択してください' };
      if (record.unitPrice <= 0) return { status: 'error', message: (i + 1) + '件目の単価を入力してください' };
      if (payload.hasRowIdx && (!payload.rowIdx || payload.rowIdx < 2)) {
        return { status: 'error', message: (i + 1) + '件目の更新対象が不正です' };
      }
      const canUseExistingRowProduct = payload.hasRowIdx
        && payload.rowIdx <= lastRow
        && String(sheet.getRange(payload.rowIdx, 2).getValue() || '').trim() === record.productName;
      if (!productMap[record.productName] && !canUseExistingRowProduct) {
        return { status: 'error', message: (i + 1) + '件目の商品が商品管理に見つかりません' };
      }
    }

    const appendValues = [];
    let updatedCount = 0;
    let createdCount = 0;

    records.forEach(function (payload, index) {
      const record = payload.record;
      const values = [[record.date, record.productName, record.qty, record.unitPrice, record.unitCost, record.note]];
      if (payload.hasRowIdx) {
        if (payload.rowIdx > lastRow) {
          throw new Error((index + 1) + '件目の更新対象が見つかりません');
        }
        sheet.getRange(payload.rowIdx, 1, 1, PRODUCT_REVENUE_HEADERS.length).setValues(values);
        updatedCount += 1;
      } else {
        appendValues.push(values[0]);
        createdCount += 1;
      }
    });

    if (appendValues.length) {
      sheet.getRange(lastRow + 1, 1, appendValues.length, PRODUCT_REVENUE_HEADERS.length).setValues(appendValues);
    }
    return { status: 'ok', savedCount: records.length, updatedCount: updatedCount, createdCount: createdCount };
  } catch (err) {
    Logger.log('handleSaveProductRevenueRecord error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function handleDeleteProductRevenueRecord(data) {
  try {
    const rowIdx = Number(data && data.rowIdx);
    if (!rowIdx || rowIdx < 2) return { status: 'error', message: '削除対象が不正です' };
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.PRODUCT_REVENUE);
    if (!sheet) return { status: 'error', message: '商品収益シートが見つかりません' };
    sheet.deleteRow(rowIdx);
    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleDeleteProductRevenueRecord error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function rebuildCombinedAnalyticsTotals_(monthStats) {
  if (!monthStats) return;
  monthStats.combinedRevenue = Number(monthStats.sales || 0) + Number(monthStats.menuRevenueTotal || 0);
  monthStats.combinedCost = Number(monthStats.cost || 0) + Number(monthStats.menuCostTotal || 0);
  monthStats.combinedProfit = Number(monthStats.profit || 0) + Number(monthStats.menuProfitTotal || 0);
  monthStats.combinedSales = monthStats.combinedRevenue;
}

function buildRegistrationRouteAnalytics_() {
  const routeLabels = ['新規登録', '復元', '引き継ぎコード利用', '重複候補からの復旧', '自動復旧'];
  const result = {
    routeLabels: routeLabels,
    months: [],
    totals: {},
    matrix: {}
  };
  routeLabels.forEach(function (label) {
    result.totals[label] = 0;
  });

  const ss = getOrCreateSpreadsheet();
  const sheet = getOrCreateUsersSheet_(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;

  const values = sheet.getRange(2, 1, lastRow - 1, USER_HEADERS.length).getValues();
  values.forEach(function (row) {
    if (String(row[USER_COL.DELETE_STATUS - 1] || '').trim() === SOFT_DELETE_STATUS) return;
    const routeInfo = buildUserRegistrationSourceFromRow_(row);
    const source = routeLabels.indexOf(routeInfo.source) !== -1 ? routeInfo.source : '新規登録';
    const occurredAt = routeInfo.source === '新規登録'
      ? formatMaybeDateTime_(row[USER_COL.TIMESTAMP - 1])
      : routeInfo.updatedAt;
    const monthKey = String(occurredAt || '').slice(0, 7);
    if (!monthKey) return;
    if (!result.matrix[monthKey]) {
      result.matrix[monthKey] = {};
      routeLabels.forEach(function (label) {
        result.matrix[monthKey][label] = 0;
      });
    }
    result.matrix[monthKey][source] += 1;
    result.totals[source] += 1;
  });

  result.months = Object.keys(result.matrix).sort().reverse();
  return result;
}

function sortCategoryUsageEntries_(counts) {
  return Object.keys(counts || {}).map(function (category) {
    return {
      category: category,
      count: Number(counts[category] || 0)
    };
  }).sort(function (a, b) {
    if (Number(b.count || 0) !== Number(a.count || 0)) {
      return Number(b.count || 0) - Number(a.count || 0);
    }
    return String(a.category || '').localeCompare(String(b.category || ''), 'ja');
  });
}

function buildCategoryUsageAnalytics_() {
  const result = {
    news: { label: 'NEWS', items: [], total: 0 },
    products: { label: '商品', items: [], total: 0 },
    menus: { label: 'メニュー', items: [], total: 0 },
    calendar: { label: 'カレンダー', items: [], total: 0 }
  };

  const bump = function (bucket, category) {
    const normalized = String(category || '').trim() || 'カテゴリ未設定';
    if (!bucket[normalized]) bucket[normalized] = 0;
    bucket[normalized] += 1;
  };

  const newsCounts = {};
  const news = getAdminBlogs();
  if (news && news.status === 'ok') {
    (news.blogs || []).forEach(function (item) {
      bump(newsCounts, item.category);
    });
  }

  const productCounts = {};
  const products = getAdminProducts();
  if (products && products.status === 'ok') {
    (products.products || []).forEach(function (item) {
      bump(productCounts, item.category);
    });
  }

  const menuCounts = {};
  const menus = getAdminMenus();
  if (menus && menus.status === 'ok') {
    (menus.menus || []).forEach(function (item) {
      bump(menuCounts, item.category);
    });
  }

  const calendarCounts = {};
  const ss = getOrCreateSpreadsheet();
  const calendarSheet = ss.getSheetByName(SHEETS.CALENDAR);
  if (calendarSheet && calendarSheet.getLastRow() >= 2) {
    ensureUpdatedAtColumn_(calendarSheet, '更新日時');
    ensureSortOrderColumn_(calendarSheet);
    ensureNoticeVisibilityColumn_(calendarSheet, 5, '公開');
    ensureSoftDeleteColumns_(calendarSheet);
    const categoryCol = ensureNamedColumn_(calendarSheet, 'カテゴリ', 140);
    const deleteCols = ensureSoftDeleteColumns_(calendarSheet);
    const values = calendarSheet.getRange(2, 1, calendarSheet.getLastRow() - 1, calendarSheet.getLastColumn()).getValues();
    values.forEach(function (row) {
      if (isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol)) return;
      if (!String(row[1] || '').trim()) return;
      bump(calendarCounts, row[categoryCol - 1]);
    });
  }

  result.news.items = sortCategoryUsageEntries_(newsCounts);
  result.products.items = sortCategoryUsageEntries_(productCounts);
  result.menus.items = sortCategoryUsageEntries_(menuCounts);
  result.calendar.items = sortCategoryUsageEntries_(calendarCounts);
  result.news.total = result.news.items.reduce(function (sum, item) { return sum + Number(item.count || 0); }, 0);
  result.products.total = result.products.items.reduce(function (sum, item) { return sum + Number(item.count || 0); }, 0);
  result.menus.total = result.menus.items.reduce(function (sum, item) { return sum + Number(item.count || 0); }, 0);
  result.calendar.total = result.calendar.items.reduce(function (sum, item) { return sum + Number(item.count || 0); }, 0);
  return result;
}

function getAnalyticsData() {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ORDERS);
  const menuSheet = ss.getSheetByName(SHEETS.MENU_REVENUE);
  const productRevenueSheet = ss.getSheetByName(SHEETS.PRODUCT_REVENUE);
  if (!sheet && !menuSheet && !productRevenueSheet) {
    return {
      status: 'ok',
      months: [],
      products: [],
      menuTypes: MENU_REVENUE_TYPES.slice(),
      matrix: {},
      registrationRoutes: buildRegistrationRouteAnalytics_(),
      categoryUsage: buildCategoryUsageAnalytics_()
    };
  }

  const summaryMenuTypes = ['母乳外来', 'ビジリス', '教室', 'その他'];

  const stats = {};
  const productSet = new Set();
  const ensureMonthStats = function (monthKey) {
    if (!stats[monthKey]) {
      stats[monthKey] = {
        products: {},
        sales: 0,
        cost: 0,
        profit: 0,
        menuRevenue: summaryMenuTypes.reduce(function (acc, type) {
          acc[type] = 0;
          return acc;
        }, {}),
        menuCost: summaryMenuTypes.reduce(function (acc, type) {
          acc[type] = 0;
          return acc;
        }, {}),
        menuProfit: summaryMenuTypes.reduce(function (acc, type) {
          acc[type] = 0;
          return acc;
        }, {}),
        menuCount: summaryMenuTypes.reduce(function (acc, type) {
          acc[type] = 0;
          return acc;
        }, {}),
        menuBreakdown: summaryMenuTypes.reduce(function (acc, type) {
          acc[type] = {};
          return acc;
        }, {}),
        menuDetails: {}, // 追加: 全てのメニュー種別ごとの詳細（内訳用）
        menuRevenueTotal: 0,
        menuCostTotal: 0,
        menuProfitTotal: 0,
        combinedRevenue: 0,
        combinedCost: 0,
        combinedProfit: 0,
        combinedSales: 0,
        productDetails: {} // 追加: 売価別の内訳を保持
      };
    }
    return stats[monthKey];
  };

  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const date = row[1];
        const pName = String(row[3] || '').trim();
        const qty = Number(row[4] || 0);
        const status = String(row[11] || '').trim();

        if (status === 'キャンセル済') continue;
        if (!date || isNaN(new Date(date).getTime())) continue;

        const d = new Date(date);
        const monthKey = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
        const monthStats = ensureMonthStats(monthKey);

        if (!monthStats.products[pName]) {
          monthStats.products[pName] = { qty: 0, sales: 0, cost: 0, profit: 0 };
        }
        
        const unitPrice = Number(row[5] || 0); // 売価
        const costPerItem = Number(row[6] || 0);
        const subtotal = Number(row[7] || 0);
        const netProfit = Number(row[8] || 0);

        // -- 商品名ごとのサマリー --
        monthStats.products[pName].qty += qty;
        monthStats.products[pName].sales += subtotal;
        monthStats.products[pName].cost += (costPerItem * qty);
        monthStats.products[pName].profit += netProfit;

        // -- 売価別の詳細内訳 --
        if (isDashiProductName_(pName)) {
          const effectiveCost = costPerItem > 0 ? costPerItem : DASHI_COST_PER_ITEM;
          const breakdown = buildDashiTierBreakdown_(qty);
          ensureDashiDetailRows_(monthStats.productDetails, pName);
          DASHI_PRICE_TIERS.forEach(function (tier) {
            const tierQty = breakdown[tier.key] || 0;
            if (!tierQty) return;
            const detailKey = getDashiDetailKey_(pName, tier.key);
            monthStats.productDetails[detailKey].qty += tierQty;
            monthStats.productDetails[detailKey].sales += tier.price * tierQty;
            monthStats.productDetails[detailKey].cost += effectiveCost * tierQty;
            monthStats.productDetails[detailKey].profit += (tier.price - effectiveCost) * tierQty;
          });
        } else {
          const detailKey = pName + '|' + unitPrice;
          if (!monthStats.productDetails[detailKey]) {
            monthStats.productDetails[detailKey] = {
              name: pName,
              price: unitPrice,
              priceLabel: '',
              sortOrder: 1000 - unitPrice,
              qty: 0,
              sales: 0,
              cost: 0,
              profit: 0
            };
          }
          monthStats.productDetails[detailKey].qty += qty;
          monthStats.productDetails[detailKey].sales += subtotal;
          monthStats.productDetails[detailKey].cost += (costPerItem * qty);
          monthStats.productDetails[detailKey].profit += netProfit;
        }

        // -- 月間全体のサマリー --
        monthStats.sales += subtotal;
        monthStats.cost += (costPerItem * qty);
        monthStats.profit += netProfit;
        rebuildCombinedAnalyticsTotals_(monthStats);

        if (pName) productSet.add(pName);
      }
    }
  }

  if (menuSheet) {
    const menuLastRow = menuSheet.getLastRow();
    if (menuLastRow >= 2) {
      const menuData = menuSheet.getRange(2, 1, menuLastRow - 1, MENU_REVENUE_HEADERS.length).getValues();
      for (let i = 0; i < menuData.length; i++) {
        const record = sanitizeMenuRevenueRecord_({
          date: menuData[i][0],
          menuType: menuData[i][1],
          count: menuData[i][2],
          unitPrice: menuData[i][3],
          unitCost: menuData[i][4],
          note: menuData[i][5]
        });
        if (!record.date || !record.menuType || record.totalAmount <= 0) continue;

        const d = new Date(record.date);
        if (isNaN(d.getTime())) continue;

        const monthKey = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
        const monthStats = ensureMonthStats(monthKey);
        
        // --- 集計用カテゴリの判定 ---
        let summaryType = record.menuType;
        if (summaryMenuTypes.slice(0, 3).indexOf(summaryType) === -1) {
          summaryType = 'その他';
        }

        monthStats.menuRevenue[summaryType] += record.totalAmount;
        monthStats.menuCost[summaryType] += record.totalCost;
        monthStats.menuProfit[summaryType] += record.profit;
        monthStats.menuCount[summaryType] += record.count;
        const unitKey = String(record.unitPrice);
        if (!monthStats.menuBreakdown[summaryType][unitKey]) {
          monthStats.menuBreakdown[summaryType][unitKey] = { count: 0, revenue: 0, cost: 0, profit: 0 };
        }
        monthStats.menuBreakdown[summaryType][unitKey].count += record.count;
        monthStats.menuBreakdown[summaryType][unitKey].revenue += record.totalAmount;
        monthStats.menuBreakdown[summaryType][unitKey].cost += record.totalCost;
        monthStats.menuBreakdown[summaryType][unitKey].profit += record.profit;

        // --- 詳細データ（メニュー名別）の保持 ---
        if (!monthStats.menuDetails[record.menuType]) {
          monthStats.menuDetails[record.menuType] = { name: record.menuType, count: 0, revenue: 0, cost: 0, profit: 0, breakdown: {} };
        }
        monthStats.menuDetails[record.menuType].count += record.count;
        monthStats.menuDetails[record.menuType].revenue += record.totalAmount;
        monthStats.menuDetails[record.menuType].cost += record.totalCost;
        monthStats.menuDetails[record.menuType].profit += record.profit;
        if (!monthStats.menuDetails[record.menuType].breakdown[unitKey]) {
          monthStats.menuDetails[record.menuType].breakdown[unitKey] = { count: 0, revenue: 0, cost: 0, profit: 0 };
        }
        monthStats.menuDetails[record.menuType].breakdown[unitKey].count += record.count;
        monthStats.menuDetails[record.menuType].breakdown[unitKey].revenue += record.totalAmount;
        monthStats.menuDetails[record.menuType].breakdown[unitKey].cost += record.totalCost;
        monthStats.menuDetails[record.menuType].breakdown[unitKey].profit += record.profit;

        monthStats.menuRevenueTotal += record.totalAmount;
        monthStats.menuCostTotal += record.totalCost;
        monthStats.menuProfitTotal += record.profit;
        rebuildCombinedAnalyticsTotals_(monthStats);
      }
    }
  }

  if (productRevenueSheet) {
    const productRevenueLastRow = productRevenueSheet.getLastRow();
    if (productRevenueLastRow >= 2) {
      const productRevenueData = productRevenueSheet.getRange(2, 1, productRevenueLastRow - 1, PRODUCT_REVENUE_HEADERS.length).getValues();
      for (let i = 0; i < productRevenueData.length; i++) {
        const record = sanitizeProductRevenueRecord_({
          date: productRevenueData[i][0],
          productName: productRevenueData[i][1],
          qty: productRevenueData[i][2],
          unitPrice: productRevenueData[i][3],
          unitCost: productRevenueData[i][4],
          note: productRevenueData[i][5]
        });
        if (!record.date || !record.productName || record.totalAmount <= 0) continue;

        const d = new Date(record.date);
        if (isNaN(d.getTime())) continue;

        const monthKey = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
        const monthStats = ensureMonthStats(monthKey);

        if (!monthStats.products[record.productName]) {
          monthStats.products[record.productName] = { qty: 0, sales: 0, cost: 0, profit: 0 };
        }

        monthStats.products[record.productName].qty += record.qty;
        monthStats.products[record.productName].sales += record.totalAmount;
        monthStats.products[record.productName].cost += record.totalCost;
        monthStats.products[record.productName].profit += record.profit;

        if (isDashiProductName_(record.productName)) {
          ensureDashiDetailRows_(monthStats.productDetails, record.productName);
          const breakdown = record.priceBreakdown || buildDashiTierBreakdown_(record.qty);
          DASHI_PRICE_TIERS.forEach(function (tier) {
            const tierQty = Number(breakdown[tier.key] || 0);
            if (!tierQty) return;
            const detailKey = getDashiDetailKey_(record.productName, tier.key);
            monthStats.productDetails[detailKey].qty += tierQty;
            monthStats.productDetails[detailKey].sales += tier.price * tierQty;
            monthStats.productDetails[detailKey].cost += record.unitCost * tierQty;
            monthStats.productDetails[detailKey].profit += (tier.price - record.unitCost) * tierQty;
          });
        } else {
          const detailKey = record.productName + '|' + String(record.unitPrice);
          if (!monthStats.productDetails[detailKey]) {
            monthStats.productDetails[detailKey] = {
              name: record.productName,
              price: record.unitPrice,
              priceLabel: '',
              sortOrder: 1000 - record.unitPrice,
              qty: 0,
              sales: 0,
              cost: 0,
              profit: 0
            };
          }
          monthStats.productDetails[detailKey].qty += record.qty;
          monthStats.productDetails[detailKey].sales += record.totalAmount;
          monthStats.productDetails[detailKey].cost += record.totalCost;
          monthStats.productDetails[detailKey].profit += record.profit;
        }

        monthStats.sales += record.totalAmount;
        monthStats.cost += record.totalCost;
        monthStats.profit += record.profit;
        rebuildCombinedAnalyticsTotals_(monthStats);

        productSet.add(record.productName);
      }
    }
  }

  return {
    status: 'ok',
    months: Object.keys(stats).sort().reverse(),
    products: Array.from(productSet).sort(),
    menuTypes: summaryMenuTypes,
    matrix: stats,
    registrationRoutes: buildRegistrationRouteAnalytics_(),
    categoryUsage: buildCategoryUsageAnalytics_()
  };
}

/**
 * 「天然だし調味粉」専用の価格・利益計算ロジック
 * 収益計算、注文作成、管理画面での売上集計すべてで共通で使用する
 */
function calculateDashiPricing_(qty) {
  const count = Math.max(0, Number(qty) || 0);
  if (count <= 0) {
    return { totalRevenue: 0, avgUnitPrice: DASHI_BASE_PRICE, totalCost: 0, profit: 0, priceBreakdown: buildDashiTierBreakdown_(0) };
  }

  const breakdown = buildDashiTierBreakdown_(count);
  let totalRevenue = 0;
  DASHI_PRICE_TIERS.forEach(function (tier) {
    totalRevenue += tier.price * (breakdown[tier.key] || 0);
  });

  const roundedTotal = Math.floor(totalRevenue);
  const avgUnitPrice = roundedTotal / count; // スプレッドシート記録用の実質単価
  const totalCost = DASHI_COST_PER_ITEM * count;
  const profit = roundedTotal - totalCost;

  return {
    totalRevenue: roundedTotal,
    avgUnitPrice: avgUnitPrice,
    totalCost: totalCost,
    profit: profit,
    priceBreakdown: breakdown
  };
}

// ========== ドライブのアクセス権を明示的に要求するためのダミー関数 ==========
// ※ この関数はアプリから呼び出されませんが、GASエディタ上で一度だけ手動実行することで、
// Googleドライブへのアクセス権限（DriveApp）をユーザーに許可させるために必要です。
function authorizeDriveAccess() {
  try {
    const folders = DriveApp.getFoldersByName('まゆみ助産院_画像');
    Logger.log('Driveアクセス成功');
  } catch (e) {
    Logger.log('Driveアクセスエラー: ' + e.toString());
  }
}

// ========== 管理者用：カレンダー取得 ==========

function getAdminCalendar() {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CALENDAR);
  if (!sheet) return { status: 'ok', events: [] };
  ensureUpdatedAtColumn_(sheet, '更新日時');

  const data = sheet.getDataRange().getValues();
  const events = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || !row[1]) continue; // 必須項目

    const dateVal = row[0];
    let dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy-MM-dd');
    } else {
      dateStr = String(dateVal).trim();
    }

    events.push({
      rowIdx: i + 1,
      date: dateStr,
      title: String(row[1]),
      desc: String(row[2] || ''),
      color: String(row[3] || '#e57373'),
      status: String(row[4] || '公開'), // E列: 公開設定
      image: String(row[5] || ''), // F列: 画像URL
      updatedAt: formatMaybeDateTime_(row[6])
    });
  }

  // 日付の降順
  events.sort((a, b) => b.date.localeCompare(a.date));
  return { status: 'ok', events: events };
}

// ========== 管理者用：カレンダー追加 ==========

function handleAddCalendar(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    let sheet = ss.getSheetByName(SHEETS.CALENDAR);
    if (!sheet) {
      getCalendarEvents();
      sheet = ss.getSheetByName(SHEETS.CALENDAR);
    }
    ensureUpdatedAtColumn_(sheet, '更新日時');
    const noticeCol = ensureNoticeVisibilityColumn_(sheet, 5, '公開');
    const publishAtCol = ensurePublishAtColumn_(sheet);
    const categoryCol = ensureNamedColumn_(sheet, 'カテゴリ', 140);

    const rowsToAdd = [];
    const updatedAt = formatDateTime_(new Date());

    if (data.dates && Array.isArray(data.dates) && data.dates.length > 0) {
      data.dates.forEach(dStr => {
        rowsToAdd.push([
          dStr,
          data.title,
          data.desc || '',
          data.color || '#e57373',
          data.status || '公開',
          data.image || '',
          updatedAt
        ]);
      });
    } else {
      rowsToAdd.push([
        data.date,
        data.title,
        data.desc || '',
        data.color || '#e57373',
        data.status || '公開',
        data.image || '',
        updatedAt
      ]);
    }

    if (rowsToAdd.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 7).setValues(rowsToAdd);
      const startRow = sheet.getLastRow() - rowsToAdd.length + 1;
      const noticeValues = rowsToAdd.map(function () {
        return [normalizePublishVisibilityStatus_(data.noticeStatus || data.status || '公開')];
      });
      sheet.getRange(startRow, noticeCol, noticeValues.length, 1).setValues(noticeValues);
      if (categoryCol) {
        const categoryValues = rowsToAdd.map(function () {
          return [String(data.category || '').trim()];
        });
        sheet.getRange(startRow, categoryCol, categoryValues.length, 1).setValues(categoryValues);
      }
      if (publishAtCol) {
        const publishValues = rowsToAdd.map(function () {
          return [normalizePublishAtValue_(data.publishAt)];
        });
        sheet.getRange(startRow, publishAtCol, publishValues.length, 1).setValues(publishValues);
      }
    }

    // 自動プッシュ通知
    if (String(data.status || '公開') !== '非公開' && isPublishAtAvailable_(data.publishAt)) {
      sendAutoPush('📅 ' + (data.title || '新しいイベント'), 'カレンダーが更新されました', {
        targetPage: 'calendar'
      });
    }

    return { status: 'ok', count: rowsToAdd.length };
  } catch (err) {
    Logger.log('handleAddCalendar error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

// ========== 管理者用：カレンダー更新 ==========

function handleUpdateCalendar(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.CALENDAR);
    if (!sheet) return { status: 'error', message: 'シートが見つかりません' };
    ensureUpdatedAtColumn_(sheet, '更新日時');
    const noticeCol = ensureNoticeVisibilityColumn_(sheet, 5, '公開');
    const publishAtCol = ensurePublishAtColumn_(sheet);
    const categoryCol = ensureNamedColumn_(sheet, 'カテゴリ', 140);

    const rowIdx = Number(data.rowIdx);
    if (rowIdx < 2) return { status: 'error', message: '更新対象が見つかりません' };

    sheet.getRange(rowIdx, 1, 1, 7).setValues([[
      data.date,
      data.title,
      data.desc || '',
      data.color || '#e57373',
      data.status || sheet.getRange(rowIdx, 5).getValue(),
      data.image || '',
      formatDateTime_(new Date())
    ]]);
    if (data.noticeStatus) {
      sheet.getRange(rowIdx, noticeCol).setValue(normalizePublishVisibilityStatus_(data.noticeStatus));
    }
    if (categoryCol && data.category !== undefined) {
      sheet.getRange(rowIdx, categoryCol).setValue(String(data.category || '').trim());
    }
    if (publishAtCol && data.publishAt !== undefined) {
      sheet.getRange(rowIdx, publishAtCol).setValue(normalizePublishAtValue_(data.publishAt));
    }

    // 自動プッシュ通知
    const effectiveStatus = data.status || sheet.getRange(rowIdx, 5).getValue();
    const effectivePublishAt = publishAtCol ? sheet.getRange(rowIdx, publishAtCol).getValue() : '';
    if (String(effectiveStatus || '公開') !== '非公開' && isPublishAtAvailable_(effectivePublishAt)) {
      sendAutoPush('📅 ' + (data.title || 'イベント更新'), 'カレンダーが更新されました', {
        targetPage: 'calendar'
      });
    }

    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleUpdateCalendar error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

// ========== 会員データ関連 ==========

function handleUpdateUser(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);

    const memberId = data.memberId;
    if (!memberId) return { status: 'error', message: 'IDが指定されていません' };

    const targetRow = findUserRowByMemberId_(sheet, memberId);

    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/M/d H:mm');
    const rewardStatus = sanitizeRewardStatus_(data);

    if (targetRow === -1) {
      const rowData = new Array(USER_HEADERS.length).fill('');
      rowData[USER_COL.MEMBER_ID - 1] = memberId;
      rowData[USER_COL.TIMESTAMP - 1] = timestamp;
      rowData[USER_COL.NAME - 1] = data.name || '';
      rowData[USER_COL.KANA - 1] = data.kana || '';
      rowData[USER_COL.PHONE - 1] = data.phone || '';
      rowData[USER_COL.AVATAR_URL - 1] = data.avatar || '';
      rowData[USER_COL.MEMO - 1] = data.memo || '';
      rowData[USER_COL.PUSH - 1] = data.pushSubscription || '';
      rowData[USER_COL.STATUS - 1] = data.status || '';
      rowData[USER_COL.BIRTHDAY - 1] = data.birthday || '';
      rowData[USER_COL.ADDRESS - 1] = data.address || '';
      rowData[USER_COL.PASSCODE - 1] = data.passcode || '';
      rowData[USER_COL.DEVICE_SESSIONS - 1] = '[]';
      rowData[USER_COL.DELETE_STATUS - 1] = '';
      rowData[USER_COL.DELETED_AT - 1] = '';
      rowData[USER_COL.MERGED_INTO - 1] = '';
      clearTransferCodeFromRow_(rowData);
      setUserRegistrationSource_(
        rowData,
        data.registrationSource || '新規登録',
        data.registrationSourceDetail || '',
        data.registrationSourceUpdatedAt || timestamp
      );
      const nextRow = applyRewardStatusToRow_(rowData, rewardStatus);
      sheet.appendRow(nextRow);
    } else {
      const range = sheet.getRange(targetRow, 1, 1, USER_HEADERS.length);
      const currentValues = range.getValues()[0];
      const currentRewardStatus = getRewardStatusFromRow_(currentValues);
      const nextRewardStatus = sanitizeRewardStatus_({
        stampCount: data.stampCount !== undefined ? data.stampCount : currentRewardStatus.stampCount,
        stampCardNum: data.stampCardNum !== undefined ? data.stampCardNum : currentRewardStatus.stampCardNum,
        rewards: data.rewards !== undefined ? data.rewards : currentRewardStatus.rewards,
        stampHistory: data.stampHistory !== undefined ? data.stampHistory : currentRewardStatus.stampHistory,
        lastStampDate: data.lastStampDate !== undefined ? data.lastStampDate : currentRewardStatus.lastStampDate,
        stampAchievedDate: data.stampAchievedDate !== undefined ? data.stampAchievedDate : currentRewardStatus.stampAchievedDate
      });
      const updatedRow = currentValues.slice();
      updatedRow[USER_COL.MEMBER_ID - 1] = memberId;
      updatedRow[USER_COL.TIMESTAMP - 1] = timestamp;
      updatedRow[USER_COL.NAME - 1] = data.name !== undefined ? data.name : currentValues[USER_COL.NAME - 1];
      updatedRow[USER_COL.KANA - 1] = data.kana !== undefined ? data.kana : currentValues[USER_COL.KANA - 1];
      updatedRow[USER_COL.PHONE - 1] = data.phone !== undefined ? data.phone : currentValues[USER_COL.PHONE - 1];
      updatedRow[USER_COL.AVATAR_URL - 1] = data.avatar !== undefined ? data.avatar : currentValues[USER_COL.AVATAR_URL - 1];
      updatedRow[USER_COL.MEMO - 1] = data.memo !== undefined ? data.memo : currentValues[USER_COL.MEMO - 1];
      updatedRow[USER_COL.PUSH - 1] = data.pushSubscription !== undefined ? data.pushSubscription : currentValues[USER_COL.PUSH - 1];
      updatedRow[USER_COL.STATUS - 1] = data.status !== undefined ? data.status : currentValues[USER_COL.STATUS - 1];
      updatedRow[USER_COL.BIRTHDAY - 1] = data.birthday !== undefined ? data.birthday : currentValues[USER_COL.BIRTHDAY - 1];
      updatedRow[USER_COL.ADDRESS - 1] = data.address !== undefined ? data.address : currentValues[USER_COL.ADDRESS - 1];
      updatedRow[USER_COL.PASSCODE - 1] = data.passcode !== undefined ? data.passcode : currentValues[USER_COL.PASSCODE - 1];
      updatedRow[USER_COL.DEVICE_SESSIONS - 1] = currentValues[USER_COL.DEVICE_SESSIONS - 1];
      updatedRow[USER_COL.DELETE_STATUS - 1] = currentValues[USER_COL.DELETE_STATUS - 1];
      updatedRow[USER_COL.DELETED_AT - 1] = currentValues[USER_COL.DELETED_AT - 1];
      updatedRow[USER_COL.MERGED_INTO - 1] = currentValues[USER_COL.MERGED_INTO - 1];
      updatedRow[USER_COL.TRANSFER_CODE - 1] = currentValues[USER_COL.TRANSFER_CODE - 1];
      updatedRow[USER_COL.TRANSFER_CODE_ISSUED_AT - 1] = currentValues[USER_COL.TRANSFER_CODE_ISSUED_AT - 1];
      updatedRow[USER_COL.REGISTRATION_SOURCE - 1] = currentValues[USER_COL.REGISTRATION_SOURCE - 1];
      updatedRow[USER_COL.REGISTRATION_SOURCE_DETAIL - 1] = currentValues[USER_COL.REGISTRATION_SOURCE_DETAIL - 1];
      updatedRow[USER_COL.REGISTRATION_SOURCE_UPDATED_AT - 1] = currentValues[USER_COL.REGISTRATION_SOURCE_UPDATED_AT - 1];
      if (data.passcode !== undefined) {
        clearTransferCodeFromRow_(updatedRow);
      }
      if (data.registrationSource !== undefined || data.registrationSourceDetail !== undefined || data.registrationSourceUpdatedAt !== undefined) {
        setUserRegistrationSource_(
          updatedRow,
          data.registrationSource || currentValues[USER_COL.REGISTRATION_SOURCE - 1] || '新規登録',
          data.registrationSourceDetail !== undefined ? data.registrationSourceDetail : currentValues[USER_COL.REGISTRATION_SOURCE_DETAIL - 1],
          data.registrationSourceUpdatedAt || currentValues[USER_COL.REGISTRATION_SOURCE_UPDATED_AT - 1] || timestamp
        );
      }
      range.setValues([applyRewardStatusToRow_(updatedRow, nextRewardStatus)]);
    }

    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleUpdateUser error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function normalizePhoneForMatch_(value) {
  let str = String(value || '').trim().replace(/\D/g, '');
  if (str.length >= 9 && !str.startsWith('0')) {
    str = '0' + str;
  }
  return str;
}

function normalizeNameForMatch_(value) {
  return String(value || '').trim().replace(/[\s\u3000]+/g, '');
}

function normalizePasscodeForMatch_(value) {
  return String(value || '').trim();
}

function normalizeTransferCodeForMatch_(value) {
  return String(value || '').replace(/\D/g, '').slice(0, TRANSFER_CODE_LENGTH);
}

function clearTransferCodeFromRow_(rowData) {
  if (!rowData) return rowData;
  rowData[USER_COL.TRANSFER_CODE - 1] = '';
  rowData[USER_COL.TRANSFER_CODE_ISSUED_AT - 1] = '';
  return rowData;
}

function normalizeRegistrationSource_(value) {
  const source = String(value || '').trim();
  if (!source) return '新規登録';
  const normalized = source.toLowerCase();
  if (normalized === 'new' || normalized === 'newregistration' || normalized === 'register') return '新規登録';
  if (normalized === 'recover' || normalized === 'identity') return '復元';
  if (normalized === 'transfer' || normalized === 'transfercode') return '引き継ぎコード利用';
  if (normalized === 'merge' || normalized === 'duplicate') return '重複候補からの復旧';
  if (normalized === 'activity' || normalized === 'auto') return '自動復旧';
  return source;
}

function setUserRegistrationSource_(rowData, source, detail, occurredAt) {
  if (!rowData) return rowData;
  rowData[USER_COL.REGISTRATION_SOURCE - 1] = normalizeRegistrationSource_(source);
  rowData[USER_COL.REGISTRATION_SOURCE_DETAIL - 1] = String(detail || '').trim();
  rowData[USER_COL.REGISTRATION_SOURCE_UPDATED_AT - 1] = normalizeDateTimeValue_(occurredAt) || formatDateTime_(new Date());
  return rowData;
}

function buildUserRegistrationSourceFromRow_(row) {
  const source = normalizeRegistrationSource_(row[USER_COL.REGISTRATION_SOURCE - 1]);
  const detail = String(row[USER_COL.REGISTRATION_SOURCE_DETAIL - 1] || '').trim();
  const updatedAt = formatMaybeDateTime_(row[USER_COL.REGISTRATION_SOURCE_UPDATED_AT - 1]) || formatMaybeDateTime_(row[USER_COL.TIMESTAMP - 1]);
  return {
    source: source,
    detail: detail,
    updatedAt: updatedAt
  };
}

function fillUserFieldIfBlank_(rowData, columnIndex, value) {
  if (!rowData || !columnIndex) return false;
  const currentValue = String(rowData[columnIndex - 1] || '').trim();
  const nextValue = String(value || '').trim();
  if (currentValue || !nextValue) return false;
  rowData[columnIndex - 1] = nextValue;
  return true;
}

function ensureUserRowFromActivity_(sheet, data) {
  if (!sheet) return { rowIndex: -1, rowData: null, created: false };

  const memberId = String(data && data.memberId || '').trim();
  if (!memberId) return { rowIndex: -1, rowData: null, created: false };

  const name = String(data && (data.name || data.customerName) || '').trim();
  const phone = String(data && data.phone || '').trim();
  const birthday = normalizeDateOnlyValue_(data && data.birthday);
  const address = String(data && data.address || '').trim();
  const kana = String(data && data.kana || '').trim();
  const avatar = String(data && data.avatar || '').trim();
  const memo = String(data && data.memo || '').trim();
  const pushSubscription = String(data && data.pushSubscription || '').trim();
  const status = String(data && data.status || '').trim();
  const passcode = normalizePasscodeForMatch_(data && data.passcode);
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/M/d H:mm');
  const targetRow = findUserRowByMemberId_(sheet, memberId);

  if (targetRow === -1) {
    const rowData = new Array(USER_HEADERS.length).fill('');
    rowData[USER_COL.MEMBER_ID - 1] = memberId;
    rowData[USER_COL.TIMESTAMP - 1] = timestamp;
    rowData[USER_COL.NAME - 1] = name;
    rowData[USER_COL.KANA - 1] = kana;
    rowData[USER_COL.PHONE - 1] = phone;
    rowData[USER_COL.AVATAR_URL - 1] = avatar;
    rowData[USER_COL.MEMO - 1] = memo;
    rowData[USER_COL.PUSH - 1] = pushSubscription;
    rowData[USER_COL.STATUS - 1] = status;
    rowData[USER_COL.BIRTHDAY - 1] = birthday;
    rowData[USER_COL.ADDRESS - 1] = address;
    rowData[USER_COL.PASSCODE - 1] = passcode;
    rowData[USER_COL.DEVICE_SESSIONS - 1] = '[]';
    rowData[USER_COL.DELETE_STATUS - 1] = '';
    rowData[USER_COL.DELETED_AT - 1] = '';
    rowData[USER_COL.MERGED_INTO - 1] = '';
    clearTransferCodeFromRow_(rowData);
    setUserRegistrationSource_(rowData, '自動復旧', 'スタンプ・注文の活動から再作成', timestamp);
    sheet.appendRow(rowData);
    return { rowIndex: sheet.getLastRow(), rowData: rowData, created: true };
  }

  const range = sheet.getRange(targetRow, 1, 1, USER_HEADERS.length);
  const currentRow = range.getValues()[0];
  const nextRow = currentRow.slice();
  let changed = false;

  changed = fillUserFieldIfBlank_(nextRow, USER_COL.NAME, name) || changed;
  changed = fillUserFieldIfBlank_(nextRow, USER_COL.KANA, kana) || changed;
  changed = fillUserFieldIfBlank_(nextRow, USER_COL.PHONE, phone) || changed;
  changed = fillUserFieldIfBlank_(nextRow, USER_COL.AVATAR_URL, avatar) || changed;
  changed = fillUserFieldIfBlank_(nextRow, USER_COL.MEMO, memo) || changed;
  changed = fillUserFieldIfBlank_(nextRow, USER_COL.PUSH, pushSubscription) || changed;
  changed = fillUserFieldIfBlank_(nextRow, USER_COL.STATUS, status) || changed;
  changed = fillUserFieldIfBlank_(nextRow, USER_COL.BIRTHDAY, birthday) || changed;
  changed = fillUserFieldIfBlank_(nextRow, USER_COL.ADDRESS, address) || changed;

  if (fillUserFieldIfBlank_(nextRow, USER_COL.PASSCODE, passcode)) {
    clearTransferCodeFromRow_(nextRow);
    changed = true;
  }
  if (String(nextRow[USER_COL.DELETE_STATUS - 1] || '').trim()) {
    nextRow[USER_COL.DELETE_STATUS - 1] = '';
    nextRow[USER_COL.DELETED_AT - 1] = '';
    nextRow[USER_COL.MERGED_INTO - 1] = '';
    setUserRegistrationSource_(nextRow, '自動復旧', 'スタンプ・注文の活動から再表示', timestamp);
    changed = true;
  }

  if (changed) {
    range.setValues([nextRow]);
    return { rowIndex: targetRow, rowData: nextRow, created: false };
  }

  return { rowIndex: targetRow, rowData: currentRow, created: false };
}

function getTransferCodeExpiryDate_(issuedAtValue) {
  const normalized = normalizeDateTimeValue_(issuedAtValue);
  if (!normalized) return null;
  const issuedAt = new Date(normalized);
  if (isNaN(issuedAt.getTime())) return null;
  issuedAt.setHours(issuedAt.getHours() + TRANSFER_CODE_TTL_HOURS);
  return issuedAt;
}

function formatTransferCodeDateTime_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

function generateUniqueTransferCode_(values) {
  const usedCodes = {};
  (values || []).forEach(function (row) {
    const code = normalizeTransferCodeForMatch_(row[USER_COL.TRANSFER_CODE - 1]);
    if (code) usedCodes[code] = true;
  });

  for (let i = 0; i < 30; i++) {
    const nextCode = String(Math.floor(Math.random() * Math.pow(10, TRANSFER_CODE_LENGTH))).padStart(TRANSFER_CODE_LENGTH, '0');
    if (!usedCodes[nextCode]) return nextCode;
  }

  return String(new Date().getTime()).slice(-TRANSFER_CODE_LENGTH);
}

function buildRecoverAccountUserFromRow_(row) {
  const rowPasscode = String(row[USER_COL.PASSCODE - 1] || '').trim();
  return {
    memberId: String(row[USER_COL.MEMBER_ID - 1]),
    name: String(row[USER_COL.NAME - 1]),
    kana: String(row[USER_COL.KANA - 1]),
    phone: String(row[USER_COL.PHONE - 1]),
    avatar: String(row[USER_COL.AVATAR_URL - 1]),
    memo: String(row[USER_COL.MEMO - 1]),
    status: String(row[USER_COL.STATUS - 1]),
    birthday: normalizeDateOnlyValue_(row[USER_COL.BIRTHDAY - 1]),
    address: String(row[USER_COL.ADDRESS - 1]),
    passcode: rowPasscode,
    deviceSessions: getUserDeviceSessionsFromRow_(row),
    stampCount: Number(row[USER_COL.STAMP_COUNT - 1]) || 0,
    stampCardNum: Number(row[USER_COL.STAMP_CARD_NUM - 1]) || 1,
    rewards: String(row[USER_COL.REWARDS - 1] || '[]'),
    stampHistory: String(row[USER_COL.STAMP_HISTORY_JSON - 1] || '[]'),
    lastStampDate: String(row[USER_COL.LAST_STAMP_DATE - 1] || ''),
    lastStampAt: String(row[USER_COL.LAST_STAMP_AT - 1] || ''),
    stampAchievedAt: String(row[USER_COL.STAMP_ACHIEVED_AT - 1] || ''),
    regDate: formatDateOnly_(new Date(row[USER_COL.TIMESTAMP - 1]))
  };
}

function getRecoveryCandidates(params) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', candidates: [] };

    const name = normalizeNameForMatch_(params && params.name);
    const phone = normalizePhoneForMatch_(params && params.phone);
    const birthday = normalizeDateOnlyValue_(params && params.birthday);
    if (!name && !phone && !birthday) return { status: 'ok', candidates: [] };

    const values = sheet.getRange(2, 1, lastRow - 1, USER_HEADERS.length).getValues();
    const candidates = values.map(function (row, index) {
      if (String(row[USER_COL.DELETE_STATUS - 1] || '').trim() === SOFT_DELETE_STATUS) return null;
      const rowName = normalizeNameForMatch_(row[USER_COL.NAME - 1]);
      const rowPhone = normalizePhoneForMatch_(row[USER_COL.PHONE - 1]);
      const rowBirthday = normalizeDateOnlyValue_(row[USER_COL.BIRTHDAY - 1]);
      const reasons = [];
      if (name && rowName === name) reasons.push('氏名');
      if (phone && rowPhone && rowPhone === phone) reasons.push('電話番号');
      if (birthday && rowBirthday && rowBirthday === birthday) reasons.push('生年月日');
      if (!reasons.length) return null;
      const score = reasons.reduce(function (sum, label) {
        return sum + (label === '氏名' ? 2 : 3);
      }, 0);
      return {
        rowIdx: index + 2,
        memberId: String(row[USER_COL.MEMBER_ID - 1] || ''),
        name: String(row[USER_COL.NAME - 1] || ''),
        phone: String(row[USER_COL.PHONE - 1] || ''),
        birthday: normalizeDateOnlyValue_(row[USER_COL.BIRTHDAY - 1]),
        updatedAt: formatMaybeDateTime_(row[USER_COL.TIMESTAMP - 1]),
        reasons: reasons,
        score: score
      };
    }).filter(Boolean).sort(function (a, b) {
      if (b.score !== a.score) return b.score - a.score;
      return parseLooseDateToTimestamp_(b.updatedAt) - parseLooseDateToTimestamp_(a.updatedAt);
    }).slice(0, 5);

    return { status: 'ok', candidates: candidates };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleRecoverAccount(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'error', message: '会員が見つかりません' };

    const name = normalizeNameForMatch_(data.name);
    const phone = normalizePhoneForMatch_(data.phone);
    const birthday = normalizeDateOnlyValue_(data.birthday);
    const passcode = normalizePasscodeForMatch_(data.passcode);
    const newPasscode = normalizePasscodeForMatch_(data.newPasscode) || passcode;
    const transferCode = normalizeTransferCodeForMatch_(data.transferCode);

    if (!transferCode && !name) {
      return { status: 'error', message: 'お名前を入力してください。' };
    }
    if (!transferCode && !phone && !birthday && !passcode) {
      return { status: 'error', message: '電話番号・生年月日・現在のパスコードのうち1つ以上を入力してください。' };
    }
    if (!/^(?:\d{4}|\d{6})$/.test(newPasscode)) {
      return { status: 'error', message: 'この端末で使うパスコードを4桁または6桁の数字で入力してください。' };
    }

    const values = sheet.getRange(2, 1, lastRow - 1, USER_HEADERS.length).getValues();
    let matchedRowIndex = -1;
    let matchedRowValues = null;
    let usedTransferCode = false;

    if (transferCode) {
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const rowTransferCode = normalizeTransferCodeForMatch_(row[USER_COL.TRANSFER_CODE - 1]);
        if (rowTransferCode !== transferCode) continue;

        const expiryDate = getTransferCodeExpiryDate_(row[USER_COL.TRANSFER_CODE_ISSUED_AT - 1]);
        if (!expiryDate) {
          const invalidRow = row.slice();
          clearTransferCodeFromRow_(invalidRow);
          sheet.getRange(i + 2, 1, 1, USER_HEADERS.length).setValues([invalidRow]);
          return { status: 'error', message: 'この引き継ぎコードは無効です。元の端末で新しいコードを発行してください。' };
        }
        if (expiryDate.getTime() < new Date().getTime()) {
          const expiredRow = row.slice();
          clearTransferCodeFromRow_(expiredRow);
          sheet.getRange(i + 2, 1, 1, USER_HEADERS.length).setValues([expiredRow]);
          return { status: 'error', message: 'この引き継ぎコードは期限切れです。元の端末で新しいコードを発行してください。' };
        }

        matchedRowIndex = i + 2;
        matchedRowValues = row.slice();
        usedTransferCode = true;
        break;
      }
    } else {
      const candidates = [];
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const rowName = normalizeNameForMatch_(row[USER_COL.NAME - 1]);
        const rowPhone = normalizePhoneForMatch_(row[USER_COL.PHONE - 1]);
        const rowBirthday = normalizeDateOnlyValue_(row[USER_COL.BIRTHDAY - 1]);
        const rowPasscode = normalizePasscodeForMatch_(row[USER_COL.PASSCODE - 1]);

        if (rowName !== name) continue;
        if (phone && rowPhone !== phone) continue;
        if (birthday && rowBirthday !== birthday) continue;
        if (passcode && rowPasscode !== passcode) continue;

        candidates.push({
          rowIndex: i + 2,
          rowValues: row.slice()
        });
      }

      if (!candidates.length) {
        return { status: 'error', message: '一致する会員情報が見つかりませんでした。入力内容をご確認ください。' };
      }
      if (candidates.length > 1) {
        return { status: 'error', message: '一致する会員情報が複数見つかりました。電話番号・生年月日・現在のパスコードをもう1つ追加するか、引き継ぎコードをご利用ください。' };
      }

      matchedRowIndex = candidates[0].rowIndex;
      matchedRowValues = candidates[0].rowValues;
    }

    if (matchedRowIndex < 2 || !matchedRowValues) {
      return { status: 'error', message: '一致する会員情報が見つかりませんでした。入力内容をご確認ください。' };
    }

    const updatedRow = matchedRowValues.slice();
    updatedRow[USER_COL.PASSCODE - 1] = newPasscode;
    updatedRow[USER_COL.DELETE_STATUS - 1] = '';
    updatedRow[USER_COL.DELETED_AT - 1] = '';
    updatedRow[USER_COL.MERGED_INTO - 1] = '';
    clearTransferCodeFromRow_(updatedRow);
    setUserRegistrationSource_(
      updatedRow,
      usedTransferCode ? '引き継ぎコード利用' : '復元',
      usedTransferCode ? '引き継ぎコードで復元' : '本人情報で復元',
      formatDateTime_(new Date())
    );
    sheet.getRange(matchedRowIndex, 1, 1, USER_HEADERS.length).setValues([updatedRow]);

    const matchedUser = buildRecoverAccountUserFromRow_(updatedRow);
    matchedUser.passcode = newPasscode;
    return {
      status: 'ok',
      user: matchedUser,
      recoveredBy: usedTransferCode ? 'transferCode' : 'identity'
    };
  } catch (err) {
    Logger.log('handleRecoverAccount error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function handleIssueTransferCode(data) {
  try {
    const memberId = String(data.memberId || '').trim();
    if (!memberId) return { status: 'error', message: '会員IDが指定されていません。' };

    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    const rowIndex = findUserRowByMemberId_(sheet, memberId);
    if (rowIndex < 2) return { status: 'error', message: '会員情報が見つかりませんでした。' };

    const lastRow = sheet.getLastRow();
    const values = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, USER_HEADERS.length).getValues() : [];
    const range = sheet.getRange(rowIndex, 1, 1, USER_HEADERS.length);
    const row = range.getValues()[0];
    const transferCode = generateUniqueTransferCode_(values);
    const issuedAt = new Date();
    const expiresAt = new Date(issuedAt.getTime());
    expiresAt.setHours(expiresAt.getHours() + TRANSFER_CODE_TTL_HOURS);
    const updatedRow = row.slice();
    updatedRow[USER_COL.TRANSFER_CODE - 1] = transferCode;
    updatedRow[USER_COL.TRANSFER_CODE_ISSUED_AT - 1] = formatDateTime_(issuedAt);
    range.setValues([updatedRow]);

    return {
      status: 'ok',
      transferCode: transferCode,
      issuedAt: formatDateTime_(issuedAt),
      issuedAtLabel: formatTransferCodeDateTime_(issuedAt),
      expiresAt: formatDateTime_(expiresAt),
      expiresAtLabel: formatTransferCodeDateTime_(expiresAt),
      ttlHours: TRANSFER_CODE_TTL_HOURS
    };
  } catch (err) {
    Logger.log('handleIssueTransferCode error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function handleResetForgottenPasscode(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'error', message: '会員が見つかりません' };

    const memberId = String(data.memberId || '').trim();
    const name = normalizeNameForMatch_(data.name);
    const phone = normalizePhoneForMatch_(data.phone);
    const birthday = normalizeDateOnlyValue_(data.birthday);
    const newPasscode = String(data.newPasscode || '').trim();

    if (!name || !phone || !birthday || !newPasscode) {
      return { status: 'error', message: 'お名前・電話番号・生年月日・新しいパスコードを入力してください。' };
    }
    if (!/^(?:\d{4}|\d{6})$/.test(newPasscode)) {
      return { status: 'error', message: '新しいパスコードは4桁または6桁の数字で入力してください。' };
    }

    const values = sheet.getRange(2, 1, lastRow - 1, USER_HEADERS.length).getValues();
    let matchedRow = -1;
    let matchedUser = null;

    if (memberId) {
      const targetRow = findUserRowByMemberId_(sheet, memberId);
      if (targetRow >= 2) {
        const row = values[targetRow - 2];
        const rowPhone = normalizePhoneForMatch_(row[USER_COL.PHONE - 1]);
        const rowName = normalizeNameForMatch_(row[USER_COL.NAME - 1]);
        const rowBirthday = normalizeDateOnlyValue_(row[USER_COL.BIRTHDAY - 1]);
        if (rowPhone === phone && rowName === name && rowBirthday === birthday) {
          matchedRow = targetRow;
          matchedUser = buildRecoverAccountUserFromRow_(row);
        }
      }
    }

    if (!matchedUser) {
      const candidates = [];
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const rowPhone = normalizePhoneForMatch_(row[USER_COL.PHONE - 1]);
        const rowName = normalizeNameForMatch_(row[USER_COL.NAME - 1]);
        const rowBirthday = normalizeDateOnlyValue_(row[USER_COL.BIRTHDAY - 1]);
        if (rowPhone !== phone) continue;
        if (rowName !== name) continue;
        if (rowBirthday !== birthday) continue;
        candidates.push({
          rowIndex: i + 2,
          user: buildRecoverAccountUserFromRow_(row)
        });
      }

      if (candidates.length > 1) {
        return { status: 'error', message: '一致する会員情報が複数見つかりました。院へお問い合わせください。' };
      }
      if (!candidates.length) {
        return { status: 'error', message: '一致する会員情報が見つかりませんでした。入力内容をご確認ください。' };
      }

      matchedRow = candidates[0].rowIndex;
      matchedUser = candidates[0].user;
    }

    if (matchedRow < 2 || !matchedUser) {
      return { status: 'error', message: '本人確認に失敗しました。入力内容をご確認ください。' };
    }

    const range = sheet.getRange(matchedRow, 1, 1, USER_HEADERS.length);
    const updatedRow = range.getValues()[0];
    updatedRow[USER_COL.PASSCODE - 1] = newPasscode;
    updatedRow[USER_COL.DELETE_STATUS - 1] = '';
    updatedRow[USER_COL.DELETED_AT - 1] = '';
    updatedRow[USER_COL.MERGED_INTO - 1] = '';
    clearTransferCodeFromRow_(updatedRow);
    range.setValues([updatedRow]);
    matchedUser = buildRecoverAccountUserFromRow_(updatedRow);
    matchedUser.passcode = newPasscode;
    return { status: 'ok', user: matchedUser };
  } catch (err) {
    Logger.log('handleResetForgottenPasscode error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}


function handleSyncUserRewardStatus(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    const memberId = String(data.memberId || '').trim();
    if (!memberId) return { status: 'error', message: '会員IDが指定されていません' };

    const ensuredUser = ensureUserRowFromActivity_(sheet, {
      memberId: memberId,
      name: data.name,
      phone: data.phone,
      birthday: data.birthday,
      address: data.address,
      passcode: data.passcode
    });
    const targetRow = ensuredUser.rowIndex;
    if (targetRow < 2) {
      return { status: 'error', message: '会員情報の復元に失敗しました' };
    }
    const range = sheet.getRange(targetRow, 1, 1, USER_HEADERS.length);
    const currentRow = ensuredUser.rowData || range.getValues()[0];
    const currentRewardStatus = getRewardStatusFromRow_(currentRow);
    const rewardStatus = sanitizeRewardStatus_({
      stampCount: data.stampCount !== undefined ? data.stampCount : currentRewardStatus.stampCount,
      stampCardNum: data.stampCardNum !== undefined ? data.stampCardNum : currentRewardStatus.stampCardNum,
      rewards: data.rewards !== undefined ? data.rewards : currentRewardStatus.rewards,
      stampHistory: data.stampHistory !== undefined ? data.stampHistory : currentRewardStatus.stampHistory,
      lastStampDate: data.lastStampDate !== undefined ? data.lastStampDate : currentRewardStatus.lastStampDate,
      lastStampAt: data.lastStampAt !== undefined ? data.lastStampAt : currentRewardStatus.lastStampAt,
      stampAchievedDate: data.stampAchievedDate !== undefined ? data.stampAchievedDate : currentRewardStatus.stampAchievedDate
    });
    range.setValues([applyRewardStatusToRow_(currentRow, rewardStatus)]);

    return {
      status: 'ok',
      rewardStatus: rewardStatus
    };
  } catch (err) {
    Logger.log('handleSyncUserRewardStatus error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function handleDrawRewardGacha(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const memberId = String(data && data.memberId || '').trim();
    if (!memberId) {
      return { status: 'error', message: '会員IDが指定されていません' };
    }

    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    const rowIdx = findUserRowByMemberId_(sheet, memberId);
    if (rowIdx === -1) {
      return { status: 'error', message: '会員情報が見つかりません' };
    }

    const range = sheet.getRange(rowIdx, 1, 1, USER_HEADERS.length);
    const currentRow = range.getValues()[0];
    const rewardStatus = getRewardStatusFromRow_(currentRow);
    const currentCardNum = Math.max(1, Number(rewardStatus.stampCardNum) || 1);
    const existingReward = (rewardStatus.rewards || []).find(function (reward) {
      return Math.max(1, Number(reward && reward.cardNum) || 1) === currentCardNum;
    });

    if (existingReward) {
      return {
        status: 'ok',
        alreadyDrawn: true,
        rewardStatus: rewardStatus,
        drawnReward: buildRewardGachaResult_(existingReward, true)
      };
    }

    if (Math.max(0, Number(rewardStatus.stampCount) || 0) < 10) {
      return { status: 'error', message: 'スタンプが10個たまっていません' };
    }

    const selectedPrize = pickWeightedRewardGachaPrize_();
    const rewardEntry = buildRewardGachaEntry_(rewardStatus, selectedPrize);
    const nextRewardStatus = sanitizeRewardStatus_({
      stampCount: rewardStatus.stampCount,
      stampCardNum: rewardStatus.stampCardNum,
      rewards: [rewardEntry].concat(rewardStatus.rewards || []),
      stampHistory: rewardStatus.stampHistory,
      lastStampDate: rewardStatus.lastStampDate,
      stampAchievedDate: rewardStatus.stampAchievedDate || rewardEntry.earnedDate
    });

    range.setValues([applyRewardStatusToRow_(currentRow, nextRewardStatus)]);

    return {
      status: 'ok',
      rewardStatus: nextRewardStatus,
      drawnReward: buildRewardGachaResult_(rewardEntry, false)
    };
  } catch (err) {
    Logger.log('handleDrawRewardGacha error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseErr) { }
  }
}

function handleUnsubscribePush(data) {
  try {
    const warnings = [];
    const memberId = String(data.memberId || '').trim();
    const playerId = String(data.playerId || '').trim();

    if (playerId && CONFIG.ONESIGNAL_APP_ID && CONFIG.ONESIGNAL_REST_API_KEY) {
      try {
        const response = UrlFetchApp.fetch(
          'https://onesignal.com/api/v1/players/' + encodeURIComponent(playerId) + '?app_id=' + encodeURIComponent(CONFIG.ONESIGNAL_APP_ID),
          {
            method: 'delete',
            headers: {
              Authorization: 'Basic ' + CONFIG.ONESIGNAL_REST_API_KEY
            },
            muteHttpExceptions: true
          }
        );
        const code = response.getResponseCode();
        if (code >= 300) {
          warnings.push('OneSignal unsubscribe failed: ' + response.getContentText());
        }
      } catch (oneSignalErr) {
        Logger.log('handleUnsubscribePush OneSignal error: ' + oneSignalErr.toString());
        warnings.push('OneSignal unsubscribe failed: ' + oneSignalErr.toString());
      }
    }

    if (memberId) {
      clearUserPushSubscription_(memberId);
    }

    return {
      status: 'ok',
      warning: warnings.length ? warnings.join(' / ') : ''
    };
  } catch (err) {
    Logger.log('handleUnsubscribePush error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function handleSyncUserDeviceSession(data) {
  try {
    const memberId = String(data && data.memberId || '').trim();
    const deviceId = String(data && data.deviceId || '').trim();
    if (!memberId || !deviceId) {
      return { status: 'error', message: '会員IDまたは端末IDが不足しています' };
    }

    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    const rowIdx = findUserRowByMemberId_(sheet, memberId);
    if (rowIdx < 2) return { status: 'error', message: '会員情報が見つかりません' };

    const sessions = upsertUserDeviceSession_(sheet, rowIdx, data);
    return { status: 'ok', devices: sessions };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getUserDevices(params) {
  try {
    const memberId = String(params && params.memberId || '').trim();
    if (!memberId) return { status: 'error', message: '会員IDが必要です' };
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    const rowIdx = findUserRowByMemberId_(sheet, memberId);
    if (rowIdx < 2) return { status: 'ok', devices: [] };
    const row = sheet.getRange(rowIdx, 1, 1, USER_HEADERS.length).getValues()[0];
    return { status: 'ok', devices: getUserDeviceSessionsFromRow_(row) };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleRemoveUserDeviceSession(data) {
  try {
    const memberId = String(data && data.memberId || '').trim();
    const deviceId = String(data && data.deviceId || '').trim();
    if (!memberId || !deviceId) {
      return { status: 'error', message: '会員IDまたは端末IDが不足しています' };
    }
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    const rowIdx = findUserRowByMemberId_(sheet, memberId);
    if (rowIdx < 2) return { status: 'error', message: '会員情報が見つかりません' };
    const range = sheet.getRange(rowIdx, 1, 1, USER_HEADERS.length);
    const row = range.getValues()[0];
    const sessions = getUserDeviceSessionsFromRow_(row).filter(function (session) {
      return session.deviceId !== deviceId;
    });
    range.setValues([applyUserDeviceSessionsToRow_(row, sessions)]);
    return { status: 'ok', devices: sessions };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getUserRewardStatus(params) {
  try {
    if (!params || !params.memberId) {
      return { status: 'error', message: '会員IDが必要です' };
    }
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    const rowIdx = findUserRowByMemberId_(sheet, params.memberId);
    if (rowIdx === -1) {
      return { status: 'ok', rewardStatus: getDefaultRewardStatus_() };
    }
    const row = sheet.getRange(rowIdx, 1, 1, USER_HEADERS.length).getValues()[0];
    return { status: 'ok', rewardStatus: getRewardStatusFromRow_(row) };
  } catch (err) {
    Logger.log('getUserRewardStatus error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function findUserRowForAdminMutation_(sheet, data) {
  const memberId = String(data && data.memberId || '').trim();
  if (memberId) {
    const found = findUserRowByMemberId_(sheet, memberId);
    if (found !== -1) return found;
  }

  const rowIdx = Number(data && data.rowIdx || 0);
  if (sheet && rowIdx > 1 && rowIdx <= sheet.getLastRow()) {
    return rowIdx;
  }

  return -1;
}

function handleUpdateAdminUser(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    const rowIdx = findUserRowForAdminMutation_(sheet, data);
    if (!(sheet && rowIdx > 1)) {
      return { status: 'error', message: '会員が見つかりません' };
    }

    const range = sheet.getRange(rowIdx, 1, 1, USER_HEADERS.length);
    const currentRow = range.getValues()[0];
    const updatedRow = currentRow.slice();
    updatedRow[USER_COL.NAME - 1] = data.name !== undefined ? data.name : currentRow[USER_COL.NAME - 1];
    updatedRow[USER_COL.KANA - 1] = data.kana !== undefined ? data.kana : currentRow[USER_COL.KANA - 1];
    updatedRow[USER_COL.PHONE - 1] = data.phone !== undefined ? data.phone : currentRow[USER_COL.PHONE - 1];
    updatedRow[USER_COL.MEMO - 1] = data.memo !== undefined ? data.memo : currentRow[USER_COL.MEMO - 1];
    updatedRow[USER_COL.BIRTHDAY - 1] = data.birthday !== undefined ? data.birthday : currentRow[USER_COL.BIRTHDAY - 1];
    updatedRow[USER_COL.ADDRESS - 1] = data.address !== undefined ? data.address : currentRow[USER_COL.ADDRESS - 1];
    range.setValues([updatedRow]);
    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleUpdateAdminUser error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function handleDeleteUser(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    const rowIdx = findUserRowForAdminMutation_(sheet, data);
    if (!(sheet && rowIdx > 1)) {
      return { status: 'error', message: '会員が見つかりません' };
    }
    const row = sheet.getRange(rowIdx, 1, 1, USER_HEADERS.length).getValues()[0];
    row[USER_COL.DELETE_STATUS - 1] = SOFT_DELETE_STATUS;
    row[USER_COL.DELETED_AT - 1] = formatDateTime_(new Date());
    sheet.getRange(rowIdx, 1, 1, USER_HEADERS.length).setValues([row]);
    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleDeleteUser error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function handleUpdateAdminRewardStatus(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    const rowIdx = Number(data.rowIdx);
    if (!(sheet && rowIdx > 1)) {
      return { status: 'error', message: '会員が見つかりません' };
    }

    const range = sheet.getRange(rowIdx, 1, 1, USER_HEADERS.length);
    const currentRow = range.getValues()[0];
    const currentStatus = getRewardStatusFromRow_(currentRow);
    const nextStatus = sanitizeRewardStatus_({
      stampCount: data.stampCount !== undefined ? data.stampCount : currentStatus.stampCount,
      stampCardNum: data.stampCardNum !== undefined ? data.stampCardNum : currentStatus.stampCardNum,
      rewards: data.rewards !== undefined ? data.rewards : currentStatus.rewards,
      stampHistory: data.stampHistory !== undefined ? data.stampHistory : currentStatus.stampHistory,
      lastStampDate: data.lastStampDate !== undefined ? data.lastStampDate : currentStatus.lastStampDate,
      lastStampAt: data.lastStampAt !== undefined ? data.lastStampAt : currentStatus.lastStampAt,
      stampAchievedDate: data.stampAchievedDate !== undefined ? data.stampAchievedDate : currentStatus.stampAchievedDate
    });
    
    // デモ用：取得制限（日付）の強制クリア
    if (data.clearLastStampDate === true) {
      nextStatus.lastStampDate = '';
      nextStatus.lastStampAt = '';
    }
    
    range.setValues([applyRewardStatusToRow_(currentRow, nextStatus)]);
    return { status: 'ok', rewardStatus: nextStatus };
  } catch (err) {
    Logger.log('handleUpdateAdminRewardStatus error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function buildOrderStatsByMemberId_() {
  const stats = {};
  const ss = getOrCreateSpreadsheet();
  const sheet = ensureOrdersSheetStructure_(ss.getSheetByName(SHEETS.ORDERS));
  if (!sheet || sheet.getLastRow() < 2) return stats;
  const deleteCols = ensureSoftDeleteColumns_(sheet);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const orderSeen = {};
  rows.forEach(function (row) {
    if (isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol)) return;
    const memberId = String(row[13] || '').trim();
    const orderId = String(row[0] || '').trim();
    if (!memberId || !orderId) return;
    if (!stats[memberId]) {
      stats[memberId] = {
        orderCount: 0,
        pendingOrderCount: 0,
        lastOrderAt: '',
        orderTotal: 0
      };
    }
    const key = memberId + '::' + orderId;
    if (!orderSeen[key]) {
      orderSeen[key] = true;
      stats[memberId].orderCount += 1;
      if (normalizeOrderStatus_(row[11] || '') === '受付中') {
        stats[memberId].pendingOrderCount += 1;
      }
      const currentTs = parseLooseDateToTimestamp_(row[1]);
      const lastTs = parseLooseDateToTimestamp_(stats[memberId].lastOrderAt);
      if (currentTs >= lastTs) {
        stats[memberId].lastOrderAt = formatMaybeDateTime_(row[1]);
      }
      const orderTotal = Number(String(row[9] || '').replace(/[^0-9.-]+/g, '')) || 0;
      stats[memberId].orderTotal += orderTotal;
    }
  });
  return stats;
}

function getAdminUsers() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    const orderStats = buildOrderStatsByMemberId_();

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', users: [] };

    const data = sheet.getRange(2, 1, lastRow - 1, USER_HEADERS.length).getValues();
    const users = data.map((row, index) => {
      if (String(row[USER_COL.DELETE_STATUS - 1] || '').trim() === SOFT_DELETE_STATUS) return null;
      const rewardStatus = getRewardStatusFromRow_(row);
      const registrationSource = buildUserRegistrationSourceFromRow_(row);
      const memberId = String(row[USER_COL.MEMBER_ID - 1] || '');
      const userOrderStats = orderStats[memberId] || { orderCount: 0, pendingOrderCount: 0, lastOrderAt: '', orderTotal: 0 };
      const latestStampActivityAt = rewardStatus.lastStampAt
        || (rewardStatus.stampHistory && rewardStatus.stampHistory[0] && rewardStatus.stampHistory[0].acquiredDate)
        || rewardStatus.lastStampDate
        || rewardStatus.stampAchievedDate
        || '';
      return {
        rowIdx: index + 2,
        memberId: memberId,
        timestamp: (row[USER_COL.TIMESTAMP - 1] instanceof Date) ? Utilities.formatDate(row[USER_COL.TIMESTAMP - 1], 'Asia/Tokyo', 'yyyy/M/d H:mm') : String(row[USER_COL.TIMESTAMP - 1] || ''),
        name: String(row[USER_COL.NAME - 1] || ''),
        kana: String(row[USER_COL.KANA - 1] || ''),
        phone: String(row[USER_COL.PHONE - 1] || ''),
        avatarUrl: String(row[USER_COL.AVATAR_URL - 1] || ''),
        memo: String(row[USER_COL.MEMO - 1] || ''),
        pushEnabled: !!row[USER_COL.PUSH - 1],
        status: String(row[USER_COL.STATUS - 1] || ''),
        birthday: (row[USER_COL.BIRTHDAY - 1] instanceof Date) ? Utilities.formatDate(row[USER_COL.BIRTHDAY - 1], 'Asia/Tokyo', 'yyyy/M/d') : String(row[USER_COL.BIRTHDAY - 1] || '').replace(/-/g, '/'),
        address: String(row[USER_COL.ADDRESS - 1] || ''),
        deviceSessions: getUserDeviceSessionsFromRow_(row),
        deviceCount: getUserDeviceSessionsFromRow_(row).length,
        stampCount: rewardStatus.stampCount,
        stampCardNum: rewardStatus.stampCardNum,
        rewards: rewardStatus.rewards,
        stampHistory: rewardStatus.stampHistory,
        lastStampDate: rewardStatus.lastStampDate,
        lastStampAt: rewardStatus.lastStampAt,
        stampAchievedDate: rewardStatus.stampAchievedDate,
        latestStampActivityAt: latestStampActivityAt,
        registrationSource: registrationSource.source,
        registrationSourceDetail: registrationSource.detail,
        registrationSourceUpdatedAt: registrationSource.updatedAt,
        orderCount: userOrderStats.orderCount,
        pendingOrderCount: userOrderStats.pendingOrderCount,
        lastOrderAt: userOrderStats.lastOrderAt,
        orderTotal: userOrderStats.orderTotal
      };
    }).filter(Boolean).reverse(); // 最新を上に

    return { status: 'ok', users: users };
  } catch (err) {
    Logger.log('getAdminUsers error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function mergeRewardEntries_(leftRewards, rightRewards) {
  const map = {};
  (Array.isArray(leftRewards) ? leftRewards : []).concat(Array.isArray(rightRewards) ? rightRewards : []).map(function (reward, index) {
    return sanitizeRewardEntry_(reward, 'merged-reward-' + index);
  }).forEach(function (reward) {
    const key = [reward.cardNum, reward.rewardName, reward.earnedDate].join('|');
    if (!map[key]) {
      map[key] = reward;
      return;
    }
    if (reward.used && !map[key].used) {
      map[key] = reward;
    }
  });
  return Object.keys(map).map(function (key) { return map[key]; });
}

function mergeStampHistoryEntries_(leftHistory, rightHistory) {
  return sanitizeStampHistoryList_([].concat(leftHistory || [], rightHistory || []));
}

function buildDuplicateUsersFromRows_(rows) {
  const groups = [];
  const seen = {};
  const pushGroup = function (reason, list) {
    const filtered = list.filter(function (item) { return item; });
    if (filtered.length < 2) return;
    const key = filtered.map(function (item) { return item.memberId; }).sort().join('|');
    if (seen[key]) {
      seen[key].reasons[reason] = true;
      return;
    }
    seen[key] = {
      key: key,
      reasons: {}
    };
    seen[key].reasons[reason] = true;
    groups.push({
      key: key,
      reason: reason,
      reasons: [reason],
      users: filtered.sort(function (a, b) {
        return parseLooseDateToTimestamp_(b.updatedAt) - parseLooseDateToTimestamp_(a.updatedAt);
      })
    });
  };

  const phoneMap = {};
  const nameBirthdayMap = {};
  const nameMap = {};
  rows.forEach(function (item) {
    const phone = normalizePhoneForMatch_(item.phone);
    const name = normalizeNameForMatch_(item.name);
    const birthday = normalizeDateOnlyValue_(item.birthday);
    if (phone) {
      phoneMap[phone] = phoneMap[phone] || [];
      phoneMap[phone].push(item);
    }
    if (name && birthday) {
      const key = name + '|' + birthday;
      nameBirthdayMap[key] = nameBirthdayMap[key] || [];
      nameBirthdayMap[key].push(item);
    }
    if (name) {
      nameMap[name] = nameMap[name] || [];
      nameMap[name].push(item);
    }
  });

  Object.keys(phoneMap).forEach(function (key) {
    if (phoneMap[key].length > 1) pushGroup('電話番号が一致', phoneMap[key]);
  });
  Object.keys(nameBirthdayMap).forEach(function (key) {
    if (nameBirthdayMap[key].length > 1) pushGroup('氏名と生年月日が一致', nameBirthdayMap[key]);
  });
  Object.keys(nameMap).forEach(function (key) {
    if (nameMap[key].length > 1) pushGroup('氏名が一致', nameMap[key]);
  });

  return groups.map(function (group) {
    group.reasons = Object.keys(seen[group.key].reasons);
    return group;
  });
}

function getDuplicateUsers() {
  try {
    const admin = getAdminUsers();
    if (admin.status !== 'ok') return admin;
    const rows = (admin.users || []).map(function (user) {
      return {
        rowIdx: user.rowIdx,
        memberId: user.memberId,
        name: user.name,
        phone: user.phone,
        birthday: user.birthday,
        updatedAt: user.timestamp,
        stampCount: user.stampCount,
        orderCount: user.orderCount,
        pendingOrderCount: user.pendingOrderCount,
        lastOrderAt: user.lastOrderAt,
        orderTotal: user.orderTotal,
        deviceCount: user.deviceCount
      };
    });
    return { status: 'ok', groups: buildDuplicateUsersFromRows_(rows) };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleMergeUsers(data) {
  try {
    const targetMemberId = String(data && data.targetMemberId || '').trim();
    const sourceMemberIds = Array.isArray(data && data.sourceMemberIds)
      ? data.sourceMemberIds.map(function (value) { return String(value || '').trim(); }).filter(Boolean)
      : [];
    if (!targetMemberId || !sourceMemberIds.length) {
      return { status: 'error', message: '統合対象が不足しています' };
    }

    const ss = getOrCreateSpreadsheet();
    const usersSheet = getOrCreateUsersSheet_(ss);
    const targetRowIdx = findUserRowByMemberId_(usersSheet, targetMemberId);
    if (targetRowIdx < 2) return { status: 'error', message: '統合先会員が見つかりません' };

    const targetRange = usersSheet.getRange(targetRowIdx, 1, 1, USER_HEADERS.length);
    const targetRow = targetRange.getValues()[0];
    let mergedRow = targetRow.slice();
    let mergedRewards = getRewardStatusFromRow_(targetRow);
    let mergedSessions = getUserDeviceSessionsFromRow_(targetRow);

    sourceMemberIds.forEach(function (memberId) {
      const rowIdx = findUserRowByMemberId_(usersSheet, memberId);
      if (rowIdx < 2 || rowIdx === targetRowIdx) return;
      const range = usersSheet.getRange(rowIdx, 1, 1, USER_HEADERS.length);
      const row = range.getValues()[0];
      if (String(row[USER_COL.DELETE_STATUS - 1] || '').trim() === SOFT_DELETE_STATUS &&
        String(row[USER_COL.MERGED_INTO - 1] || '').trim()) {
        return;
      }

      [
        USER_COL.NAME,
        USER_COL.KANA,
        USER_COL.PHONE,
        USER_COL.AVATAR_URL,
        USER_COL.MEMO,
        USER_COL.PUSH,
        USER_COL.STATUS,
        USER_COL.BIRTHDAY,
        USER_COL.ADDRESS,
        USER_COL.PASSCODE
      ].forEach(function (col) {
        if (!String(mergedRow[col - 1] || '').trim() && String(row[col - 1] || '').trim()) {
          mergedRow[col - 1] = row[col - 1];
        }
      });

      mergedRewards = sanitizeRewardStatus_({
        stampCount: Math.max(Number(mergedRewards.stampCount || 0), Number(getRewardStatusFromRow_(row).stampCount || 0)),
        stampCardNum: Math.max(Number(mergedRewards.stampCardNum || 1), Number(getRewardStatusFromRow_(row).stampCardNum || 1)),
        rewards: mergeRewardEntries_(mergedRewards.rewards, getRewardStatusFromRow_(row).rewards),
        stampHistory: mergeStampHistoryEntries_(mergedRewards.stampHistory, getRewardStatusFromRow_(row).stampHistory),
        lastStampDate: mergedRewards.lastStampDate || getRewardStatusFromRow_(row).lastStampDate,
        lastStampAt: mergedRewards.lastStampAt || getRewardStatusFromRow_(row).lastStampAt,
        stampAchievedDate: mergedRewards.stampAchievedDate || getRewardStatusFromRow_(row).stampAchievedDate
      });
      mergedSessions = mergeDeviceSessionLists_(mergedSessions, getUserDeviceSessionsFromRow_(row));

      const nextSourceRow = row.slice();
      nextSourceRow[USER_COL.DELETE_STATUS - 1] = SOFT_DELETE_STATUS;
      nextSourceRow[USER_COL.DELETED_AT - 1] = formatDateTime_(new Date());
      nextSourceRow[USER_COL.MERGED_INTO - 1] = targetMemberId;
      range.setValues([nextSourceRow]);
    });

    mergedRow = applyRewardStatusToRow_(mergedRow, mergedRewards);
    mergedRow = applyUserDeviceSessionsToRow_(mergedRow, mergedSessions);
    mergedRow[USER_COL.DELETE_STATUS - 1] = '';
    mergedRow[USER_COL.DELETED_AT - 1] = '';
    mergedRow[USER_COL.MERGED_INTO - 1] = '';
    setUserRegistrationSource_(mergedRow, '重複候補からの復旧', '会員統合で情報を集約', formatDateTime_(new Date()));
    targetRange.setValues([mergedRow]);

    const orderSheet = ss.getSheetByName(SHEETS.ORDERS);
    if (orderSheet && orderSheet.getLastRow() >= 2) {
      const orderData = orderSheet.getRange(2, 1, orderSheet.getLastRow() - 1, 14).getValues();
      orderData.forEach(function (row, index) {
        if (sourceMemberIds.indexOf(String(row[13] || '').trim()) !== -1) {
          orderSheet.getRange(index + 2, 14).setValue(targetMemberId);
        }
      });
    }

    return { status: 'ok' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/**
 * お知らせ（PUSH_NOTICES）を取得
 */
function getPushNotices() {
  try {
    const ensured = ensurePushNoticeSheetStructure_();
    const sheet = ensured.sheet;
    const columns = ensured.columns;
    const deleteCols = ensureSoftDeleteColumns_(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', notices: [] };

    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const notices = data.map(function (row, index) {
      if (isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol)) return null;
      const sentAt = formatMaybeDateTime_(row[columns.sentAt - 1]);
      const scheduledAt = formatMaybeDateTime_(row[columns.scheduledAt - 1]);
      const updatedAt = formatMaybeDateTime_(row[columns.updatedAt - 1]);
      return {
        rowIdx: index + 2,
        date: parseLooseDateToTimestamp_(sentAt || scheduledAt || updatedAt),
        sentAt: sentAt,
        scheduledAt: scheduledAt,
        updatedAt: updatedAt,
        title: String(row[columns.title - 1] || ''),
        body: String(row[columns.body - 1] || ''),
        targetStatus: String(row[columns.targetStatus - 1] || 'all'),
        targetDetail: String(row[columns.targetDetail - 1] || ''),
        recipientCount: Number(row[columns.recipientCount - 1] || 0),
        status: String(row[columns.status - 1] || ''),
        targetPage: normalizeTargetPage_(row[columns.targetPage - 1]),
        previewBody: String(row[columns.previewBody - 1] || ''),
        notificationId: String(row[columns.notificationId - 1] || ''),
        result: String(row[columns.result - 1] || '')
      };
    }).filter(Boolean).sort(function (a, b) {
      return (b.date || 0) - (a.date || 0);
    });

    return { status: 'ok', notices: notices };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
/**
 * Push通知対象のユーザーを取得
 */
function getPushUsers() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateUsersSheet_(ss);
    if (!sheet) return { status: 'ok', users: [] };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', users: [] };

    const orderStats = buildOrderStatsByMemberId_();
    const data = sheet.getRange(2, 1, lastRow - 1, USER_HEADERS.length).getValues();
    const users = [];
    data.forEach(row => {
      if (String(row[USER_COL.DELETE_STATUS - 1] || '').trim() === SOFT_DELETE_STATUS) return;
      if (row[USER_COL.PUSH - 1]) {
        const memberId = String(row[USER_COL.MEMBER_ID - 1] || '').trim();
        const rewardStatus = getRewardStatusFromRow_(row);
        const deviceSessions = getUserDeviceSessionsFromRow_(row);
        const userOrderStats = orderStats[memberId] || { orderCount: 0, pendingOrderCount: 0, lastOrderAt: '', orderTotal: 0 };
        users.push({
          memberId: memberId,
          name: row[USER_COL.NAME - 1],
          phone: row[USER_COL.PHONE - 1],
          birthday: normalizeDateOnlyValue_(row[USER_COL.BIRTHDAY - 1]),
          status: String(row[USER_COL.STATUS - 1] || ''),
          subscription: row[USER_COL.PUSH - 1],
          stampCount: rewardStatus.stampCount,
          rewardCount: (rewardStatus.rewards || []).filter(function (reward) { return !reward.used; }).length,
          orderCount: userOrderStats.orderCount,
          pendingOrderCount: userOrderStats.pendingOrderCount,
          lastOrderAt: userOrderStats.lastOrderAt,
          deviceCount: deviceSessions.length
        });
      }
    });

    return { status: 'ok', users: users };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/**
 * 全ユーザーへ通知を配信
 */
function broadcastPush(data) {
  try {
    const ensured = ensurePushNoticeSheetStructure_();
    const sheet = ensured.sheet;
    const columns = ensured.columns;
    const mode = String(data && data.mode || 'send').trim();
    const targetUsers = Array.isArray(data && data.targetUsers) ? data.targetUsers : [];
    const record = buildPushNoticeRowFromInput_(data, mode === 'draft' ? PUSH_STATUS_DRAFT : mode === 'schedule' ? PUSH_STATUS_SCHEDULED : PUSH_STATUS_SENT);
    const rowIdx = writePushNoticeRow_(sheet, Number(data && data.rowIdx || 0), columns, record);

    if (mode === 'draft') {
      return { status: 'ok', message: '通知を下書き保存しました', rowIdx: rowIdx, recipientCount: record.recipientCount };
    }
    if (mode === 'schedule') {
      if (!record.scheduledAt) {
        return { status: 'error', message: '配信予定日時を入力してください' };
      }
      return { status: 'ok', message: '通知を予約しました', rowIdx: rowIdx, recipientCount: record.recipientCount };
    }
    if (mode === 'test') {
      if (!targetUsers.length) {
        return { status: 'error', message: 'テスト送信先を選択してください' };
      }
      const result = sendPushNoticePayload_(record);
      record.sentAt = formatDateTime_(new Date());
      record.status = PUSH_STATUS_SENT;
      record.notificationId = result.id;
      record.result = 'テスト送信済み';
      writePushNoticeRow_(sheet, rowIdx, columns, record);
      return { status: 'ok', message: 'テスト通知を送信しました', rowIdx: rowIdx, recipientCount: targetUsers.length };
    }

    const result = sendPushNoticePayload_(record);
    record.sentAt = formatDateTime_(new Date());
    record.status = PUSH_STATUS_SENT;
    record.notificationId = result.id;
    record.result = '送信済み';
    writePushNoticeRow_(sheet, rowIdx, columns, record);
    return { status: 'ok', message: '通知を送信しました', rowIdx: rowIdx, recipientCount: record.recipientCount };
  } catch (err) {
    Logger.log("broadcastPush error: " + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function getCalendarEvents() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.CALENDAR);
    if (!sheet) return { status: 'ok', events: [] };
    ensureUpdatedAtColumn_(sheet, '更新日時');
    ensureSortOrderColumn_(sheet);
    const noticeCol = ensureNoticeVisibilityColumn_(sheet, 5, '公開');
    const categoryCol = ensureNamedColumn_(sheet, 'カテゴリ', 140);
    const deleteCols = ensureSoftDeleteColumns_(sheet);
    const publishAtCol = ensurePublishAtColumn_(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', events: [] };

    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const sortCol = header.indexOf('表示順') + 1;
    const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    const events = values.map(function (row, index) {
      if (row[4] !== '公開' ||
        isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol) ||
        !isPublishAtAvailable_(row[publishAtCol - 1])) {
        return null;
      }
      return {
        rowIdx: index + 2,
        date: formatMaybeDateTime_(row[0]),
        title: row[1],
        desc: row[2],
        color: row[3],
        category: String(row[categoryCol - 1] || ''),
        image: String(row[5] || ''),
        updatedAt: formatMaybeDateTime_(row[6]),
        publishAt: formatMaybeDateTime_(row[publishAtCol - 1]),
        noticeStatus: normalizePublishVisibilityStatus_(row[noticeCol - 1] || row[4] || '公開'),
        sortOrder: sortCol > 0 ? Number(row[sortCol - 1] || 0) : 0
      };
    }).filter(Boolean);

    return { status: 'ok', events: events };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getAdminCalendar() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.CALENDAR);
    if (!sheet) return { status: 'ok', events: [] };
    ensureUpdatedAtColumn_(sheet, '更新日時');
    ensureSortOrderColumn_(sheet);
    const noticeCol = ensureNoticeVisibilityColumn_(sheet, 5, '公開');
    const categoryCol = ensureNamedColumn_(sheet, 'カテゴリ', 140);
    const deleteCols = ensureSoftDeleteColumns_(sheet);
    const publishAtCol = ensurePublishAtColumn_(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', events: [] };

    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const sortCol = header.indexOf('表示順') + 1;
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    const events = data.map(function (row, index) {
      if (isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol)) return null;
      return {
        rowIdx: index + 2,
        date: formatMaybeDateTime_(row[0]),
        title: row[1],
        desc: row[2],
        color: row[3],
        category: String(row[categoryCol - 1] || ''),
        publishStatus: row[4],
        image: String(row[5] || ''),
        updatedAt: formatMaybeDateTime_(row[6]),
        publishAt: formatMaybeDateTime_(row[publishAtCol - 1]),
        noticeStatus: normalizePublishVisibilityStatus_(row[noticeCol - 1] || row[4] || '公開'),
        sortOrder: sortCol > 0 ? Number(row[sortCol - 1] || 0) : 0
      };
    }).filter(Boolean);

    return { status: 'ok', events: events };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleUpdateItemOrder(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const updatesBySheet = {};
    (data.updates || []).forEach(u => {
      if (!updatesBySheet[u.sheet]) updatesBySheet[u.sheet] = [];
      updatesBySheet[u.sheet].push(u);
    });

    for (const sheetName in updatesBySheet) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      const sortCol = ensureSortOrderColumn_(sheet);
      const updates = updatesBySheet[sheetName];
      updates.forEach(u => {
        sheet.getRange(u.rowIdx, sortCol).setValue(u.sortOrder);
      });
    }

    // キャッシュクリア
    bumpDataCacheVersion_();
    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleUpdateItemOrder error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

/**
 * 自動プッシュ通知送信（ブログ・カレンダー追加時に自動呼び出し）
 */
function sendAutoPush(title, body, options) {
  try {
    if (!CONFIG.ONESIGNAL_APP_ID || !CONFIG.ONESIGNAL_REST_API_KEY) return;
    const record = buildPushNoticeRowFromInput_({
      title: title,
      body: body,
      targetStatus: options && options.targetStatus ? options.targetStatus : 'all',
      targetPage: options && options.targetPage ? options.targetPage : 'home',
      previewBody: options && options.previewBody ? options.previewBody : body
    }, PUSH_STATUS_AUTO_SENT);
    const ensured = ensurePushNoticeSheetStructure_();
    const rowIdx = writePushNoticeRow_(ensured.sheet, 0, ensured.columns, record);
    const result = sendPushNoticePayload_(record);
    record.sentAt = formatDateTime_(new Date());
    record.notificationId = result.id;
    record.result = '自動送信済み';
    writePushNoticeRow_(ensured.sheet, rowIdx, ensured.columns, record);
  } catch (err) {
    Logger.log('sendAutoPush error: ' + err.toString());
  }
}

/**
 * 初期化に必要なデータを一括取得
 */
function getInitialData() {
  try {
    return {
      status: 'ok',
      news: getBlogNews(),
      products: getProducts(),
      calendar: getCalendarEvents(),
      pushNotices: getPushNotices(),
      categories: getCategories()
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
/**
 * アンケート回答を新しいシートに展開する
 */
function handleExportSurvey(data) {
  try {
    const surveyId = data.surveyId;
    const ss = getOrCreateSpreadsheet();

    // 1. アンケート情報の取得
    const masterSheet = ss.getSheetByName(SHEETS.SURVEY_MASTER);
    let surveyTitle = "アンケート";
    let questions = [];
    if (masterSheet && masterSheet.getLastRow() >= 2) {
      const masterValues = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 4).getValues();
      const surveyRow = masterValues.find(row => String(row[0]).split('.')[0] === String(surveyId).split('.')[0]);
      if (surveyRow) {
        surveyTitle = surveyRow[1];
        try { questions = JSON.parse(surveyRow[3]); } catch (e) { }
      }
    }

    // 2. 回答データの取得
    const responseResult = getSurveyResponses({ surveyId: surveyId });
    if (responseResult.status !== 'ok') return responseResult;
    const responses = responseResult.responses;

    if (responses.length === 0) {
      return { status: 'error', message: '回答がまだありません' };
    }

    // 3. 書き出し用シートの準備
    // シート名に使用できない記号 \ / [ ] ? * : をアンダースコアに置換し、長さを制限する
    const sanitizedTitle = surveyTitle.replace(/[\\\/\[\]\?\*:]/g, '_').substring(0, 50);
    const exportSheetName = `回答集計_${sanitizedTitle}`;
    let exportSheet = ss.getSheetByName(exportSheetName);
    if (exportSheet) {
      exportSheet.clear();
    } else {
      exportSheet = ss.insertSheet(exportSheetName);
    }

    // 4. ヘッダーの作成
    const headers = ['回答日時'];
    questions.forEach(q => {
      headers.push(q.title);
    });
    exportSheet.appendRow(headers);
    exportSheet.getRange(1, 1, 1, headers.length).setBackground('#d4e8c8').setFontWeight('bold');

    // 5. データの流し込み
    const rows = responses.map(r => {
      const row = [r.submittedAt];
      questions.forEach(q => {
        let ansRaw = r.answers[q.id];
        // 回答がオブジェクト { title, value } の場合は value を取り出す
        let ans = (ansRaw && typeof ansRaw === 'object' && !Array.isArray(ansRaw)) ? ansRaw.value : ansRaw;
        
        if (Array.isArray(ans)) ans = ans.join(', ');
        row.push(ans || '');
      });
      return row;
    });

    if (rows.length > 0) {
      exportSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    exportSheet.activate();

    return {
      status: 'ok',
      message: `スプレッドシートに「${exportSheetName}」シートを作成しました。`,
      sheetUrl: ss.getUrl() + "#gid=" + exportSheet.getSheetId()
    };
  } catch (err) {
    Logger.log('handleExportSurvey error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

// ========== カテゴリ管理 ==========

/**
 * カテゴリ一覧を取得（なければ初期作成）
 */
function getCategorySheet_() {
  const ss = getOrCreateSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.CATEGORIES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.CATEGORIES);
    sheet.appendRow(['カテゴリ名', '分類']);
  }

  if (sheet.getLastRow() < 2) {
    const categoryDefaults = [
      ['お知らせ', 'お知らせ'],
      ['休診情報', 'お知らせ'],
      ['ブログ', 'ブログ'],
      ['イベント', 'ブログ'],
      ['商品情報', 'ブログ'],
    ];
    sheet.getRange(2, 1, categoryDefaults.length, 2).setValues(categoryDefaults);
  }

  return sheet;
}

function getCategories() {
  try {
    const sheet = getCategorySheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', categories: [] };

    const values = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
    const seen = {};
    const categories = values.map(function (row) {
      const name = String(row[0] || '').trim();
      const rawType = String(row[1] || '').trim();
      // デフォルトはブログ。有効な種別（メニュー、通知、お知らせ、ブログ）ならそのまま返す
      let type = rawType || 'ブログ';
      if (['お知らせ', 'ブログ', 'メニュー', '通知'].indexOf(type) === -1) {
        type = 'ブログ';
      }
      return { name: name, type: type };
    }).filter(function (item) {
      if (!item.name || seen[item.name]) return false;
      seen[item.name] = true;
      return true;
    });

    return { status: 'ok', categories: categories };
  } catch (err) {
    Logger.log('getCategories error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function handleAddCategory(data) {
  try {
    const name = String(data.name || '').trim();
    if (!name) throw new Error('カテゴリ名を入力してください');

    const validTypes = ['お知らせ', 'ブログ', 'メニュー', '通知'];
    const type = validTypes.indexOf(data.categoryType) !== -1 ? data.categoryType : 'ブログ';
    const sheet = getCategorySheet_();
    const lastRow = sheet.getLastRow();
    const existing = lastRow >= 2
      ? sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().some(function (row) {
        return String(row[0] || '').trim() === name;
      })
      : false;

    if (existing) throw new Error('同じカテゴリ名が既に登録されています');

    sheet.appendRow([name, type]);
    return { status: 'ok', message: 'カテゴリを追加しました' };
  } catch (err) {
    Logger.log('handleAddCategory error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function handleUpdateCategory(data) {
  try {
    const oldName = String(data.oldName || '').trim();
    const newName = String(data.newName || '').trim();
    if (!oldName) throw new Error('更新対象のカテゴリが見つかりません');
    if (!newName) throw new Error('カテゴリ名を入力してください');

    const validTypes = ['お知らせ', 'ブログ', 'メニュー', '通知'];
    const type = validTypes.indexOf(String(data.categoryType || '').trim()) !== -1
      ? String(data.categoryType || '').trim()
      : 'ブログ';
    const sheet = getCategorySheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('カテゴリデータがありません');

    const values = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
    let targetRowIdx = 0;

    values.forEach(function (row, index) {
      const currentName = String(row[0] || '').trim();
      const rowIdx = index + 2;
      if (currentName === oldName) targetRowIdx = rowIdx;
      if (currentName === newName && currentName !== oldName) {
        throw new Error('同じカテゴリ名が既に登録されています');
      }
    });

    if (!targetRowIdx) throw new Error('更新対象のカテゴリが見つかりません');

    sheet.getRange(targetRowIdx, 1, 1, 2).setValues([[newName, type]]);
    return { status: 'ok', message: 'カテゴリを更新しました' };
  } catch (err) {
    Logger.log('handleUpdateCategory error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function getBlogNews() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.BLOG);
    if (!sheet) return { status: 'ok', news: [] };
    ensureUpdatedAtColumn_(sheet, '更新日時');
    ensureSortOrderColumn_(sheet);
    const noticeCol = ensureNoticeVisibilityColumn_(sheet, 6, '公開');
    const deleteCols = ensureSoftDeleteColumns_(sheet);
    const publishAtCol = ensurePublishAtColumn_(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', news: [] };

    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const sortCol = header.indexOf('表示順') + 1;
    const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    const news = values.map(function (row, index) {
      if (row[5] !== '公開' ||
        isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol) ||
        !isPublishAtAvailable_(row[publishAtCol - 1])) {
        return null;
      }
      return {
        rowIdx: index + 2,
        date: formatMaybeDateTime_(row[0]),
        title: row[1],
        category: row[2],
        icon: row[3],
        body: row[4],
        image: row[7] || '', // H列に画像URLがある想定
        updatedAt: formatMaybeDateTime_(row[6]),
        publishAt: formatMaybeDateTime_(row[publishAtCol - 1]),
        noticeStatus: normalizePublishVisibilityStatus_(row[noticeCol - 1] || row[5] || '公開'),
        sortOrder: sortCol > 0 ? Number(row[sortCol - 1] || 0) : 0
      };
    }).filter(Boolean);

    const categoriesRes = getCategories();
    return {
      status: 'ok',
      news: news,
      categories: categoriesRes && categoriesRes.status === 'ok' ? categoriesRes.categories : []
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getAdminBlogs() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.BLOG);
    if (!sheet) return { status: 'ok', blogs: [] };
    ensureUpdatedAtColumn_(sheet, '更新日時');
    ensureSortOrderColumn_(sheet);
    const noticeCol = ensureNoticeVisibilityColumn_(sheet, 6, '公開');
    const deleteCols = ensureSoftDeleteColumns_(sheet);
    const publishAtCol = ensurePublishAtColumn_(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', blogs: [] };

    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const sortCol = header.indexOf('表示順') + 1;
    const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    const blogs = values.map(function (row, index) {
      if (isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol)) return null;
      return {
        rowIdx: index + 2,
        date: formatMaybeDateTime_(row[0]),
        title: row[1],
        category: row[2],
        icon: row[3],
        body: row[4],
        publishStatus: row[5],
        updatedAt: formatMaybeDateTime_(row[6]),
        imageUrl: row[7] || '',
        publishAt: formatMaybeDateTime_(row[publishAtCol - 1]),
        noticeStatus: normalizePublishVisibilityStatus_(row[noticeCol - 1] || row[5] || '公開'),
        sortOrder: sortCol > 0 ? Number(row[sortCol - 1] || 0) : 0
      };
    }).filter(Boolean);
    return { status: 'ok', blogs: blogs };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleDeleteCategory(data) {
  try {
    const name = String(data.name || '').trim();
    if (!name) throw new Error('削除対象のカテゴリが見つかりません');

    const sheet = getCategorySheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('カテゴリデータがありません');

    const values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    const rowsToDelete = [];
    values.forEach(function (row, index) {
      if (String(row[0] || '').trim() === name) rowsToDelete.push(index + 2);
    });

    if (!rowsToDelete.length) throw new Error('削除対象のカテゴリが見つかりません');

    rowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (rowIdx) {
      sheet.deleteRow(rowIdx);
    });

    return { status: 'ok', message: 'カテゴリを削除しました' };
  } catch (err) {
    Logger.log('handleDeleteCategory error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

function getDefaultSupportFaqRows_(updatedAt) {
  const now = updatedAt || getCurrentTime();
  return [
    ['公開', 'アプリ全般', 'このアプリでできることを知りたい', 'アプリ,使い方,できること,何ができる,機能,全体', 'このアプリでは、ホーム、ショップ、カレンダー、NEWS、マイページ、お知らせ一覧、スタンプQR読み取り、商品注文、注文履歴確認、特典確認、通知設定、起動時パスコード、データ引き継ぎ・復元、引き継ぎコード発行、使い方サポートが利用できます。予約確定や個別相談は公式LINEをご利用ください。', 160, now],
    ['公開', 'アプリ全般', 'アプリの画面構成を教えてください', '画面,構成,タブ,ナビ,メニュー,下部,上部', '画面下部にはホーム、ショップ、カレンダー、NEWS、マイページがあります。画面上部からは、最新情報の更新、お知らせ一覧、カート、マイページショートカットを利用できます。', 158, now],
    ['公開', '会員登録', '初回起動時はどちらを選べばいいですか？', '初回,最初,はじめて,以前登録した方,はじめて登録する方,どちら', 'アプリを開いた最初の選択画面で、以前に登録したことがある方は「以前登録した方はこちら ↺」、今回が初めての方は「はじめて登録する方はこちら」を選んでください。再インストール後、機種変更後、ブラウザ版からホーム画面追加した後も、以前登録したことがある方は復元側から進んでください。', 156, now],
    ['公開', 'プロフィール', 'プロフィールの登録方法を知りたい', 'プロフィール,登録,会員,名前,電話,住所,生年月日,初回登録', '初回起動時は、まず「以前登録した方」と「はじめて登録する方」の選択画面が出ます。はじめて登録する方は、そのままプロフィール登録へ進み、お名前・電話番号・生年月日・住所を入力してください。保存すると会員IDが発行され、アプリを使い始められます。', 154, now],
    ['公開', 'プロフィール', 'プロフィールの変更方法を知りたい', 'プロフィール,変更,編集,名前,電話,住所,生年月日', 'マイページを開き、「✏️ プロフィールを編集」を押してください。お名前、電話番号、生年月日、住所、アイコン画像、バナー画像を変更して保存できます。', 152, now],
    ['公開', 'プロフィール', '会員IDはどこで確認できますか？', '会員ID,会員番号,memberid,どこ,確認', '会員IDはマイページ上部に表示されます。プロフィール登録または復元が完了すると発行されます。', 150, now],
    ['公開', 'ログイン', '起動時のパスコード設定について知りたい', 'パスコード,ログイン,起動時,4桁,6桁,設定', '新しく登録する方も、すでに登録済みの方も、まずは4桁または6桁のパスコードを設定して使います。既存会員の方はアプリ起動時に設定画面が表示されます。', 148, now],
    ['公開', 'ログイン', 'ログイン時のパスコードを毎回入力したくないです', 'ログイン時のパスコード,毎回,入力したくない,オフ,省略', 'パスコードを一度設定したあと、マイページの「ログイン時のパスコード」からオン・オフを切り替えられます。オフにすると、次回からアプリ起動時のパスコード入力を省略できます。', 146, now],
    ['公開', 'ログイン', 'パスコードの変更方法を知りたい', 'パスコード,変更,変える,ログイン,再設定', 'マイページを開き、「🔐 パスコードを変更」を押してください。現在のパスコードを確認したあと、新しい4桁または6桁のパスコードへ変更できます。', 144, now],
    ['公開', 'ログイン', 'パスコードを忘れたときの再設定方法を知りたい', 'パスコード,忘れた,再設定,ログインできない', 'ログイン画面、または「ログイン・引き継ぎ」画面にある「パスコードを忘れた場合の再設定はこちら」から再設定できます。登録したお名前・電話番号・生年月日を入力し、新しい4桁または6桁のパスコードを設定してください。', 142, now],
    ['公開', '引き継ぎ', 'データの引き継ぎ・復元方法を知りたい', '引き継ぎ,復元,機種変更,データ移行,ログイン,再インストール', 'ログイン画面、または初回画面で「以前登録した方はこちら ↺」を選ぶと復元画面へ進めます。引き継ぎコードがある場合は、引き継ぎコードと新しいパスコードを入力してください。引き継ぎコードがない場合は、お名前に加えて電話番号・生年月日・現在のパスコードのうち1つ以上を入力すると復元できます。', 140, now],
    ['公開', '引き継ぎ', '引き継ぎコードの発行方法を知りたい', '引き継ぎコード,発行,機種変更,再インストール,コード', 'マイページの「↺ 引き継ぎコードの発行」から発行できます。機種変更や再インストールの前に発行しておくと、新しい端末の「データの引き継ぎ・復元」で使えます。', 138, now],
    ['公開', '引き継ぎ', '引き継ぎコードの有効期限を知りたい', '引き継ぎコード,有効期限,いつまで,何日,使えない', '引き継ぎコードは1回限りで、発行から1週間有効です。期限切れ、または一度使用したコードは使えません。必要な場合はマイページから新しいコードを発行してください。', 136, now],
    ['公開', '会員登録', '会員登録が重複しないようにする方法を知りたい', '重複,二重,会員登録,同じ名前,会員ID,ブラウザ,ホーム画面', '以前登録したことがある方は、新規登録へ進まず、必ず「以前登録した方はこちら ↺」から復元してください。ブラウザで先に登録したあとホーム画面に追加した方、再インストールした方、機種変更した方も同じです。新規登録をすると、同じお名前でも別の会員IDが作られることがあります。', 134, now],
    ['公開', 'ホーム画面', 'ホーム画面への追加方法を知りたい', 'ホーム画面,追加,インストール,iphone,android,safari,chrome', 'iPhone は Safari でサイトを開き、共有ボタンから「ホーム画面に追加」を選びます。Android は Chrome のメニューから「ホーム画面に追加」または「アプリをインストール」を選びます。すでに登録済みの方は、追加後に新規登録せず「以前登録した方はこちら ↺」から入ってください。', 132, now],
    ['公開', 'ホーム', 'ホーム画面の見方を知りたい', 'ホーム,トップ,home,見方', 'ホーム画面では、スタンプQRの読み取り、現在のスタンプカード確認、おすすめ商品、最新のNEWS、メニュー一覧への移動、公式サイト・SNSへのリンクを確認できます。', 130, now],
    ['公開', '注文', '商品の注文方法を知りたい', '注文,買い方,ショップ,カート,購入', '下部メニューの「ショップ」を開き、商品を選んで詳細を確認し、「カートに追加する」を押してください。画面上部の🛒からカートを開き、内容を確認して「ご注文を確定する」で注文できます。', 128, now],
    ['公開', '注文', 'カートの使い方を知りたい', 'カート,買い物かご,個数,削除,変更', '商品を「カートに追加する」で入れたあと、画面上部の🛒を押すとカートを開けます。カートでは個数変更や削除ができます。商品が入っていない場合は空の画面が表示されます。', 126, now],
    ['公開', '注文', '支払い方法を知りたい', '支払い,決済,現金,支払方法', '商品注文のお支払い方法は現在「現金払い」です。ご来院時に受付でお支払いください。', 124, now],
    ['公開', '注文', '商品の受け取り方法を知りたい', '受取,受け取り,配送,宅配,受領', '商品は院内受け取りです。注文後、ご来院時にスタッフへお声がけください。配送には対応していません。', 122, now],
    ['公開', '注文', '注文履歴の見方を知りたい', '注文履歴,履歴,受取,受付中,キャンセル', 'マイページの「📋 ご注文履歴」から確認できます。受付中の注文や受け取り前の注文が表示され、状況を確認できます。', 120, now],
    ['公開', '注文', '注文をキャンセルしたいです', '注文,キャンセル,取り消し,受付中', 'マイページの「📋 ご注文履歴」を開き、受付中の注文に表示される「キャンセルする」を押してください。キャンセル済みの注文は履歴から表示されなくなります。', 118, now],
    ['公開', '注文', '受け取りましたボタンの使い方を知りたい', '受け取りました,受取完了,受取報告,注文履歴', 'マイページの「📋 ご注文履歴」を開き、受け取り済みにしたい注文の「受け取りました」を押してください。更新後、その注文は履歴に残らなくなります。', 116, now],
    ['公開', 'メニュー', 'メニュー一覧の見方を知りたい', 'メニュー,施術,一覧,カテゴリ,見方', 'ホーム画面の「メニュー一覧を見る🍴」を押すと一覧が開きます。必要に応じてカテゴリで絞り込みでき、各メニューをタップすると詳細や画像を確認できます。', 114, now],
    ['公開', '予約', '予約はアプリからできますか？', '予約,よやく,line,予約方法,相談', '現在、アプリから予約確定はできません。予約や個別相談は公式LINEをご利用ください。ホーム画面またはマイページの「🔗 公式サイト・SNS」から公式LINEを開けます。', 112, now],
    ['公開', 'スタンプ', 'スタンプの集め方を知りたい', 'スタンプ,QR,QRコード,来院,カメラ', 'ホーム画面の「📷 カメラを起動して読み取る」を押し、表示された案内でカメラを許可してから院内QRコードを読み取ってください。読み取りに成功するとスタンプが1つ追加されます。', 110, now],
    ['公開', 'スタンプ', 'スタンプは1日何回取得できますか？', 'スタンプ,1日,一日,何回,回数', '来院スタンプは1日1回までです。同じ日に再度読み取ると、すでに取得済みの案内が表示されます。', 108, now],
    ['公開', 'トラブル', 'カメラが起動しないときはどうすればいいですか？', 'カメラ,起動しない,許可,権限,QR,読めない', 'スタンプ取得時にカメラ許可の確認が出た場合は「許可」を選んでください。すでに拒否している場合は、表示される「設定を開く」から設定画面へ進み、iPhone や Android のカメラ許可をオンにしてから、もう一度「📷 カメラを起動して読み取る」を押してください。', 106, now],
    ['公開', 'スタンプ特典', 'スタンプが10個たまったらどうなりますか？', 'スタンプ,10個,達成,ガチャ,特典', 'スタンプが10個たまると、ホーム画面から特典ガチャを回せます。結果はマイページの「🎁 スタンプ・特典履歴」で確認できます。ガチャ後はホーム画面の「🌸 新しいスタンプカードを取得」から次のカードを始められます。', 104, now],
    ['公開', 'スタンプ特典', '特典はどこで確認できますか？', '特典,どこ,確認,プレゼント,ガチャ', '特典はマイページの「🎁 スタンプ・特典履歴」で確認できます。未使用の特典、使用済みの特典、受取期限を確認できます。', 102, now],
    ['公開', 'スタンプ特典', '特典の有効期限を知りたい', '特典,期限,有効期限,いつまで', '特典の受取期限は、スタンプ10個を達成した日から1か月です。期限はマイページの「🎁 スタンプ・特典履歴」に表示されます。', 100, now],
    ['公開', '通知', '通知をオン・オフにしたい', '通知,オン,オフ,push,プッシュ,許可', 'マイページの「🔔 通知設定」からオン・オフを切り替えられます。アプリ内でオンにしても届かない場合は、iPhone や Android 本体側の通知許可もご確認ください。', 98, now],
    ['公開', '通知', '通知が届かないときはどうすればいいですか？', '通知,届かない,push,プッシュ,こない', 'まずマイページの「🔔 通知設定」がオンか確認してください。そのうえで、iPhone や Android 本体側の通知許可、通信状態、アプリの最新化をご確認ください。必要に応じて画面上部の🔄で最新情報を再取得してください。', 96, now],
    ['公開', '更新', '最新情報への更新方法を知りたい', '更新,最新,再読み込み,リロード,refresh,最新情報', '画面上部の「🔄」ボタンを押すと、最新のNEWS、商品、カレンダー、メニュー、FAQ、注文履歴などを更新できます。通常の情報更新は再インストール不要です。', 94, now],
    ['公開', '更新', 'アップデートが必要と表示されたらどうすればいいですか？', 'アップデート,更新が必要,app store,最新版,バージョン', '「アップデートが必要です」と表示された場合は、案内の「アップデートする」から最新版へ更新してください。画面上部の🔄は情報更新用で、必須アップデートの代わりにはなりません。', 92, now],
    ['公開', 'カレンダー', 'イベントカレンダーの見方を知りたい', 'カレンダー,イベント,予定,日程', '下部メニューの「カレンダー」で予定を確認できます。左右の矢印で別の月に切り替えられ、日付を押すと詳細を確認できます。下部には今月のイベント一覧も表示されます。', 90, now],
    ['公開', 'カレンダー', 'カレンダーの記号の意味を知りたい', 'カレンダー,記号,意味,休,往,イ', 'カレンダーでは、休＝休診日、往＝往診日、イ＝イベントを表しています。日付を押すと、その日の詳しい内容を確認できます。', 88, now],
    ['公開', 'NEWS', 'NEWSページの使い方を知りたい', 'NEWS,ニュース,お知らせ,記事,カテゴリ', '下部メニューの「NEWS」を開くと記事一覧を確認できます。記事をタップすると詳細が開き、右上のカテゴリ選択で絞り込みもできます。', 86, now],
    ['公開', 'NEWS', 'お知らせ一覧の見方を知りたい', 'お知らせ一覧,通知一覧,拡声器,📢', '画面上部の📢ボタンを押すと「お知らせ一覧」を開けます。ここでは NEWS、カレンダー、ショップ、ホームの更新情報を新しい順で確認できます。カテゴリの絞り込みもできます。', 84, now],
    ['公開', 'NEWS', 'NEWSのカテゴリ切り替え方法を知りたい', 'NEWS,カテゴリ,切り替え,絞り込み,全て', 'NEWSページ右上のカテゴリ選択を押すと、カテゴリごとに絞り込みできます。「全て」を選ぶとすべての記事が表示されます。', 82, now],
    ['公開', 'NEWS', 'まゆみのつぶやきはどこで見られますか？', 'つぶやき,NEWS,カテゴリ,まゆみのつぶやき', '「まゆみのつぶやき」は NEWS ページ右上のカテゴリ選択から「まゆみのつぶやき」を選ぶと表示されます。院長からの短いメッセージや大切なお知らせを確認できます。', 80, now],
    ['公開', 'NEWS', 'まゆみのブログとは何ですか？', 'まゆみのブログ,ブログ,外部ブログ', '「まゆみのブログ」はマイページやホームの「🔗 公式サイト・SNS」から開ける外部ブログです。NEWS内の「まゆみのつぶやき」とは別の場所です。', 78, now],
    ['公開', 'リンク', '公式LINEやSNSの開き方を知りたい', 'LINE,ライン,instagram,facebook,ホームページ,公式サイト,SNS,問い合わせ', 'ホーム画面またはマイページの「🔗 公式サイト・SNS」を開くと、公式ホームページ、Instagram、Facebook、公式LINE、まゆみのブログを選んで開けます。', 76, now],
    ['公開', '使い方サポート', '使い方チャットでは何を質問できますか？', 'チャット,サポート,ボット,相談,何が聞ける', '使い方チャットでは、登録、復元、パスコード、注文、注文履歴、スタンプ、特典、通知、NEWS、お知らせ一覧、カレンダー、メニュー一覧、更新方法など、アプリの使い方について質問できます。診療相談や個別予約は公式LINEをご利用ください。', 74, now],
    ['公開', '引き継ぎ', '再インストールしたあとの入り方を知りたい', '再インストール,削除,アンインストール,復元,入り方', 'アプリを入れ直したあとは、新規登録ではなく、初回画面またはログイン画面の「以前登録した方はこちら ↺」から復元してください。引き継ぎコードがある場合はコードで、ない場合はお名前と電話番号・生年月日・現在のパスコードのうち1つ以上で復元できます。', 72, now],
    ['公開', 'トラブル', '画面表示がおかしい・アプリが重いときはどうすればいいですか？', '表示されない,おかしい,崩れ,不具合,バグ,重い,遅い,フリーズ', 'まず画面上部の🔄ボタンで最新情報を再取得してください。それでも改善しない場合は、アプリを一度閉じて再起動し、通信状態もご確認ください。再インストールが必要な場合は、先に引き継ぎコードを発行するか、復元方法を確認してから行ってください。', 70, now],
  ];
}

function getSupportFaqSheet_() {
  const ss = getOrCreateSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.APP_SUPPORT_FAQ);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.APP_SUPPORT_FAQ);
    sheet.appendRow(['状態', 'カテゴリ', '質問', 'キーワード', '回答', '優先度', '更新日時']);
  }

  if (sheet.getLastRow() < 2) {
    const defaults = getDefaultSupportFaqRows_(getCurrentTime());
    sheet.getRange(2, 1, defaults.length, 7).setValues(defaults);
  }

  return sheet;
}

function normalizeSupportText_(text) {
  return String(text || '')
    .toLowerCase()
    .normalize('NFKC')
    .replace(/\s+/g, '');
}

function splitSupportKeywords_(keywords) {
  return String(keywords || '')
    .split(/[\n,、，\/\s]+/)
    .map(function (item) { return item.trim(); })
    .filter(function (item) { return item; });
}

function getSupportFaqEntries_(includePrivate) {
  const sheet = getSupportFaqSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 7).getDisplayValues();
  return values.map(function (row, index) {
    return {
      rowIdx: index + 2,
      status: row[0] || '公開',
      category: row[1] || '',
      question: row[2] || '',
      keywords: row[3] || '',
      answer: row[4] || '',
      priority: Number(row[5] || 0),
      updatedAt: row[6] || ''
    };
  }).filter(function (item) {
    if (!item.question || !item.answer) return false;
    return includePrivate ? true : item.status === '公開';
  }).sort(function (a, b) {
    if (b.priority !== a.priority) return b.priority - a.priority;
    return a.rowIdx - b.rowIdx;
  });
}

function scoreSupportFaqEntry_(messageNorm, entry) {
  if (!messageNorm) return 0;

  let score = 0;
  const questionNorm = normalizeSupportText_(entry.question);
  const categoryNorm = normalizeSupportText_(entry.category);
  const keywords = splitSupportKeywords_(entry.keywords);

  if (questionNorm) {
    if (messageNorm === questionNorm) score += 14;
    if (messageNorm.indexOf(questionNorm) !== -1 || questionNorm.indexOf(messageNorm) !== -1) score += 8;
  }

  if (categoryNorm && messageNorm.indexOf(categoryNorm) !== -1) {
    score += 3;
  }

  keywords.forEach(function (keyword) {
    const norm = normalizeSupportText_(keyword);
    if (!norm) return;
    if (messageNorm.indexOf(norm) !== -1) {
      score += norm.length >= 4 ? 5 : 3;
    }
  });

  return score;
}

function buildSupportFallbackAnswer_(message) {
  const norm = normalizeSupportText_(message);

  if (norm.indexOf('パスコード') !== -1 || norm.indexOf('ログイン') !== -1) {
    return 'パスコードは4桁または6桁の数字で設定します。変更はマイページの「パスコードを変更」から行え、起動時に毎回入力するかどうかはマイページの「ログイン時のパスコード」からオン・オフを切り替えられます。忘れた場合はログイン画面の「パスコードを忘れた場合の再設定はこちら」から再設定してください。';
  }
  if (norm.indexOf('引き継ぎ') !== -1 || norm.indexOf('復元') !== -1 || norm.indexOf('機種変更') !== -1) {
    return 'ログイン画面、または初回画面の「以前登録した方はこちら ↺」から、引き継ぎコード、またはお名前と電話番号・生年月日・現在のパスコードのうち1つ以上で復元できます。機種変更前や再インストール前はマイページで引き継ぎコードを発行しておくとスムーズです。';
  }
  if (norm.indexOf('ホーム画面') !== -1 || norm.indexOf('追加') !== -1 || norm.indexOf('インストール') !== -1) {
    return 'iPhone は Safari の共有ボタンから「ホーム画面に追加」、Android は Chrome のメニューから「ホーム画面に追加」または「アプリをインストール」を選んでください。すでに会員登録済みの場合は、新規登録ではなく「以前登録した方はこちら ↺」からお入りください。';
  }
  if (norm.indexOf('注文') !== -1 || norm.indexOf('カート') !== -1 || norm.indexOf('ショップ') !== -1 || norm.indexOf('購入') !== -1) {
    return '下部メニューの「ショップ」から商品を選び、「カートに追加する」を押してご注文ください。注文後はマイページで履歴を確認できます。';
  }
  if (norm.indexOf('プロフィール') !== -1 || norm.indexOf('登録') !== -1 || norm.indexOf('会員') !== -1) {
    return 'マイページの「プロフィールを編集」から、お名前や電話番号などを登録・変更できます。初回起動時は案内に沿って登録してください。';
  }
  if (norm.indexOf('スタンプ') !== -1 || norm.indexOf('qr') !== -1 || norm.indexOf('qrcode') !== -1) {
    return 'ホーム画面にある「カメラを起動して読み取る」から院内QRコードを読み取るとスタンプが追加されます。';
  }
  if (norm.indexOf('通知') !== -1 || norm.indexOf('push') !== -1 || norm.indexOf('プッシュ') !== -1) {
    return '通知はマイページの「通知設定」からオンオフを切り替えられます。うまくいかない場合は端末の通知設定も確認してください。';
  }
  if (norm.indexOf('カレンダー') !== -1 || norm.indexOf('イベント') !== -1) {
    return '下部メニューの「カレンダー」でイベント予定を確認できます。左右の矢印で別の月にも切り替えられます。';
  }
  if (norm.indexOf('news') !== -1 || norm.indexOf('お知らせ') !== -1 || norm.indexOf('ブログ') !== -1) {
    return 'NEWSは下部メニューの「NEWS」で確認でき、更新情報の一覧は画面上部の📢ボタンから確認できます。お知らせ一覧では NEWS、カレンダー、ショップ、ホームの更新情報を新しい順で見られます。';
  }

  return 'この質問はまだ個別の自動回答に登録されていません。ホーム、ショップ、カレンダー、NEWS、マイページ、お知らせ一覧、スタンプ、注文、通知、復元、パスコードなどの使い方をご案内できます。解決しない場合は公式LINEや院への直接お問い合わせをご案内してください。';
}

function buildSupportSuggestions_(entries, matchedRowIdx) {
  return entries
    .filter(function (item) { return item.rowIdx !== matchedRowIdx; })
    .slice(0, 3)
    .map(function (item) { return item.question; });
}

function getSupportFaq() {
  try {
    const faqs = getSupportFaqEntries_(false).map(function (item) {
      return {
        rowIdx: item.rowIdx,
        category: item.category,
        question: item.question,
        keywords: item.keywords,
        answer: item.answer,
        priority: item.priority,
        updatedAt: item.updatedAt || ''
      };
    });
    return { status: 'ok', faqs: faqs };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getAdminSupportFaq() {
  try {
    return { status: 'ok', faqs: getSupportFaqEntries_(true) };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleSaveSupportFaq(data) {
  try {
    const question = String(data.question || '').trim();
    const answer = String(data.answer || '').trim();
    if (!question) throw new Error('質問を入力してください');
    if (!answer) throw new Error('回答を入力してください');

    const row = [
      data.status === '非公開' ? '非公開' : '公開',
      String(data.category || '').trim(),
      question,
      String(data.keywords || '').trim(),
      answer,
      Number(data.priority || 0),
      getCurrentTime()
    ];

    const sheet = getSupportFaqSheet_();
    const rowIdx = Number(data.rowIdx || 0);
    if (rowIdx > 1) {
      sheet.getRange(rowIdx, 1, 1, 7).setValues([row]);
      return { status: 'ok', message: 'FAQを更新しました' };
    }

    sheet.appendRow(row);
    return { status: 'ok', message: 'FAQを追加しました' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleDeleteSupportFaq(data) {
  try {
    const rowIdx = Number(data.rowIdx || 0);
    if (rowIdx <= 1) throw new Error('削除対象が見つかりません');

    const sheet = getSupportFaqSheet_();
    sheet.deleteRow(rowIdx);
    return { status: 'ok', message: 'FAQを削除しました' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function askSupportChat(data) {
  try {
    const message = String(data.message || '').trim();
    if (!message) throw new Error('質問を入力してください');

    const entries = getSupportFaqEntries_(false);
    const messageNorm = normalizeSupportText_(message);
    const ranked = entries.map(function (item) {
      return {
        item: item,
        score: scoreSupportFaqEntry_(messageNorm, item)
      };
    }).sort(function (a, b) {
      if (b.score !== a.score) return b.score - a.score;
      return b.item.priority - a.item.priority;
    });

    const top = ranked[0];
    if (top && top.score > 0) {
      const related = ranked
        .filter(function (row) { return row.score > 0 && row.item.rowIdx !== top.item.rowIdx; })
        .map(function (row) { return row.item.question; })
        .slice(0, 3);

      const response = {
        status: 'ok',
        answer: top.item.answer,
        matchedQuestion: top.item.question,
        suggestions: related.length ? related : buildSupportSuggestions_(entries, top.item.rowIdx)
      };
      appendSupportChatLog_({
        message: message,
        matchedQuestion: top.item.question,
        category: top.item.category,
        memberId: data.memberId,
        answer: top.item.answer
      });
      return response;
    }

    const fallback = {
      status: 'ok',
      answer: buildSupportFallbackAnswer_(message),
      matchedQuestion: '',
      suggestions: buildSupportSuggestions_(entries, -1)
    };
    appendSupportChatLog_({
      message: message,
      matchedQuestion: '',
      category: '',
      memberId: data.memberId,
      answer: fallback.answer
    });
    return fallback;
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// ========== メニュー管理（MENUS） ==========

function ensureMenusSheetStructure_(sheet) {
  if (!sheet) return null;

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 8).setValues([[
      '登録日', 'メニュー名', '画像URL', '概要説明', '予約状況', '公開設定', '更新日時', 'カテゴリ'
    ]]);
  }

  ensureUpdatedAtColumn_(sheet, '更新日時');

  const lastCol = Math.max(sheet.getLastColumn(), 8);
  if (sheet.getLastColumn() < lastCol) {
    sheet.insertColumnsAfter(sheet.getLastColumn(), lastCol - sheet.getLastColumn());
  }

  let header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const sortCol = header.indexOf('表示順') + 1;
  let categoryCol = header.indexOf('カテゴリ') + 1;
  if (categoryCol === 0) {
    if (sortCol === 8) {
      sheet.insertColumnBefore(8);
    } else if (sheet.getLastColumn() < 8) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), 8 - sheet.getLastColumn());
    }
    sheet.getRange(1, 8).setValue('カテゴリ');
  }

  sheet.getRange(1, 1, 1, 8).setValues([[
    '登録日', 'メニュー名', '画像URL', '概要説明', '予約状況', '公開設定', '更新日時', 'カテゴリ'
  ]]);
  ensureSortOrderColumn_(sheet);
  styleHeader(sheet, sheet.getLastColumn(), '#4caf50');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 320);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 140);
  ensureNoticeVisibilityColumn_(sheet, 6, '公開');
  ensurePublishAtColumn_(sheet);
  ensureSoftDeleteColumns_(sheet);
  return sheet;
}

/**
 * ユーザー用：公開済みメニュー一覧取得
 */
function getMenus() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ensureMenusSheetStructure_(ss.getSheetByName(SHEETS.MENUS));
    if (!sheet) return { status: 'ok', menus: [] };
    const noticeCol = ensureNoticeVisibilityColumn_(sheet, 6, '公開');
    const publishAtCol = ensurePublishAtColumn_(sheet);
    const deleteCols = ensureSoftDeleteColumns_(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', menus: [] };

    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const sortCol = header.indexOf('表示順') + 1;
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    const menus = data
      .map(function (row, index) {
        if (row[5] !== '公開' ||
          isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol) ||
          !isPublishAtAvailable_(row[publishAtCol - 1])) {
          return null;
        }
        return {
          rowIdx: index + 2,
          date: formatMaybeDateTime_(row[0]),
          name: row[1],
          imageUrl: row[2],
          description: row[3],
          reservationStatus: row[4],
          category: row[7] || '',
          updatedAt: formatMaybeDateTime_(row[6]),
          publishAt: formatMaybeDateTime_(row[publishAtCol - 1]),
          noticeStatus: normalizePublishVisibilityStatus_(row[noticeCol - 1] || row[5] || '公開'),
          sortOrder: sortCol > 0 ? Number(row[sortCol - 1] || 0) : 0
        };
      }).filter(Boolean);

    return { status: 'ok', menus: menus };
  } catch (err) {
    Logger.log('getMenus error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

/**
 * 管理者用：全メニュー一覧取得
 */
function getAdminMenus() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ensureMenusSheetStructure_(ss.getSheetByName(SHEETS.MENUS));
    if (!sheet) return { status: 'ok', menus: [] };
    const noticeCol = ensureNoticeVisibilityColumn_(sheet, 6, '公開');
    const publishAtCol = ensurePublishAtColumn_(sheet);
    const deleteCols = ensureSoftDeleteColumns_(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'ok', menus: [] };

    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const sortCol = header.indexOf('表示順') + 1;
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    const menus = data.map(function (row, index) {
      if (isSoftDeletedByColumns_(row, deleteCols.statusCol, deleteCols.deletedAtCol)) return null;
      return {
        rowIdx: index + 2,
        date: formatMaybeDateTime_(row[0]),
        name: row[1],
        imageUrl: row[2],
        description: row[3],
        reservationStatus: row[4],
        publishStatus: row[5],
        category: row[7] || '',
        updatedAt: formatMaybeDateTime_(row[6]),
        publishAt: formatMaybeDateTime_(row[publishAtCol - 1]),
        noticeStatus: normalizePublishVisibilityStatus_(row[noticeCol - 1] || row[5] || '公開'),
        sortOrder: sortCol > 0 ? Number(row[sortCol - 1] || 0) : 0
      };
    }).filter(Boolean);

    return { status: 'ok', menus: menus };
  } catch (err) {
    Logger.log('getAdminMenus error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

/**
 * メニュー追加
 */
function handleAddMenu(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    let sheet = ss.getSheetByName(SHEETS.MENUS);
    if (!sheet) sheet = ss.insertSheet(SHEETS.MENUS);
    sheet = ensureMenusSheetStructure_(sheet);
    const noticeCol = ensureNoticeVisibilityColumn_(sheet, 6, '公開');
    const publishAtCol = ensurePublishAtColumn_(sheet);

    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
    const updatedAt = formatDateTime_(new Date());
    const sortOrder = new Date().getTime();
    const sortCol = ensureSortOrderColumn_(sheet);
    const rowData = new Array(Math.max(sheet.getLastColumn(), sortCol)).fill('');
    rowData[0] = today;
    rowData[1] = data.name || '';
    rowData[2] = data.imageUrl || '';
    rowData[3] = data.description || '';
    rowData[4] = data.reservationStatus || '予約対象外';
    rowData[5] = data.publishStatus || '非公開';
    rowData[6] = updatedAt;
    rowData[7] = data.category || '';
    rowData[sortCol - 1] = sortOrder;
    sheet.appendRow(rowData);
    const rowIdx = sheet.getLastRow();
    sheet.getRange(rowIdx, noticeCol).setValue(
      normalizePublishVisibilityStatus_(data.noticeStatus || data.publishStatus || '公開')
    );
    if (publishAtCol) {
      sheet.getRange(rowIdx, publishAtCol).setValue(normalizePublishAtValue_(data.publishAt));
    }

    if (String(data.publishStatus || '非公開') === '公開' && isPublishAtAvailable_(data.publishAt)) {
      sendAutoPush('🍴 ' + (data.name || 'ホーム更新'), 'ホームのメニュー一覧が更新されました', {
        targetPage: 'menu-list'
      });
    }

    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleAddMenu error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

/**
 * メニュー更新
 */
function handleUpdateMenu(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ensureMenusSheetStructure_(ss.getSheetByName(SHEETS.MENUS));
    const noticeCol = ensureNoticeVisibilityColumn_(sheet, 6, '公開');
    const publishAtCol = ensurePublishAtColumn_(sheet);
    const rowIdx = Number(data.rowIdx);

    if (data.name) sheet.getRange(rowIdx, 2).setValue(data.name);
    if (data.imageUrl !== undefined) sheet.getRange(rowIdx, 3).setValue(data.imageUrl);
    if (data.description !== undefined) sheet.getRange(rowIdx, 4).setValue(data.description);
    if (data.reservationStatus) sheet.getRange(rowIdx, 5).setValue(data.reservationStatus);
    if (data.publishStatus) sheet.getRange(rowIdx, 6).setValue(data.publishStatus);
    if (data.category !== undefined) sheet.getRange(rowIdx, 8).setValue(data.category);
    sheet.getRange(rowIdx, 7).setValue(formatDateTime_(new Date()));
    if (data.noticeStatus) {
      sheet.getRange(rowIdx, noticeCol).setValue(normalizePublishVisibilityStatus_(data.noticeStatus));
    }
    if (publishAtCol && data.publishAt !== undefined) {
      sheet.getRange(rowIdx, publishAtCol).setValue(normalizePublishAtValue_(data.publishAt));
    }

    const effectiveStatus = data.publishStatus !== undefined ? data.publishStatus : sheet.getRange(rowIdx, 6).getValue();
    const effectivePublishAt = publishAtCol ? sheet.getRange(rowIdx, publishAtCol).getValue() : '';
    if (String(effectiveStatus || '非公開') === '公開' && isPublishAtAvailable_(effectivePublishAt)) {
      sendAutoPush('🍴 ' + (data.name || 'ホーム更新'), 'ホームのメニュー一覧が更新されました', {
        targetPage: 'menu-list'
      });
    }

    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleUpdateMenu error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

/**
 * メニュー削除
 */
function handleDeleteMenu(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ensureMenusSheetStructure_(ss.getSheetByName(SHEETS.MENUS));
    const rowIdx = Number(data.rowIdx);
    if (sheet && rowIdx > 1) {
      markRowSoftDeleted_(sheet, rowIdx, '管理画面から削除');
      return { status: 'ok' };
    }
    return { status: 'error', message: '削除失敗' };
  } catch (err) {
    Logger.log('handleDeleteMenu error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}

/**
 * メニューの並び替え（行の入れ替え）
 */
function handleMoveMenu(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ensureMenusSheetStructure_(ss.getSheetByName(SHEETS.MENUS));
    if (!sheet) return { status: 'error', message: 'Sheet not found' };

    const rowIdx = Number(data.rowIdx);
    const direction = data.direction; // 'up' or 'down'

    if (isNaN(rowIdx) || rowIdx <= 1) return { status: 'error', message: 'Invalid row index' };

    const lastRow = sheet.getLastRow();
    const targetIdx = (direction === 'up') ? rowIdx - 1 : rowIdx + 1;

    if (targetIdx <= 1 || targetIdx > lastRow) {
      return { status: 'error', message: 'Cannot move further in this direction' };
    }

    // 行の内容を入れ替える
    const range1 = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn());
    const range2 = sheet.getRange(targetIdx, 1, 1, sheet.getLastColumn());
    const values1 = range1.getValues();
    const values2 = range2.getValues();

    range1.setValues(values2);
    range2.setValues(values1);

    // 更新日時を両方の行で更新
    const now = formatDateTime_(new Date());
    sheet.getRange(rowIdx, 7).setValue(now);
    sheet.getRange(targetIdx, 7).setValue(now);

    return { status: 'ok' };
  } catch (err) {
    Logger.log('handleMoveMenu error: ' + err.toString());
    return { status: 'error', message: err.toString() };
  }
}
/**
 * Google Reviews API Integration
 * Requires OAuth2 library: 1B7FSOmYNVu6WVm_baPG_olPgdP_SXvfgmUpgab7L8Zd95HCPYp40F9K1
 */
