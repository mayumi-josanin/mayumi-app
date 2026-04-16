const CUSTOMER_KEY = "mayumi_survey_customer";
const DRAFTS_KEY = "mayumi_survey_drafts";
const PENDING_KEY = "mayumi_survey_pending_submission";
const TICKET_CARD_OVERRIDE_KEY = "mayumi_survey_ticket_card_overrides";
const BIJIRIS_FAVORITES_KEY = "mayumi_bijiris_favorites";
const BIJIRIS_READER_STATE_KEY = "mayumi_bijiris_reader_state";
const PUSH_ENABLED_STORAGE_KEY = "mayumi_customer_push_enabled";
const PUSH_APP_ID_STORAGE_KEY = "mayumi_customer_push_app_id";
const PHOTO_FILE_LIMIT = 6;
const PHOTO_MAX_SIZE = 1400;
const PHOTO_JPEG_QUALITY = 0.74;
const RESPONSE_EDIT_WINDOW_MS = 24 * 60 * 60 * 1000;
const BIJIRIS_NEW_BADGE_DAYS = 7;
const BIJIRIS_HISTORY_LIMIT = 8;
const APP_VERSION = "20260416-09";
const CACHE_PREFIX = "mayumi-customer-survey-";
const ACTIVE_CACHE_NAME = "mayumi-customer-survey-v70";
const AUTO_CACHE_MAINTENANCE_INTERVAL_MS = 6 * 60 * 60 * 1000;
const AUTO_CACHE_MAINTENANCE_KEY = "mayumi_customer_cache_maintenance_at";
const DEFAULT_ONESIGNAL_APP_ID = "88023099-c99e-44c6-9f7c-2ef08d363768";
const SESSION_SURVEY_ID = "survey_bijiris_session";
const SESSION_TYPE_QUESTION_ID = "q_bijiris_session_type";
const SESSION_TICKET_PLAN_QUESTION_ID = "q_bijiris_session_ticket_plan";
const SESSION_TICKET_SHEET_QUESTION_ID = "q_bijiris_session_ticket_sheet";
const SESSION_TICKET_ROUND_QUESTION_ID = "q_bijiris_session_ticket_round";
const LEGACY_SESSION_MONITOR_PHOTOS_QUESTION_ID = "q_bijiris_session_monitor_photos";
const LEGACY_SESSION_TICKET_END_PHOTOS_QUESTION_ID = "q_bijiris_session_ticket_end_photos";
const SESSION_MONITOR_PHOTOS_QUESTION_IDS = [
  "q_bijiris_session_monitor_photos_6",
  "q_bijiris_session_monitor_photos_10",
  LEGACY_SESSION_MONITOR_PHOTOS_QUESTION_ID,
];
const SESSION_TICKET_END_PHOTOS_QUESTION_IDS = [
  "q_bijiris_session_ticket_end_photos_6",
  "q_bijiris_session_ticket_end_photos_10",
  LEGACY_SESSION_TICKET_END_PHOTOS_QUESTION_ID,
];
const SESSION_CONCERN_QUESTION_ID = "q_bijiris_session_concern";
const SESSION_CONCERN_OTHER_QUESTION_ID = "q_bijiris_session_concern_other";
const SESSION_CONCERN_OTHER_OPTION = "その他（長文）";
const SESSION_LIFE_CHANGES_QUESTION_ID = "q_bijiris_session_life_changes";
const SESSION_LIFE_CHANGES_OTHER_QUESTION_ID = "q_bijiris_session_life_changes_other";
const TICKET_END_SURVEY_ID = "survey_bijiris_ticket_end";
const BIJIRIS_RECOMMENDATION_TOPICS = [
  {
    id: "toilet",
    answerTerms: ["トイレ", "尿", "尿もれ", "尿意", "夜中", "パッド", "くしゃみ", "咳"],
    postTerms: ["トイレ", "尿", "尿もれ", "頻尿", "夜間", "尿意", "咳", "くしゃみ", "パッド"],
  },
  {
    id: "pelvic",
    answerTerms: ["骨盤", "骨盤底筋", "体幹", "インナーマッスル", "内側", "締める感覚"],
    postTerms: ["骨盤", "骨盤底筋", "体幹", "インナー", "内側", "支える力"],
  },
  {
    id: "belly",
    answerTerms: ["お腹", "下腹", "便秘", "便通", "ぽっこり"],
    postTerms: ["お腹", "下腹", "便秘", "便通", "腹圧", "ぽっこり"],
  },
  {
    id: "posture",
    answerTerms: ["姿勢", "猫背", "反り腰", "立ち姿", "歩き方"],
    postTerms: ["姿勢", "猫背", "反り腰", "立ち方", "歩き方"],
  },
  {
    id: "lower-body",
    answerTerms: ["下半身", "股関節", "お尻", "ヒップ", "太もも", "腰"],
    postTerms: ["下半身", "股関節", "お尻", "ヒップ", "太もも", "腰"],
  },
  {
    id: "postpartum-aging",
    answerTerms: ["産後", "出産後", "年齢", "更年期", "予防", "将来"],
    postTerms: ["産後", "出産後", "年齢", "更年期", "予防", "将来"],
  },
  {
    id: "daily-life",
    answerTerms: ["外出", "旅行", "移動", "会議", "授業", "運動", "抱っこ", "日常"],
    postTerms: ["外出", "旅行", "移動", "会議", "授業", "運動", "抱っこ", "日常"],
  },
  {
    id: "self-care",
    answerTerms: ["セルフケア", "呼吸", "ストレッチ", "整える"],
    postTerms: ["セルフケア", "呼吸", "ストレッチ", "整える", "ホームケア"],
  },
];
const LEGACY_TICKET_INFO_QUESTION_IDS = {
  size: "q_ticket_end_ticket_size",
  sheet: "q_ticket_end_ticket_sheet",
  round: "q_ticket_end_ticket_round",
};
const TICKET_END_COUNT_QUESTION_ID = LEGACY_TICKET_INFO_QUESTION_IDS.size;
const TICKET_END_SHEET_QUESTION_ID = LEGACY_TICKET_INFO_QUESTION_IDS.sheet;
const TICKET_END_ROUND_QUESTION_ID = LEGACY_TICKET_INFO_QUESTION_IDS.round;
const MEASUREMENT_METRICS = [
  { key: "waist", label: "ウエスト", color: "#c95f50", unit: "cm" },
  { key: "hip", label: "ヒップ", color: "#4e8c73", unit: "cm" },
  { key: "thighRight", label: "太もも右", color: "#c78a2c", unit: "cm" },
  { key: "thighLeft", label: "太もも左", color: "#6e7fba", unit: "cm" },
];
const MEASUREMENT_PERIOD_OPTIONS = [
  { value: "1m", label: "過去1ヶ月" },
  { value: "6m", label: "過去半年" },
  { value: "1y", label: "過去1年" },
  { value: "all", label: "全期間" },
];
const DEFAULT_MEASUREMENT_VISIBILITY = {
  waist: true,
  hip: true,
  thighRight: true,
  thighLeft: true,
  whr: true,
};

function parseLaunchRoute() {
  try {
    const params = new URLSearchParams(window.location.search);
    return {
      page: String(params.get("page") || "").trim(),
      postId: String(params.get("postId") || "").trim(),
    };
  } catch {
    return {
      page: "",
      postId: "",
    };
  }
}

const SESSION_CONCERN_CATEGORIES = [
  {
    id: "toilet",
    label: "【トイレ・デリケートなお悩み】",
    options: [
      "咳やくしゃみ、大笑いをした時に少し気になることがある",
      "ジャンプや運動、重いものを持った時に気になることがある",
      "急にトイレに行きたくなり、間に合うか不安になることがある",
      "以前よりトイレが近くなった気がする",
      "夜中にトイレで目が覚めることがある",
      "外出先でトイレの場所が気になりやすい",
      "トイレのあとも、すっきりしない感じがある",
      "尿の出が弱い、出にくいと感じることがある",
      "トイレに時間がかかることがある",
      "ナプキンやパッドが手放せず不安を感じることがある",
    ],
  },
  {
    id: "belly",
    label: "【お腹まわり・便通のお悩み】",
    options: [
      "便秘が気になる",
      "すっきり出にくいと感じることがある",
      "お腹に力を入れにくい感じがある",
      "下腹が張りやすい",
      "下腹ぽっこりが気になる",
      "お腹まわりの支えが弱くなった気がする",
      "インナーマッスルの衰えが気になる",
      "お腹まわりをすっきり整えたい",
      "お腹の奥に力が入りにくい感じがある",
      "体の内側から支えられていない感じがある",
    ],
  },
  {
    id: "pelvic",
    label: "【骨盤まわり・内側の筋力のお悩み】",
    options: [
      "骨盤まわりが不安定に感じる",
      "骨盤底筋をうまく使えていない気がする",
      "締める感覚がわかりにくい",
      "自分では鍛えにくい部分をケアしたい",
      "体の内側の支える力が弱くなった気がする",
      "体幹の弱さが気になる",
      "出産後から骨盤まわりの変化が気になる",
      "年齢とともに筋力の低下を感じる",
      "将来のために早めにケアしておきたい",
      "骨盤の底から支える感覚を取り戻したい",
    ],
  },
  {
    id: "posture",
    label: "【姿勢・体型のお悩み】",
    options: [
      "姿勢の崩れが気になる",
      "猫背が気になる",
      "反り腰が気になる",
      "立ち姿をきれいに見せたい",
      "歩き方や姿勢を整えたい",
      "下腹ぽっこりが気になる",
      "ヒップラインの変化が気になる",
      "ヒップアップしたい",
      "体のラインをすっきり整えたい",
      "無理なく体の土台から整えたい",
    ],
  },
  {
    id: "lower-body",
    label: "【腰まわり・下半身のお悩み】",
    options: [
      "腰まわりに負担を感じやすい",
      "股関節まわりが硬く感じる",
      "お尻の筋肉をうまく使えていない気がする",
      "太ももばかり疲れやすい",
      "長時間立っているとつらい",
      "歩くと疲れやすい",
      "階段の上り下りが気になる",
      "下半身の筋力低下が気になる",
      "下半身を安定させたい",
      "お尻や骨盤まわりをしっかり使えるようになりたい",
    ],
  },
  {
    id: "postpartum-aging",
    label: "【産後・年齢による変化】",
    options: [
      "出産後から体の変化が気になっている",
      "出産後から骨盤まわりが不安定に感じる",
      "出産後、お腹やお尻まわりが戻りにくいと感じる",
      "以前より体を支える力が弱くなった気がする",
      "年齢とともに変化を感じるようになった",
      "更年期以降、トイレや骨盤まわりの悩みが増えた",
      "今は大きな悩みはないが、予防として始めたい",
      "将来のために骨盤底筋ケアを取り入れたい",
      "出産後、咳や抱っこで気になることが増えた",
      "これから先の体の変化に備えて整えておきたい",
    ],
  },
  {
    id: "daily-life",
    label: "【日常生活で気になること】",
    options: [
      "外出や旅行の時に少し不安がある",
      "長時間の移動が気になる",
      "会議や授業中にトイレが気になることがある",
      "運動や趣味を思いきり楽しみにくい",
      "ジャンプやランニングを控えることがある",
      "重い荷物を持つ時に不安がある",
      "お子さまの抱っこなどで気になることがある",
      "夜ぐっすり眠りたい",
      "日常のちょっとした動作に不安がある",
      "トイレを気にせず過ごせる時間を増やしたい",
    ],
  },
];

const appState = {
  customer: normalizeCustomerProfile(loadLocal(CUSTOMER_KEY, { name: "", nameKana: "", historyMatchMode: "device" })),
  drafts: loadLocal(DRAFTS_KEY, {}),
  pendingSubmission: loadLocal(PENDING_KEY, null),
  ticketCardOverride: normalizeActiveTicketCardOverride(loadLocal(TICKET_CARD_OVERRIDE_KEY, null)),
  bijirisFavoritesByCustomer: loadLocal(BIJIRIS_FAVORITES_KEY, {}),
  bijirisReaderStateByCustomer: loadLocal(BIJIRIS_READER_STATE_KEY, {}),
  surveys: [],
  history: [],
  measurements: [],
  bijirisPosts: [],
  publicInfo: {
    dataPolicyText: "",
    requireConsent: true,
    consentText: "",
    version: "",
    pushAppId: "",
  },
  selectedSurveyId: "",
  installPrompt: null,
  panelMode: "form",
  confirmPayload: null,
  editingResponseId: "",
  lastSubmittedResponse: null,
  lastSubmissionWasEdit: false,
  historySurveyId: "",
  historyResponseId: "",
  historyLoading: false,
  historyLoadError: "",
  selectedBijirisPostId: "",
  selectedBijirisCategory: "all",
  bijirisSearchQuery: "",
  showFavoriteBijirisOnly: false,
  showReadLaterBijirisOnly: false,
  bijirisLoading: false,
  bijirisLoadError: "",
  concernCategoryByQuestion: {},
  selectedMeasurementPeriod: "6m",
  measurementMetricVisibility: { ...DEFAULT_MEASUREMENT_VISIBILITY },
  selectedMeasurementPhotoComparisonId: "",
  pushSupported: false,
  pushEnabled: false,
  pushActualEnabled: false,
  pushBusy: false,
  pushInitialized: false,
  pushServerSignature: "",
  launchRoute: parseLaunchRoute(),
};

const api = window.MayumiSurveyApi;
const toast = document.querySelector("#toast");
const surveyList = document.querySelector("#surveyList");
const answerPanel = document.querySelector("#answerPanel");
const historyList = document.querySelector("#historyList");
const measurementPanel = document.querySelector("#measurementPanel");
const bijirisPanel = document.querySelector("#bijirisPanel");
const homeTicketStatus = document.querySelector("#homeTicketStatus");
const customerLoginForm = document.querySelector("#customerLoginForm");
const customerForm = document.querySelector("#customerForm");
const customerMemberInfo = document.querySelector("#customerMemberInfo");
const appUpdateButton = document.querySelector("#appUpdateButton");
const installButton = document.querySelector("#installButton");
const registrationLead = document.querySelector("#registrationLead");
const customerRegisterButton = document.querySelector("#customerRegisterButton");
const recoverAccountButton = document.querySelector("#recoverAccountButton");
const bottomNav = document.querySelector("#bottomNav");
const measurementRefreshButton = document.querySelector("#measurementRefreshButton");
const bijirisRefreshButton = document.querySelector("#bijirisRefreshButton");
const pushStatusText = document.querySelector("#pushStatusText");
const pushHelpText = document.querySelector("#pushHelpText");
const pushToggleButton = document.querySelector("#pushToggleButton");
const bijirisNavBadge = document.querySelector("#bijirisNavBadge");

function loadLocal(key, fallback) {
  try {
    const value = localStorage.getItem(key);
    return value ? JSON.parse(value) : fallback;
  } catch {
    return fallback;
  }
}

function saveLocal(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Ignore localStorage quota errors.
  }
}

function removeLocal(key) {
  try {
    localStorage.removeItem(key);
  } catch {
    // Ignore storage errors.
  }
}

function cloneData(value) {
  return JSON.parse(JSON.stringify(value));
}

function getConfiguredPushAppId() {
  return String(appState.publicInfo?.pushAppId || window.MAYUMI_ONESIGNAL_APP_ID || DEFAULT_ONESIGNAL_APP_ID || "").trim();
}

function getStoredPushPreference() {
  try {
    return localStorage.getItem(PUSH_ENABLED_STORAGE_KEY) === "true";
  } catch {
    return false;
  }
}

function savePushPreference(enabled) {
  try {
    localStorage.setItem(PUSH_ENABLED_STORAGE_KEY, enabled ? "true" : "false");
  } catch {
    // Ignore storage errors.
  }
}

function getStoredPushAppId() {
  try {
    return String(localStorage.getItem(PUSH_APP_ID_STORAGE_KEY) || "").trim();
  } catch {
    return "";
  }
}

function saveStoredPushAppId(appId) {
  try {
    localStorage.setItem(PUSH_APP_ID_STORAGE_KEY, String(appId || "").trim());
  } catch {
    // Ignore storage errors.
  }
}

function normalizeText(value) {
  return String(value ?? "").trim();
}

function normalizeKana(value) {
  return String(value ?? "").replace(/\s+/g, "").trim();
}

function normalizeMeasurementValue(value) {
  if (value === null || value === undefined || value === "") return "";
  const normalized = String(value).replace(/[^\d.-]/g, "");
  if (!normalized) return "";
  const parsed = Number(normalized);
  if (!Number.isFinite(parsed)) return "";
  return Math.round(parsed * 10) / 10;
}

function normalizeMeasurementTargets(value) {
  const targets = {
    waist: normalizeMeasurementValue(value?.waist),
    hip: normalizeMeasurementValue(value?.hip),
    thighRight: normalizeMeasurementValue(value?.thighRight),
    thighLeft: normalizeMeasurementValue(value?.thighLeft),
  };
  return Object.values(targets).some((item) => item !== "") ? targets : null;
}

function isPushFeatureSupported() {
  return window.isSecureContext && "Notification" in window && "serviceWorker" in navigator;
}

function normalizeCustomerProfile(value) {
  return {
    name: normalizeText(value?.name),
    memberNumber: normalizeText(value?.memberNumber).toUpperCase(),
    nameKana: normalizeKana(value?.nameKana),
    historyMatchMode: value?.historyMatchMode === "name" ? "name" : "device",
    measurementTargets: normalizeMeasurementTargets(value?.measurementTargets),
  };
}

function normalizeActiveTicketCardOverride(value) {
  if (!value || typeof value !== "object") return null;
  const directPlan = normalizeText(value.plan);
  const directSheetNumber = Math.floor(Number(value.sheetNumber) || 0);
  if (directPlan && directSheetNumber > 0) {
    return { plan: directPlan, sheetNumber: directSheetNumber };
  }

  const legacyEntry = Object.entries(value).find(([plan, override]) => {
    const normalizedPlan = normalizeText(plan);
    const sheetNumber = Math.floor(
      Number(typeof override === "object" && override ? override.sheetNumber : override) || 0,
    );
    return normalizedPlan && sheetNumber > 0;
  });
  if (!legacyEntry) return null;
  return {
    plan: normalizeText(legacyEntry[0]),
    sheetNumber: Math.floor(
      Number(
        typeof legacyEntry[1] === "object" && legacyEntry[1]
          ? legacyEntry[1].sheetNumber
          : legacyEntry[1],
      ) || 0,
    ),
  };
}

function saveActiveTicketCardOverride() {
  saveLocal(TICKET_CARD_OVERRIDE_KEY, appState.ticketCardOverride);
}

function getActiveTicketCardOverride() {
  return normalizeActiveTicketCardOverride(appState.ticketCardOverride);
}

function setActiveTicketCardOverride(plan, sheetNumber) {
  const normalizedPlan = normalizeText(plan);
  const normalizedSheetNumber = Math.floor(Number(sheetNumber) || 0);
  if (!normalizedPlan || normalizedSheetNumber <= 0) return;
  appState.ticketCardOverride = {
    plan: normalizedPlan,
    sheetNumber: normalizedSheetNumber,
  };
  saveActiveTicketCardOverride();
}

function clearActiveTicketCardOverride() {
  if (!appState.ticketCardOverride) return;
  appState.ticketCardOverride = null;
  saveActiveTicketCardOverride();
}

function syncActiveTicketCardOverrideWithHistory() {
  const activeOverride = getActiveTicketCardOverride();
  if (!activeOverride) return;
  const latestResponse = getLatestTicketResponseByPlan(activeOverride.plan);
  if (!latestResponse) return;
  const ticketMap = new Map(getResponseTicketInfo(latestResponse).map((item) => [item.label, item.value]));
  const latestSheetNumber = parseTicketSheet(ticketMap.get("何枚目") || "");
  if (latestSheetNumber < activeOverride.sheetNumber) return;
  clearActiveTicketCardOverride();
}

function normalizeServerActiveTicketCard(value) {
  if (!value || typeof value !== "object") return null;
  const plan = normalizeText(value.plan);
  const sheetNumber = Math.floor(Number(value.sheetNumber) || parseTicketSheet(value.sheetLabel || ""));
  const round = Math.max(0, Math.floor(Number(value.round) || parseTicketStep(value.roundLabel || "")));
  if (!plan || sheetNumber <= 0) return null;
  return {
    plan,
    sheetNumber,
    round,
  };
}

function normalizeServerCustomerProfile(value) {
  return {
    name: normalizeText(value?.name),
    memberNumber: normalizeText(value?.memberNumber).toUpperCase(),
    nameKana: normalizeKana(value?.nameKana),
    activeTicketCard: normalizeServerActiveTicketCard(value?.activeTicketCard),
    measurementTargets: normalizeMeasurementTargets(value?.measurementTargets),
    pushStatus: normalizePushStatus(value?.pushStatus),
  };
}

function normalizePushPermission(value) {
  const normalized = normalizeText(value).toLowerCase();
  return ["granted", "denied", "default", "unsupported"].includes(normalized) ? normalized : "";
}

function normalizePushStatus(value) {
  if (!value || typeof value !== "object") return null;
  const hasEnabled = Object.prototype.hasOwnProperty.call(value, "enabled");
  const hasSupported = Object.prototype.hasOwnProperty.call(value, "supported");
  const permission = normalizePushPermission(value.permission);
  if (!hasEnabled && !hasSupported && !permission) return null;
  return {
    enabled: value.enabled === true,
    supported: value.supported === true,
    permission: permission || (value.supported === false ? "unsupported" : ""),
    updatedAt: normalizeText(value.updatedAt),
  };
}

function getCurrentPushStatusPayload(overrides = {}) {
  const supported = overrides.supported === true
    ? true
    : overrides.supported === false
      ? false
      : appState.pushSupported === true;
  const enabled = overrides.enabled === true;
  const permission = normalizePushPermission(
    overrides.permission !== undefined
      ? overrides.permission
      : supported
        ? Notification.permission
        : "unsupported",
  ) || (supported ? "default" : "unsupported");
  return {
    enabled,
    supported,
    permission,
  };
}

async function syncPushStatusToServer(pushStatus) {
  if (!hasCustomerSession()) return;
  const normalized = normalizePushStatus(pushStatus);
  if (!normalized) return;
  const signature = JSON.stringify({
    name: appState.customer.name,
    nameKana: appState.customer.nameKana,
    enabled: normalized.enabled,
    supported: normalized.supported,
    permission: normalized.permission,
  });
  if (signature === appState.pushServerSignature) return;
  try {
    const result = await api.request("/api/public/customer-profile/push", {
      method: "POST",
      body: {
        customer: appState.customer,
        pushStatus: normalized,
      },
    });
    appState.pushServerSignature = signature;
    if (result?.customerProfile) {
      syncCustomerProfileFromServer(result.customerProfile);
    }
  } catch (error) {
    reportClientError("customer.push.sync", error);
  }
}

function mergeCustomerProfile(baseProfile, serverProfile) {
  const base = normalizeCustomerProfile(baseProfile);
  const server = normalizeServerCustomerProfile(serverProfile);
  if (!server.name && !server.nameKana) return base;
  return normalizeCustomerProfile({
    name: server.name || base.name,
    memberNumber: server.memberNumber || base.memberNumber,
    nameKana: server.nameKana || base.nameKana,
    historyMatchMode: base.historyMatchMode,
    measurementTargets: server.measurementTargets || base.measurementTargets,
  });
}

function syncCustomerProfileFromServer(serverProfile) {
  const merged = mergeCustomerProfile(appState.customer, serverProfile);
  if (
    merged.name === appState.customer.name &&
    merged.memberNumber === appState.customer.memberNumber &&
    merged.nameKana === appState.customer.nameKana &&
    merged.historyMatchMode === appState.customer.historyMatchMode &&
    JSON.stringify(merged.measurementTargets || null) === JSON.stringify(appState.customer.measurementTargets || null)
  ) {
    return false;
  }
  appState.customer = merged;
  saveLocal(CUSTOMER_KEY, appState.customer);
  syncCustomerForms();
  return true;
}

function syncActiveTicketCardOverrideFromServer(serverProfile) {
  const server = normalizeServerCustomerProfile(serverProfile);
  const activeTicketCard = server.activeTicketCard;
  if (!activeTicketCard) return false;

  const currentOverride = getActiveTicketCardOverride();
  if (activeTicketCard.round <= 0) {
    if (!currentOverride) {
      setActiveTicketCardOverride(activeTicketCard.plan, activeTicketCard.sheetNumber);
      return true;
    }
    return false;
  }

  if (
    currentOverride &&
    currentOverride.plan === activeTicketCard.plan &&
    currentOverride.sheetNumber <= activeTicketCard.sheetNumber
  ) {
    clearActiveTicketCardOverride();
    return true;
  }
  return false;
}

function normalizeMeasurementRecord(value) {
  const measuredAt = normalizeText(value?.measuredAt);
  const createdAt = normalizeText(value?.createdAt);
  const updatedAt = normalizeText(value?.updatedAt);
  const waist = normalizeMeasurementValue(value?.waist);
  const hip = normalizeMeasurementValue(value?.hip);
  const thighRight = normalizeMeasurementValue(value?.thighRight);
  const thighLeft = normalizeMeasurementValue(value?.thighLeft);
  return {
    id: normalizeText(value?.id),
    customerName: normalizeText(value?.customerName),
    memberNumber: normalizeText(value?.memberNumber).toUpperCase(),
    measuredAt,
    waist,
    hip,
    thighRight,
    thighLeft,
    whr: hip === "" || !(Number(hip) > 0) || waist === "" ? "" : Math.round((waist / hip) * 1000) / 1000,
    createdAt,
    updatedAt,
    target: normalizeMeasurementTargets(value?.target),
  };
}

function buildHistorySearchParams(profile, options = {}) {
  const params = new URLSearchParams();
  params.set("name", normalizeText(profile?.name));
  const nameKana = normalizeKana(profile?.nameKana);
  if (nameKana) {
    params.set("nameKana", nameKana);
  }
  if (options.recovery || profile?.historyMatchMode === "name") {
    params.set("recoverByName", "1");
  }
  return params;
}

function isKatakanaName(value) {
  return /^[ァ-ヶー・ヴ　]+$/.test(String(value || ""));
}

function hasCustomerSession() {
  return Boolean(appState.customer.name && appState.customer.nameKana);
}

function getCustomerDisplayName() {
  if (!hasCustomerSession()) return "";
  return `${appState.customer.name}（${appState.customer.nameKana}）`;
}

function getCustomerMemberNumber() {
  return normalizeText(appState.customer.memberNumber).toUpperCase();
}

function isStandaloneApp() {
  return (
    window.matchMedia?.("(display-mode: standalone)")?.matches ||
    window.navigator.standalone === true
  );
}

function canRegisterFromThisContext() {
  const host = window.location.hostname;
  return isStandaloneApp() || host === "localhost" || host === "127.0.0.1";
}

function getBijirisNotificationCount() {
  if (!hasCustomerSession()) return 0;
  return getUnreadBijirisPosts().length;
}

function updateBijirisNotificationUi() {
  const count = getBijirisNotificationCount();
  if (bijirisNavBadge) {
    bijirisNavBadge.hidden = count <= 0;
    bijirisNavBadge.textContent = count > 99 ? "99+" : String(count);
  }
  if ("setAppBadge" in navigator && "clearAppBadge" in navigator) {
    if (count > 0) {
      navigator.setAppBadge(count).catch(() => {});
    } else {
      navigator.clearAppBadge().catch(() => {});
    }
  }
}

function updatePushUi() {
  appState.pushSupported = isPushFeatureSupported() && Boolean(getConfiguredPushAppId());
  if (pushToggleButton) {
    pushToggleButton.disabled = !appState.pushSupported || !hasCustomerSession();
    if (!hasCustomerSession()) {
      pushToggleButton.textContent = "ログイン後に設定";
    } else if (!appState.pushSupported) {
      pushToggleButton.textContent = "この端末では利用できません";
    } else {
      pushToggleButton.textContent = appState.pushEnabled ? "通知オフにする" : "通知オンにする";
    }
  }
  if (pushStatusText) {
    if (!hasCustomerSession()) {
      pushStatusText.textContent = "会員登録後に通知設定を利用できます。";
    } else if (!appState.pushSupported) {
      pushStatusText.textContent = "この端末または現在の表示方法では通知設定を利用できません。";
    } else if (Notification.permission === "denied") {
      pushStatusText.textContent = "端末側で通知が拒否されています。ブラウザまたはホーム画面アプリの通知設定を確認してください。";
    } else {
      pushStatusText.textContent = appState.pushEnabled ? "現在の設定: 通知オン" : "現在の設定: 通知オフ";
    }
  }
  if (pushHelpText) {
    pushHelpText.textContent = appState.pushSupported
      ? "豆知識の追加や更新を通知で受け取れます。ホーム画面に追加したアプリでの利用をおすすめします。"
      : "通知は対応端末とホーム画面に追加したアプリでの利用が必要です。";
  }
}

function getOneSignalInstance() {
  return window.OneSignalRef || null;
}

let oneSignalInitPromise = null;

function getPushWorkerScopePath() {
  try {
    return new URL("./push/", window.location.href).pathname;
  } catch {
    return "/push/";
  }
}

async function clearScopedPushRegistrations() {
  if (!("serviceWorker" in navigator)) return false;
  const registrations = await navigator.serviceWorker.getRegistrations();
  const targetScopePath = getPushWorkerScopePath();
  let cleared = false;
  for (const registration of registrations) {
    let scopePath = "";
    try {
      scopePath = new URL(registration.scope).pathname;
    } catch {
      scopePath = "";
    }
    if (scopePath !== targetScopePath && !scopePath.endsWith(targetScopePath)) continue;
    try {
      const subscription = await registration.pushManager?.getSubscription?.();
      if (subscription) {
        await subscription.unsubscribe().catch(() => {});
        cleared = true;
      }
    } catch {
      // Ignore unsubscribe failures and continue unregister.
    }
    try {
      const unregistered = await registration.unregister();
      cleared = cleared || unregistered === true;
    } catch {
      // Ignore unregister failures.
    }
  }
  return cleared;
}

async function migratePushAppIfNeeded() {
  const currentAppId = getConfiguredPushAppId();
  if (!currentAppId || !isPushFeatureSupported()) return { changed: false, desiredEnabled: getStoredPushPreference() };
  const storedAppId = getStoredPushAppId();
  if (storedAppId === currentAppId) return { changed: false, desiredEnabled: getStoredPushPreference() };
  const desiredEnabled = getStoredPushPreference();
  await clearScopedPushRegistrations();
  window.OneSignalRef = null;
  window.OneSignalDeferred = [];
  oneSignalInitPromise = null;
  appState.pushInitialized = false;
  appState.pushEnabled = false;
  appState.pushActualEnabled = false;
  appState.pushServerSignature = "";
  saveStoredPushAppId(currentAppId);
  return { changed: true, desiredEnabled };
}

async function loadOneSignalSdk() {
  if (window.OneSignalRef) return window.OneSignalRef;
  if (oneSignalInitPromise) return oneSignalInitPromise;
  const appId = getConfiguredPushAppId();
  if (!appId || !isPushFeatureSupported()) return null;

  oneSignalInitPromise = new Promise((resolve, reject) => {
    const existingScript = document.querySelector('script[data-onesignal-sdk="true"]');
    if (!existingScript) {
      const script = document.createElement("script");
      script.src = "https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.page.js";
      script.async = true;
      script.dataset.onesignalSdk = "true";
      script.addEventListener("error", () => reject(new Error("通知SDKを読み込めませんでした。")));
      document.head.appendChild(script);
    }

    window.OneSignalDeferred = window.OneSignalDeferred || [];
    window.OneSignalDeferred.push(async (OneSignal) => {
      try {
        window.OneSignalRef = OneSignal;
        const appBase = new URL("./", window.location.href).pathname;
        const workerBase = `${appBase}push/`;
        await OneSignal.init({
          appId,
          serviceWorkerPath: `${workerBase}OneSignalSDKWorker.js`,
          serviceWorkerParam: { scope: workerBase },
          allowLocalhostAsSecureOrigin: true,
        });
        appState.pushInitialized = true;
        OneSignal.Notifications?.addEventListener?.("permissionChange", () => {
          void syncPushStateFromOneSignal();
        });
        OneSignal.User?.PushSubscription?.addEventListener?.("change", () => {
          void syncPushStateFromOneSignal();
        });
        resolve(OneSignal);
      } catch (error) {
        reject(error);
      }
    });
  }).catch((error) => {
    oneSignalInitPromise = null;
    appState.pushInitialized = false;
    reportClientError("customer.push.sdk", error);
    return null;
  });

  return oneSignalInitPromise;
}

async function syncPushStateFromOneSignal(options = {}) {
  const updateDesired = options.updateDesired !== undefined ? options.updateDesired : !appState.pushBusy;
  const OneSignal = getOneSignalInstance() || await loadOneSignalSdk();
  const pushSubscription = OneSignal?.User?.PushSubscription;
  const permission = OneSignal?.Notifications?.permission;
  const permissionGranted = permission && typeof permission.then === "function" ? await permission : permission;
  const enabled = Boolean(permissionGranted) && Boolean(pushSubscription?.optedIn);
  appState.pushActualEnabled = enabled;
  if (updateDesired) {
    appState.pushEnabled = enabled;
    savePushPreference(enabled);
  }
  updatePushUi();
  await syncPushStatusToServer(
    getCurrentPushStatusPayload({
      enabled,
      supported: appState.pushSupported,
      permission:
        permissionGranted === true
          ? "granted"
          : permissionGranted === false
            ? Notification.permission || "default"
            : permissionGranted || Notification.permission || "default",
    }),
  );
  return enabled;
}

async function initializePushNotifications() {
  const migration = await migratePushAppIfNeeded();
  appState.pushSupported = isPushFeatureSupported() && Boolean(getConfiguredPushAppId());
  appState.pushEnabled = getStoredPushPreference();
  appState.pushActualEnabled = appState.pushEnabled;
  updatePushUi();
  if (!appState.pushSupported) {
    await syncPushStatusToServer(getCurrentPushStatusPayload({ enabled: false, supported: false, permission: "unsupported" }));
    return;
  }
  if (migration.changed && migration.desiredEnabled && Notification.permission === "granted") {
    const OneSignal = await loadOneSignalSdk();
    const pushSubscription = OneSignal?.User?.PushSubscription;
    if (pushSubscription && !pushSubscription.optedIn && typeof pushSubscription.optIn === "function") {
      await pushSubscription.optIn().catch(() => {});
    }
  }
  await syncPushStateFromOneSignal();
  if (migration.changed) {
    showToast("通知設定を更新しました。別アプリの通知は届かないよう再設定しています。");
  }
}

async function flushPushToggleQueue() {
  if (appState.pushBusy || !hasCustomerSession() || !appState.pushSupported) return;
  appState.pushBusy = true;
  const startEnabled = appState.pushActualEnabled;
  let lastError = null;

  try {
    const OneSignal = await loadOneSignalSdk();
    if (!OneSignal) throw new Error("通知設定を初期化できませんでした。");

    while (appState.pushActualEnabled !== appState.pushEnabled) {
      const targetEnabled = appState.pushEnabled;
      const pushSubscription = OneSignal.User?.PushSubscription;

      if (!targetEnabled) {
        if (pushSubscription?.optedIn && typeof pushSubscription.optOut === "function") {
          await pushSubscription.optOut();
        }
      } else {
        const permissionResult = await OneSignal.Notifications?.requestPermission?.();
        const granted = permissionResult === true || Notification.permission === "granted";
        if (!granted) {
          appState.pushEnabled = false;
          savePushPreference(false);
          await syncPushStateFromOneSignal({ updateDesired: false });
          throw new Error("通知の許可が必要です。");
        }
        if (pushSubscription && !pushSubscription.optedIn && typeof pushSubscription.optIn === "function") {
          await pushSubscription.optIn();
        }
      }

      await syncPushStateFromOneSignal({ updateDesired: false });
    }
  } catch (error) {
    lastError = error;
    appState.pushEnabled = appState.pushActualEnabled;
    savePushPreference(appState.pushEnabled);
    reportClientError("customer.push.toggle", error);
    showToast(error.message || "通知設定を更新できませんでした。");
  } finally {
    appState.pushBusy = false;
    updatePushUi();
  }

  if (!lastError && startEnabled !== appState.pushActualEnabled) {
    showToast(appState.pushActualEnabled ? "通知をオンにしました。" : "通知をオフにしました。");
  }

  if (appState.pushActualEnabled !== appState.pushEnabled) {
    void flushPushToggleQueue();
  }
}

async function togglePushNotifications() {
  if (!hasCustomerSession()) return;
  if (!appState.pushSupported) {
    showToast("この端末では通知を利用できません。");
    return;
  }
  const nextEnabled = !appState.pushEnabled;
  appState.pushEnabled = nextEnabled;
  savePushPreference(nextEnabled);
  updatePushUi();
  void flushPushToggleQueue();
}

function clearLaunchRoute() {
  appState.launchRoute = { page: "", postId: "" };
  try {
    const url = new URL(window.location.href);
    url.searchParams.delete("page");
    url.searchParams.delete("postId");
    url.searchParams.delete("notice");
    window.history.replaceState({}, "", url.toString());
  } catch {
    // Ignore URL rewrite errors.
  }
}

function applyLaunchRouteIfPossible() {
  const targetPage = normalizeText(appState.launchRoute?.page);
  const postId = normalizeText(appState.launchRoute?.postId);
  if (targetPage !== "bijiris" || !hasCustomerSession()) return false;
  if (postId) {
    const post = appState.bijirisPosts.find((item) => item.id === postId);
    if (!post) return false;
    appState.selectedBijirisPostId = postId;
    recordBijirisView(postId);
  }
  setPage("bijiris");
  clearLaunchRoute();
  return true;
}

function renderRegistrationGuide() {
  const canRegister = canRegisterFromThisContext();
  if (registrationLead) {
    registrationLead.textContent = canRegister
      ? "このアプリから会員登録してください。登録後は回答履歴もそのまま確認できます。"
      : "ブラウザでは初回会員登録を行わず、ホーム画面に追加したアプリから会員登録してください。";
  }
  if (customerRegisterButton) {
    customerRegisterButton.disabled = !canRegister;
    customerRegisterButton.textContent = canRegister ? "会員登録する" : "ホーム画面アプリで会員登録";
  }
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatDate(value) {
  return new Intl.DateTimeFormat("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date(value));
}

function formatDateOnly(value) {
  return new Intl.DateTimeFormat("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(new Date(value));
}

function roundMeasurementDelta(value) {
  if (!Number.isFinite(Number(value))) return "";
  return Math.round(Number(value) * 10) / 10;
}

function formatMeasurementValue(value, unit = "cm") {
  const normalized = normalizeMeasurementValue(value);
  if (normalized === "") return "-";
  return `${normalized.toFixed(1)}${unit}`;
}

function formatMeasurementDelta(value, unit = "cm") {
  if (value === "" || value === null || value === undefined || !Number.isFinite(Number(value))) {
    return '<span class="delta-badge neutral">-</span>';
  }
  const normalized = roundMeasurementDelta(value);
  const className = normalized > 0 ? "increase" : normalized < 0 ? "decrease" : "neutral";
  const sign = normalized > 0 ? "+" : normalized < 0 ? "-" : "+/-";
  const text =
    normalized === 0
      ? `0.0${unit}`
      : `${sign}${Math.abs(normalized).toFixed(1)}${unit}`;
  return `<span class="delta-badge ${className}">${escapeHtml(text)}</span>`;
}

function formatWhr(value) {
  if (value === "" || value === null || value === undefined || !Number.isFinite(Number(value))) return "-";
  return Number(value).toFixed(3);
}

function formatPhotoCapturedAt(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return formatDateOnly(date.toISOString());
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => resolve(String(reader.result || "")));
    reader.addEventListener("error", () => reject(new Error("写真を読み込めませんでした。")));
    reader.readAsDataURL(file);
  });
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.addEventListener("load", () => resolve(image));
    image.addEventListener("error", () =>
      reject(new Error("JPEG、PNG、WEBP の写真を選択してください。")),
    );
    image.src = src;
  });
}

async function preparePhotoFile(file) {
  if (!file.type.startsWith("image/")) {
    throw new Error("写真ファイルを選択してください。");
  }

  const source = await readFileAsDataUrl(file);
  const image = await loadImage(source);
  const scale = Math.min(1, PHOTO_MAX_SIZE / Math.max(image.naturalWidth, image.naturalHeight));
  const width = Math.max(1, Math.round(image.naturalWidth * scale));
  const height = Math.max(1, Math.round(image.naturalHeight * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  context.drawImage(image, 0, 0, width, height);

  return {
    name: file.name || "photo.jpg",
    type: "image/jpeg",
    capturedAt: Number.isFinite(file.lastModified) && file.lastModified > 0
      ? new Date(file.lastModified).toISOString()
      : new Date().toISOString(),
    dataUrl: canvas.toDataURL("image/jpeg", PHOTO_JPEG_QUALITY),
  };
}

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("show");
  setTimeout(() => toast.classList.remove("show"), 2600);
}

function reportClientError(source, error, detail = {}) {
  void api.logError?.(source, error, detail);
}

function getFallbackSurveys() {
  const makeSurveys = window.MayumiDefaultSurveys;
  return typeof makeSurveys === "function"
    ? makeSurveys(new Date().toISOString()).filter((survey) => survey.status === "published")
    : [];
}

function isTicketEndSurvey(survey) {
  return survey?.id === TICKET_END_SURVEY_ID;
}

function normalizeDraft(draft) {
  const updatedAt = String(draft?.updatedAt || "").trim();
  return {
    values: typeof draft?.values === "object" && draft.values ? draft.values : {},
    photos: typeof draft?.photos === "object" && draft.photos ? draft.photos : {},
    updatedAt,
  };
}

function getSurveyDraft(surveyId) {
  return normalizeDraft(appState.drafts[surveyId]);
}

function setSurveyDraft(surveyId, draft) {
  const normalized = normalizeDraft(draft);
  normalized.updatedAt = new Date().toISOString();
  appState.drafts[surveyId] = normalized;
  saveLocal(DRAFTS_KEY, appState.drafts);
}

function clearSurveyDraft(surveyId) {
  delete appState.drafts[surveyId];
  saveLocal(DRAFTS_KEY, appState.drafts);
}

function getDraftValue(surveyId, questionId) {
  const draft = getSurveyDraft(surveyId);
  return draft.values[questionId];
}

function getDraftPhotos(surveyId, questionId) {
  const draft = getSurveyDraft(surveyId);
  return Array.isArray(draft.photos[questionId]) ? draft.photos[questionId] : [];
}

function updateDraftValue(surveyId, questionId, value) {
  const draft = getSurveyDraft(surveyId);
  draft.values[questionId] = value;
  setSurveyDraft(surveyId, draft);
}

function updateDraftPhotos(surveyId, questionId, files) {
  const draft = getSurveyDraft(surveyId);
  draft.photos[questionId] = Array.isArray(files) ? files : [];
  setSurveyDraft(surveyId, draft);
}

function clearPendingSubmission() {
  appState.pendingSubmission = null;
  removeLocal(PENDING_KEY);
}

function storePendingSubmission(payload) {
  appState.pendingSubmission = payload;
  saveLocal(PENDING_KEY, payload);
}

function getAgreementDraftKey(surveyId) {
  return `__agreement_${surveyId}`;
}

function isAgreementAccepted(surveyId) {
  return Boolean(getDraftValue(surveyId, getAgreementDraftKey(surveyId)));
}

function setAgreementAccepted(surveyId, accepted) {
  updateDraftValue(surveyId, getAgreementDraftKey(surveyId), Boolean(accepted));
}

function makeValidationError(questionId, message) {
  const error = new Error(message);
  error.questionId = questionId;
  return error;
}

function scrollToQuestion(questionId) {
  if (!questionId) return;
  requestAnimationFrame(() => {
    const node = document.querySelector(`[data-question-wrap="${questionId}"]`);
    if (node) node.scrollIntoView({ behavior: "smooth", block: "center" });
  });
}

function hasDraftContent(draft) {
  const values = Object.values(draft?.values || {}).some((value) => {
    if (Array.isArray(value)) return value.some((item) => normalizeText(item));
    if (typeof value === "boolean") return value;
    return Boolean(normalizeText(value));
  });
  const photos = Object.values(draft?.photos || {}).some((files) => Array.isArray(files) && files.length);
  return values || photos;
}

function getDraftSavedAt(surveyId) {
  const updatedAt = getSurveyDraft(surveyId).updatedAt;
  if (!updatedAt) return "";
  const date = new Date(updatedAt);
  if (Number.isNaN(date.getTime())) return "";
  return date.toISOString();
}

function getSelectedSurvey() {
  return appState.surveys.find((survey) => survey.id === appState.selectedSurveyId) || null;
}

function getVisibleSurveyIdSet() {
  return new Set(
    appState.surveys
      .map((survey) => normalizeText(survey?.id))
      .filter(Boolean),
  );
}

function getVisibleHistoryResponses() {
  const visibleSurveyIds = getVisibleSurveyIdSet();
  if (!visibleSurveyIds.size) return appState.history.slice();
  return appState.history.filter((response) => visibleSurveyIds.has(normalizeText(response?.surveyId)));
}

function isSessionSurvey(survey) {
  return survey?.id === SESSION_SURVEY_ID;
}

function getSessionTypeSelection(surveyId) {
  return normalizeText(getDraftValue(surveyId, SESSION_TYPE_QUESTION_ID));
}

function getSessionTicketPlanSelection(surveyId) {
  return normalizeText(getDraftValue(surveyId, SESSION_TICKET_PLAN_QUESTION_ID));
}

function getSessionTicketSheetSelection(surveyId) {
  return normalizeText(getDraftValue(surveyId, SESSION_TICKET_SHEET_QUESTION_ID));
}

function getSessionTicketRoundSelection(surveyId) {
  return normalizeText(getDraftValue(surveyId, SESSION_TICKET_ROUND_QUESTION_ID));
}

function getLatestTicketResponseByPlan(ticketPlan) {
  const normalizedPlan = normalizeText(ticketPlan);
  if (!normalizedPlan) return null;
  return (
    getVisibleHistoryResponses()
      .filter((response) => {
        const ticketInfo = getResponseTicketInfo(response);
        return ticketInfo.some((item) => item.label === "回数券" && item.value === normalizedPlan);
      })
      .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))[0] || null
  );
}

function getNextTicketSheetLabelForPlan(ticketPlan) {
  const normalizedPlan = normalizeText(ticketPlan);
  if (!normalizedPlan) return "";
  const latestResponse = getLatestTicketResponseByPlan(normalizedPlan);
  if (!latestResponse) return "1枚目";

  const ticketMap = new Map(getResponseTicketInfo(latestResponse).map((item) => [item.label, item.value]));
  const currentSheet = normalizeText(ticketMap.get("何枚目") || "");
  const currentRound = parseTicketStep(ticketMap.get("何回目") || "");
  const ticketCount = parseTicketCount(normalizedPlan);

  if (!currentSheet) return "1枚目";
  if (ticketCount && currentRound >= ticketCount) {
    return getNextTicketSheetLabel(currentSheet) || currentSheet;
  }
  return currentSheet;
}

function resolveCurrentSessionTicketSheetLabel(ticketPlan) {
  const normalizedPlan = normalizeText(ticketPlan);
  if (!normalizedPlan) return "";
  const activeOverride = getActiveTicketCardOverride();
  if (activeOverride?.plan === normalizedPlan && activeOverride.sheetNumber > 0) {
    return `${activeOverride.sheetNumber}枚目`;
  }
  return getNextTicketSheetLabelForPlan(normalizedPlan);
}

function ensureSessionTicketSheetSelection(surveyId) {
  const currentSheet = getSessionTicketSheetSelection(surveyId);
  if (currentSheet) return currentSheet;

  const ticketPlan = getSessionTicketPlanSelection(surveyId);
  if (!ticketPlan) return "";
  if (hasCustomerSession() && appState.historyLoading) return "";

  const resolvedSheet = resolveCurrentSessionTicketSheetLabel(ticketPlan) || "1枚目";
  updateDraftValue(surveyId, SESSION_TICKET_SHEET_QUESTION_ID, resolvedSheet);
  return resolvedSheet;
}

function getTicketResponsesByPlanAndSheet(ticketPlan, ticketSheetLabel, options = {}) {
  const normalizedPlan = normalizeText(ticketPlan);
  const normalizedSheet = normalizeText(ticketSheetLabel);
  const excludeResponseId = normalizeText(options.excludeResponseId);
  if (!normalizedPlan || !normalizedSheet) return [];
  return getVisibleHistoryResponses()
    .filter((response) => {
      if (excludeResponseId && response.id === excludeResponseId) return false;
      const ticketMap = new Map(getResponseTicketInfo(response).map((item) => [item.label, item.value]));
      return (
        normalizeText(ticketMap.get("回数券") || "") === normalizedPlan &&
        normalizeText(ticketMap.get("何枚目") || "") === normalizedSheet
      );
    })
    .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt));
}

function getExpectedSessionTicketRoundLabel(ticketPlan, ticketSheetLabel, options = {}) {
  const normalizedPlan = normalizeText(ticketPlan);
  const normalizedSheet = normalizeText(ticketSheetLabel);
  const ticketCount = parseTicketCount(normalizedPlan);
  if (!normalizedPlan || !normalizedSheet || !ticketCount) return "";

  const latestSameSheetResponse = getTicketResponsesByPlanAndSheet(normalizedPlan, normalizedSheet, options)[0] || null;
  if (latestSameSheetResponse) {
    const ticketMap = new Map(getResponseTicketInfo(latestSameSheetResponse).map((item) => [item.label, item.value]));
    const currentRound = parseTicketStep(ticketMap.get("何回目") || "");
    const expectedRound = Math.min(ticketCount, Math.max(1, currentRound + 1));
    return `${expectedRound}回目`;
  }

  const activeOverride = getActiveTicketCardOverride();
  if (activeOverride?.plan === normalizedPlan && `${activeOverride.sheetNumber}枚目` === normalizedSheet) {
    return "1回目";
  }

  return "1回目";
}

function getSessionTicketRoundMismatchMessage(ticketPlan, ticketSheetLabel, ticketRoundLabel, options = {}) {
  const selectedRound = normalizeText(ticketRoundLabel);
  if (!selectedRound) return "";
  const expectedRoundLabel = getExpectedSessionTicketRoundLabel(ticketPlan, ticketSheetLabel, options);
  if (!expectedRoundLabel || expectedRoundLabel === selectedRound) return "";
  return `スタンプカード情報と違います。正しくは ${expectedRoundLabel} だと思います。`;
}

function clearSessionSelections(surveyId) {
  const draft = getSurveyDraft(surveyId);
  delete draft.values[SESSION_TYPE_QUESTION_ID];
  delete draft.values[SESSION_TICKET_PLAN_QUESTION_ID];
  delete draft.values[SESSION_TICKET_SHEET_QUESTION_ID];
  delete draft.values[SESSION_TICKET_ROUND_QUESTION_ID];
  setSurveyDraft(surveyId, draft);
}

function getSessionTypeQuestion(survey) {
  return survey.questions.find((question) => question.id === SESSION_TYPE_QUESTION_ID);
}

function getSessionTicketPlanQuestion(survey) {
  return survey.questions.find((question) => question.id === SESSION_TICKET_PLAN_QUESTION_ID);
}

function getSessionTicketSheetQuestion(survey) {
  return survey.questions.find((question) => question.id === SESSION_TICKET_SHEET_QUESTION_ID);
}

function getSessionTicketRoundQuestion(survey) {
  return survey.questions.find((question) => question.id === SESSION_TICKET_ROUND_QUESTION_ID);
}

function clearSessionTicketProgressSelections(surveyId) {
  const draft = getSurveyDraft(surveyId);
  delete draft.values[SESSION_TICKET_SHEET_QUESTION_ID];
  delete draft.values[SESSION_TICKET_ROUND_QUESTION_ID];
  setSurveyDraft(surveyId, draft);
}

function getTicketEndCountSelection(surveyId) {
  return normalizeText(getDraftValue(surveyId, LEGACY_TICKET_INFO_QUESTION_IDS.size));
}

function getTicketEndSheetSelection(surveyId) {
  return normalizeText(getDraftValue(surveyId, LEGACY_TICKET_INFO_QUESTION_IDS.sheet));
}

function getTicketEndRoundSelection(surveyId) {
  return normalizeText(getDraftValue(surveyId, LEGACY_TICKET_INFO_QUESTION_IDS.round));
}

function clearTicketEndSelections(surveyId) {
  const draft = getSurveyDraft(surveyId);
  delete draft.values[LEGACY_TICKET_INFO_QUESTION_IDS.size];
  delete draft.values[LEGACY_TICKET_INFO_QUESTION_IDS.sheet];
  delete draft.values[LEGACY_TICKET_INFO_QUESTION_IDS.round];
  setSurveyDraft(surveyId, draft);
}

function getTicketEndCountQuestion(survey) {
  return survey.questions.find((question) => question.id === LEGACY_TICKET_INFO_QUESTION_IDS.size);
}

function getTicketEndSheetQuestion(survey) {
  return survey.questions.find((question) => question.id === LEGACY_TICKET_INFO_QUESTION_IDS.sheet);
}

function getTicketEndRoundQuestion(survey) {
  return survey.questions.find((question) => question.id === LEGACY_TICKET_INFO_QUESTION_IDS.round);
}

function getTicketRoundOptions(ticketCount) {
  const max = ticketCount === "10回券" ? 10 : ticketCount === "6回券" ? 6 : 0;
  return Array.from({ length: max }, (_, index) => `${index + 1}回目`);
}

function getSessionPhotoQuestionConfig(question) {
  if (!question) return null;
  if (question.id === "q_bijiris_session_ticket_photos") {
    return { maxFiles: 4, requiredCount: 4 };
  }
  if (
    SESSION_MONITOR_PHOTOS_QUESTION_IDS.includes(question.id) ||
    SESSION_TICKET_END_PHOTOS_QUESTION_IDS.includes(question.id)
  ) {
    return { maxFiles: 2, requiredCount: 2 };
  }
  return null;
}

function isLegacyBijirisSessionPhotoQuestion(question) {
  return Boolean(
    question &&
      [
        LEGACY_SESSION_MONITOR_PHOTOS_QUESTION_ID,
        LEGACY_SESSION_TICKET_END_PHOTOS_QUESTION_ID,
      ].includes(question.id),
  );
}

function isBijirisSessionFinalPhotoVisible(answerMap = {}) {
  const sessionType = normalizeText(answerMap[SESSION_TYPE_QUESTION_ID]?.[0]);
  const ticketPlan = normalizeText(answerMap[SESSION_TICKET_PLAN_QUESTION_ID]?.[0]);
  const ticketRound = normalizeText(answerMap[SESSION_TICKET_ROUND_QUESTION_ID]?.[0]);
  return (
    sessionType === "回数券" &&
    ((ticketPlan === "6回券" && ticketRound === "6回目") ||
      (ticketPlan === "10回券" && ticketRound === "10回目"))
  );
}

function isTicketEndLastPhotoQuestion(question) {
  return question?.id === "q_ticket_end_photo_last";
}

function getPhotoQuestionFileLimit(question) {
  const sessionPhotoQuestionConfig = getSessionPhotoQuestionConfig(question);
  if (sessionPhotoQuestionConfig) return sessionPhotoQuestionConfig.maxFiles;
  return PHOTO_FILE_LIMIT;
}

function getPhotoQuestionRequiredCount(question, surveyId = appState.selectedSurveyId) {
  const sessionPhotoQuestionConfig = getSessionPhotoQuestionConfig(question);
  if (sessionPhotoQuestionConfig) return sessionPhotoQuestionConfig.requiredCount;
  if (isTicketEndLastPhotoQuestion(question) && !getQuestionVisibilityConditions(question).length) {
    return isTicketEndLastPhotoRequired(surveyId) ? 1 : 0;
  }
  return question.required === false ? 0 : 1;
}

function isTicketEndLastPhotoRequired(surveyId) {
  const ticketCount = normalizeText(getDraftValue(surveyId, LEGACY_TICKET_INFO_QUESTION_IDS.size));
  const ticketRound = normalizeText(getDraftValue(surveyId, LEGACY_TICKET_INFO_QUESTION_IDS.round));
  return (
    (ticketCount === "6回券" && ticketRound === "6回目") ||
    (ticketCount === "10回券" && ticketRound === "10回目")
  );
}

function getQuestionVisibilityConditions(question) {
  const conditions = Array.isArray(question?.visibilityConditions)
    ? question.visibilityConditions
        .map((condition) => ({
          questionId: normalizeText(condition?.questionId),
          value: normalizeText(condition?.value),
        }))
        .filter((condition) => condition.questionId && condition.value)
    : [];
  if (conditions.length) return conditions;
  const fallbackQuestionId = normalizeText(question?.visibleWhen?.questionId);
  const fallbackValue = normalizeText(question?.visibleWhen?.value);
  return fallbackQuestionId && fallbackValue
    ? [{ questionId: fallbackQuestionId, value: fallbackValue }]
    : [];
}

function doesQuestionAffectVisibility(survey, questionId) {
  return survey.questions.some((question) =>
    getQuestionVisibilityConditions(question).some((condition) => condition.questionId === questionId),
  );
}

function isQuestionRequired(question, surveyId) {
  if (question?.type === "photo") {
    return getPhotoQuestionRequiredCount(question, surveyId) > 0;
  }
  return question.required !== false;
}

function getQuestionLabel(question, surveyId = appState.selectedSurveyId) {
  if (question.id === "q_ticket_end_photo_last") {
    const ticketRound = normalizeText(getDraftValue(surveyId, LEGACY_TICKET_INFO_QUESTION_IDS.round));
    if (ticketRound) return `計測写真(${ticketRound})`;
    const ticketCount = normalizeText(getDraftValue(surveyId, LEGACY_TICKET_INFO_QUESTION_IDS.size));
    if (ticketCount === "6回券") return "計測写真(6回目)";
    if (ticketCount === "10回券") return "計測写真(10回目)";
  }
  return question.label;
}

function getSubmissionQuestionVisibility(question, answerMap, surveyId = appState.selectedSurveyId) {
  if (question.id === SESSION_TYPE_QUESTION_ID) return true;
  if (question.id === SESSION_TICKET_PLAN_QUESTION_ID) {
    return getSessionTypeSelection(surveyId) === "回数券";
  }
  if (question.id === SESSION_TICKET_SHEET_QUESTION_ID) {
    return getSessionTypeSelection(surveyId) === "回数券" && Boolean(getSessionTicketPlanSelection(surveyId));
  }
  if (question.id === SESSION_TICKET_ROUND_QUESTION_ID) {
    return (
      getSessionTypeSelection(surveyId) === "回数券" &&
      Boolean(getSessionTicketPlanSelection(surveyId) && getSessionTicketSheetSelection(surveyId))
    );
  }
  return isQuestionVisible(question, answerMap, surveyId);
}

function getConcernCategoryStateKey(surveyId, questionId) {
  return `${surveyId}:${questionId}`;
}

function getConcernActiveCategory(surveyId, questionId, selectedOptions = []) {
  const key = getConcernCategoryStateKey(surveyId, questionId);
  const current = appState.concernCategoryByQuestion[key];
  if (SESSION_CONCERN_CATEGORIES.some((category) => category.id === current)) {
    return current;
  }
  const matched = SESSION_CONCERN_CATEGORIES.find((category) =>
    category.options.some((option) => selectedOptions.includes(option)),
  );
  return matched?.id || "";
}

function setConcernActiveCategory(surveyId, questionId, categoryId) {
  const key = getConcernCategoryStateKey(surveyId, questionId);
  appState.concernCategoryByQuestion[key] = categoryId;
}

function isInlineManagedTextareaQuestion(question) {
  return question?.id === SESSION_CONCERN_OTHER_QUESTION_ID || question?.id === SESSION_LIFE_CHANGES_OTHER_QUESTION_ID;
}

function renderQuestionLabel(question, index, surveyId) {
  const required = isQuestionRequired(question, surveyId);
  return `
    <span class="question-title-row">
      <span class="question-title-text">${index + 1}. ${escapeHtml(getQuestionLabel(question, surveyId))}</span>
      <span class="badge ${required ? "required" : "optional"}">${required ? "必須" : "任意"}</span>
    </span>
  `;
}

function renderInlineOtherTextarea(questionId, surveyId, label) {
  return `
    <div class="question-inline-textarea">
      <div class="question-inline-textarea-label">${escapeHtml(label)}</div>
      <textarea data-question-id="${questionId}">${escapeHtml(getDraftValue(surveyId, questionId) || "")}</textarea>
    </div>
  `;
}

function getSurveyAvailability(survey) {
  const nowTime = Date.now();
  const nearDeadlineMs = 3 * 24 * 60 * 60 * 1000;
  if (survey?.acceptingResponses === false) {
    return { open: false, label: "受付停止中", detail: "現在このアンケートは停止しています。", nearDeadline: false };
  }
  if (survey?.startAt && new Date(survey.startAt).getTime() > nowTime) {
    return {
      open: false,
      label: "受付開始前",
      detail: `開始予定: ${formatDate(survey.startAt)}`,
      nearDeadline: false,
    };
  }
  if (survey?.endAt && new Date(survey.endAt).getTime() < nowTime) {
    return {
      open: false,
      label: "受付終了",
      detail: `終了日時: ${formatDate(survey.endAt)}`,
      nearDeadline: false,
    };
  }
  const nearDeadline = survey?.endAt
    ? new Date(survey.endAt).getTime() - nowTime <= nearDeadlineMs
    : false;
  return {
    open: true,
    label: nearDeadline ? "まもなく終了" : "受付中",
    detail: survey?.endAt ? `受付終了: ${formatDate(survey.endAt)}` : "期限指定なし",
    nearDeadline,
  };
}

function buildDraftAnswerMap(survey, draft) {
  const map = {};
  survey.questions.forEach((question) => {
    if (question.type === "photo") {
      map[question.id] = getDraftPhotos(survey.id, question.id);
      return;
    }
    const raw = draft.values[question.id];
    if (Array.isArray(raw)) {
      map[question.id] = raw.map(normalizeText).filter(Boolean);
      return;
    }
    const value = normalizeText(raw);
    map[question.id] = value ? [value] : [];
  });
  return map;
}

function isQuestionVisible(question, answerMap, surveyId = appState.selectedSurveyId) {
  if (surveyId === SESSION_SURVEY_ID && isLegacyBijirisSessionPhotoQuestion(question)) {
    return isBijirisSessionFinalPhotoVisible(answerMap);
  }
  const conditions = getQuestionVisibilityConditions(question);
  if (conditions.length) {
    return conditions.every((condition) => {
      const values = Array.isArray(answerMap[condition.questionId]) ? answerMap[condition.questionId] : [];
      return values.includes(condition.value);
    });
  }
  if (isTicketEndLastPhotoQuestion(question)) {
    return isTicketEndLastPhotoRequired(surveyId);
  }
  return true;
}

function getVisibleQuestions(survey, draft) {
  const answerMap = buildDraftAnswerMap(survey, draft);
  return survey.questions.filter((question) => {
    if (
      question.id === SESSION_TYPE_QUESTION_ID ||
      question.id === SESSION_TICKET_PLAN_QUESTION_ID ||
      question.id === SESSION_TICKET_SHEET_QUESTION_ID ||
      question.id === SESSION_TICKET_ROUND_QUESTION_ID
    ) {
      return false;
    }
    return isQuestionVisible(question, answerMap, survey.id);
  });
}

function isQuestionAnswered(question, surveyId, draft) {
  if (question.type === "photo") {
    return getDraftPhotos(surveyId, question.id).length >= getPhotoQuestionRequiredCount(question, surveyId);
  }
  const raw = draft.values[question.id];
  if (Array.isArray(raw)) return raw.some((item) => normalizeText(item));
  return Boolean(normalizeText(raw));
}

function getProgress(survey, draft) {
  const visibleQuestions = getVisibleQuestions(survey, draft);
  const requiredQuestions = visibleQuestions.filter((question) => isQuestionRequired(question, survey.id));
  const answeredRequired = requiredQuestions.filter((question) =>
    isQuestionAnswered(question, survey.id, draft),
  ).length;
  return {
    total: visibleQuestions.length,
    requiredTotal: requiredQuestions.length,
    answeredRequired,
    percent: requiredQuestions.length
      ? Math.round((answeredRequired / requiredQuestions.length) * 100)
      : 100,
  };
}

function renderPendingNotice() {
  if (!appState.pendingSubmission) return "";
  return `
    <article class="retry-card">
      <strong>未送信の回答があります</strong>
      <div class="meta">${escapeHtml(appState.pendingSubmission.surveyTitle || "アンケート回答")} / ${formatDate(appState.pendingSubmission.savedAt)}</div>
      <p>通信エラーで保存確認ができていません。再送信してください。</p>
      <button class="primary-button" type="button" data-retry-pending>再送信する</button>
    </article>
  `;
}

function setPage(page) {
  if (page !== "login" && !hasCustomerSession()) {
    page = "login";
  }
  const navPage = page === "survey" ? "home" : page;
  document.querySelectorAll(".page").forEach((node) => {
    node.classList.toggle("active", node.id === `page-${page}`);
  });
  document.querySelectorAll(".nav-button").forEach((button) => {
    button.classList.toggle("active", button.dataset.page === navPage);
  });
  if (bottomNav) {
    bottomNav.hidden = !hasCustomerSession();
  }
  updatePushUi();
  updateBijirisNotificationUi();
  if (page === "home") {
    renderSurveys();
  }
  if (page === "bijiris") {
    renderBijirisPosts();
    if (!appState.bijirisPosts.length && !appState.bijirisLoading) {
      void loadBijirisPosts();
    }
  }
  if ((page === "history" || page === "home" || page === "measurements") && hasCustomerSession()) {
    void loadHistory();
  }
}

async function loadSurveys() {
  surveyList.innerHTML = `<div class="empty">読み込み中です。</div>`;
  try {
    const result = await api.request("/api/public/surveys");
    const surveys = Array.isArray(result.surveys) ? result.surveys : [];
    appState.surveys = surveys.length ? surveys : getFallbackSurveys();
    appState.publicInfo = {
      dataPolicyText: result.dataPolicyText || "",
      requireConsent: result.requireConsent === false ? false : true,
      consentText: result.consentText || "",
      version: result.version || APP_VERSION,
      pushAppId: result.pushAppId || getConfiguredPushAppId(),
    };
    renderSurveys();
    renderAnswerPanel();
    void initializePushNotifications();
  } catch (error) {
    reportClientError("customer.loadSurveys", error);
    const fallbackSurveys = getFallbackSurveys();
    if (fallbackSurveys.length) {
      appState.surveys = fallbackSurveys;
      renderSurveys();
      renderAnswerPanel();
      void initializePushNotifications();
      showToast("アンケート取得に失敗したため、保存済みの一覧を表示しています。");
      return;
    }
    surveyList.innerHTML = `<div class="empty">${escapeHtml(error.message || "アンケートを読み込めませんでした。")}</div>`;
  }
}

function normalizeBijirisContentFile(file, kind = "photo") {
  return {
    kind: normalizeText(file?.kind || kind) || kind,
    name: normalizeText(file?.name) || (kind === "pdf" ? "資料" : "写真"),
    type: normalizeText(file?.type || file?.mimeType),
    fileId: normalizeText(file?.fileId),
    url: normalizeText(file?.url),
    previewUrl: normalizeText(file?.previewUrl),
    downloadUrl: normalizeText(file?.downloadUrl),
    thumbnailUrl: normalizeText(file?.thumbnailUrl),
  };
}

function normalizeBijirisPost(post) {
  return {
    id: normalizeText(post?.id),
    title: normalizeText(post?.title),
    category: normalizeText(post?.category),
    summary: normalizeText(post?.summary),
    body: normalizeText(post?.body),
    status: normalizeText(post?.status) || "draft",
    pinned: post?.pinned === true,
    createdAt: normalizeText(post?.createdAt),
    updatedAt: normalizeText(post?.updatedAt),
    publishedAt: normalizeText(post?.publishedAt),
    photos: (Array.isArray(post?.photos) ? post.photos : [])
      .map((file) => normalizeBijirisContentFile(file, "photo"))
      .filter((file) => file.fileId || file.url || file.previewUrl || file.downloadUrl),
    documents: (Array.isArray(post?.documents) ? post.documents : [])
      .map((file) => normalizeBijirisContentFile(file, "pdf"))
      .filter((file) => file.fileId || file.url || file.previewUrl || file.downloadUrl),
  };
}

function getSelectedBijirisPost() {
  return appState.bijirisPosts.find((post) => post.id === appState.selectedBijirisPostId) || null;
}

function getBijirisFavoriteProfileKey(customer = appState.customer) {
  const name = normalizeText(customer?.name);
  const nameKana = normalizeText(customer?.nameKana);
  if (!name && !nameKana) return "";
  return [name, nameKana].filter(Boolean).join("::");
}

function getBijirisFavoriteIds() {
  const profileKey = getBijirisFavoriteProfileKey();
  if (!profileKey) return new Set();
  const ids = Array.isArray(appState.bijirisFavoritesByCustomer?.[profileKey])
    ? appState.bijirisFavoritesByCustomer[profileKey]
    : [];
  return new Set(ids.map((value) => normalizeText(value)).filter(Boolean));
}

function saveBijirisFavoriteIds(ids) {
  const profileKey = getBijirisFavoriteProfileKey();
  if (!profileKey) return;
  const nextFavorites = {
    ...(appState.bijirisFavoritesByCustomer && typeof appState.bijirisFavoritesByCustomer === "object"
      ? appState.bijirisFavoritesByCustomer
      : {}),
  };
  const normalizedIds = Array.from(new Set((Array.isArray(ids) ? ids : []).map((value) => normalizeText(value)).filter(Boolean)));
  if (normalizedIds.length) {
    nextFavorites[profileKey] = normalizedIds;
  } else {
    delete nextFavorites[profileKey];
  }
  appState.bijirisFavoritesByCustomer = nextFavorites;
  saveLocal(BIJIRIS_FAVORITES_KEY, nextFavorites);
}

function isBijirisFavorite(postId) {
  return getBijirisFavoriteIds().has(normalizeText(postId));
}

function toggleBijirisFavorite(postId) {
  const normalizedId = normalizeText(postId);
  if (!normalizedId || !hasCustomerSession()) return;
  const favoriteIds = getBijirisFavoriteIds();
  const alreadySaved = favoriteIds.has(normalizedId);
  if (alreadySaved) {
    favoriteIds.delete(normalizedId);
  } else {
    favoriteIds.add(normalizedId);
  }
  saveBijirisFavoriteIds(Array.from(favoriteIds));
  renderBijirisPosts();
  showToast(alreadySaved ? "お気に入りを解除しました。" : "お気に入りに保存しました。");
}

function createEmptyBijirisReaderState() {
  return {
    readIds: [],
    readLaterIds: [],
    historyIds: [],
  };
}

function getBijirisPostReadToken(postOrId) {
  const post = typeof postOrId === "object"
    ? postOrId
    : appState.bijirisPosts.find((item) => item.id === normalizeText(postOrId));
  const postId = normalizeText(post?.id || postOrId);
  const version = normalizeText(post?.updatedAt || post?.publishedAt || post?.createdAt);
  if (!postId) return "";
  return version ? `${postId}::${version}` : postId;
}

function getBijirisReaderStateProfileKey(customer = appState.customer) {
  return getBijirisFavoriteProfileKey(customer);
}

function getBijirisReaderState() {
  const profileKey = getBijirisReaderStateProfileKey();
  if (!profileKey) return createEmptyBijirisReaderState();
  const saved = appState.bijirisReaderStateByCustomer?.[profileKey];
  return {
    readIds: Array.isArray(saved?.readIds) ? saved.readIds.map((value) => normalizeText(value)).filter(Boolean) : [],
    readLaterIds: Array.isArray(saved?.readLaterIds) ? saved.readLaterIds.map((value) => normalizeText(value)).filter(Boolean) : [],
    historyIds: Array.isArray(saved?.historyIds) ? saved.historyIds.map((value) => normalizeText(value)).filter(Boolean) : [],
  };
}

function saveBijirisReaderState(nextState) {
  const profileKey = getBijirisReaderStateProfileKey();
  if (!profileKey) return;
  const normalized = {
    readIds: Array.from(new Set((Array.isArray(nextState?.readIds) ? nextState.readIds : []).map((value) => normalizeText(value)).filter(Boolean))),
    readLaterIds: Array.from(new Set((Array.isArray(nextState?.readLaterIds) ? nextState.readLaterIds : []).map((value) => normalizeText(value)).filter(Boolean))),
    historyIds: Array.from(new Set((Array.isArray(nextState?.historyIds) ? nextState.historyIds : []).map((value) => normalizeText(value)).filter(Boolean))).slice(0, BIJIRIS_HISTORY_LIMIT),
  };
  const nextMap = {
    ...(appState.bijirisReaderStateByCustomer && typeof appState.bijirisReaderStateByCustomer === "object"
      ? appState.bijirisReaderStateByCustomer
      : {}),
    [profileKey]: normalized,
  };
  appState.bijirisReaderStateByCustomer = nextMap;
  saveLocal(BIJIRIS_READER_STATE_KEY, nextMap);
}

function isBijirisRead(postId) {
  const token = getBijirisPostReadToken(postId);
  if (!token) return false;
  return getBijirisReaderState().readIds.includes(token);
}

function isBijirisReadLater(postId) {
  return getBijirisReaderState().readLaterIds.includes(normalizeText(postId));
}

function toggleBijirisReadLater(postId) {
  const normalizedId = normalizeText(postId);
  if (!normalizedId || !hasCustomerSession()) return;
  const nextState = getBijirisReaderState();
  const current = new Set(nextState.readLaterIds);
  const alreadySaved = current.has(normalizedId);
  if (alreadySaved) current.delete(normalizedId);
  else current.add(normalizedId);
  nextState.readLaterIds = Array.from(current);
  saveBijirisReaderState(nextState);
  renderBijirisPosts();
  showToast(alreadySaved ? "あとで読むを解除しました。" : "あとで読むに保存しました。");
}

function recordBijirisView(postId) {
  const normalizedId = normalizeText(postId);
  if (!normalizedId || !hasCustomerSession()) return;
  const nextState = getBijirisReaderState();
  const readIds = new Set(nextState.readIds);
  const readToken = getBijirisPostReadToken(normalizedId);
  if (readToken) readIds.add(readToken);
  nextState.readIds = Array.from(readIds);
  nextState.historyIds = [normalizedId]
    .concat(nextState.historyIds.filter((value) => value !== normalizedId))
    .slice(0, BIJIRIS_HISTORY_LIMIT);
  saveBijirisReaderState(nextState);
}

function migrateBijirisReaderStateTokens() {
  if (!hasCustomerSession() || !appState.bijirisPosts.length) return;
  const currentState = getBijirisReaderState();
  let changed = false;
  const migratedReadIds = currentState.readIds
    .map((value) => {
      const normalized = normalizeText(value);
      if (!normalized || normalized.includes("::")) return normalized;
      const token = getBijirisPostReadToken(normalized);
      if (token && token !== normalized) changed = true;
      return token || normalized;
    })
    .filter(Boolean);
  if (!changed) return;
  saveBijirisReaderState({
    ...currentState,
    readIds: migratedReadIds,
  });
}

function getBijirisRecentHistoryPosts() {
  const historyIds = getBijirisReaderState().historyIds;
  if (!historyIds.length) return [];
  const postMap = new Map(appState.bijirisPosts.map((post) => [post.id, post]));
  return historyIds.map((id) => postMap.get(id)).filter(Boolean);
}

function getUnreadBijirisPosts() {
  return sortBijirisPosts(appState.bijirisPosts).filter((post) => !isBijirisRead(post));
}

function getNewUnreadBijirisPosts() {
  return getUnreadBijirisPosts().filter((post) => isBijirisPostNew(post));
}

function getBijirisSearchSource(post) {
  return [
    post?.title,
    post?.category,
    post?.summary,
    post?.body,
    ...(Array.isArray(post?.documents) ? post.documents.map((file) => file.name) : []),
    ...(Array.isArray(post?.photos) ? post.photos.map((file) => file.name) : []),
  ]
    .map((value) => normalizeText(value))
    .filter(Boolean)
    .join("\n")
    .toLocaleLowerCase();
}

function getAnswerValueTokens(value) {
  return normalizeText(value)
    .split(/\s*,\s*|\n+/)
    .map((item) => normalizeText(item))
    .filter(Boolean);
}

function collectCustomerInterestTopics() {
  const matchedTopics = new Set();
  getVisibleHistoryResponses().forEach((response) => {
    (Array.isArray(response?.answers) ? response.answers : []).forEach((answer) => {
      const value = normalizeText(answer?.value);
      if (!value) return;
      BIJIRIS_RECOMMENDATION_TOPICS.forEach((topic) => {
        if (topic.answerTerms.some((term) => value.includes(term))) {
          matchedTopics.add(topic.id);
        }
      });
    });
  });
  return Array.from(matchedTopics);
}

function getPostTopicIds(post) {
  const source = getBijirisSearchSource(post);
  return BIJIRIS_RECOMMENDATION_TOPICS
    .filter((topic) => topic.postTerms.some((term) => source.includes(term)))
    .map((topic) => topic.id);
}

function scoreBijirisPostForTopics(post, topicIds) {
  if (!topicIds.length) return 0;
  const source = getBijirisSearchSource(post);
  let score = 0;
  topicIds.forEach((topicId) => {
    const topic = BIJIRIS_RECOMMENDATION_TOPICS.find((item) => item.id === topicId);
    if (!topic) return;
    if (topic.postTerms.some((term) => source.includes(term))) {
      score += 3;
    }
  });
  return score;
}

function getRecommendedBijirisPosts(limit = 3, excludeIds = []) {
  const exclude = new Set((Array.isArray(excludeIds) ? excludeIds : []).map((value) => normalizeText(value)).filter(Boolean));
  const topics = collectCustomerInterestTopics();
  const scored = sortBijirisPosts(appState.bijirisPosts)
    .filter((post) => !exclude.has(post.id))
    .map((post) => ({
      post,
      score: scoreBijirisPostForTopics(post, topics),
    }))
    .filter((entry) => entry.score > 0)
    .sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      if (Boolean(b.post?.pinned) !== Boolean(a.post?.pinned)) return b.post?.pinned ? 1 : -1;
      return new Date(b.post?.publishedAt || b.post?.updatedAt || 0) - new Date(a.post?.publishedAt || a.post?.updatedAt || 0);
    })
    .map((entry) => entry.post);
  if (scored.length) return scored.slice(0, limit);
  return sortBijirisPosts(appState.bijirisPosts)
    .filter((post) => !exclude.has(post.id))
    .slice(0, limit);
}

function getRelatedBijirisPosts(post, limit = 3) {
  const baseTopics = new Set(getPostTopicIds(post));
  return sortBijirisPosts(appState.bijirisPosts)
    .filter((item) => item.id !== post.id)
    .map((item) => {
      let score = 0;
      if (normalizeText(item.category) && normalizeText(item.category) === normalizeText(post.category)) {
        score += 3;
      }
      const sharedTopics = getPostTopicIds(item).filter((topicId) => baseTopics.has(topicId)).length;
      score += sharedTopics * 2;
      return { post: item, score };
    })
    .filter((entry) => entry.score > 0)
    .sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      return new Date(b.post?.publishedAt || b.post?.updatedAt || 0) - new Date(a.post?.publishedAt || a.post?.updatedAt || 0);
    })
    .slice(0, limit)
    .map((entry) => entry.post);
}

function createPdfThumbnailDataUrl(fileName) {
  const label = normalizeText(fileName || "PDF").slice(0, 24);
  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="360" height="240" viewBox="0 0 360 240">
      <rect width="360" height="240" rx="18" fill="#fbf7f2"/>
      <rect x="20" y="20" width="320" height="200" rx="14" fill="#ffffff" stroke="#eadfd2"/>
      <rect x="40" y="38" width="74" height="28" rx="8" fill="#c95f50"/>
      <text x="77" y="57" text-anchor="middle" font-family="Arial, sans-serif" font-size="16" font-weight="700" fill="#ffffff">PDF</text>
      <text x="40" y="102" font-family="Arial, sans-serif" font-size="20" font-weight="700" fill="#5b493d">${escapeHtml(label)}</text>
      <text x="40" y="138" font-family="Arial, sans-serif" font-size="14" fill="#866f60">資料を開いて確認できます</text>
    </svg>
  `;
  return `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(svg)}`;
}

function getBijirisDocumentThumbnailSrc(file) {
  return normalizeText(file?.thumbnailUrl) || createPdfThumbnailDataUrl(file?.name);
}

function sortBijirisPosts(posts) {
  return posts.slice().sort((a, b) => {
    if (Boolean(b?.pinned) !== Boolean(a?.pinned)) {
      return b?.pinned ? 1 : -1;
    }
    const aTime = new Date(a?.publishedAt || a?.updatedAt || a?.createdAt || 0).getTime();
    const bTime = new Date(b?.publishedAt || b?.updatedAt || b?.createdAt || 0).getTime();
    return bTime - aTime;
  });
}

function isBijirisPostNew(post) {
  const publishedAt = normalizeText(post?.publishedAt) || normalizeText(post?.createdAt) || normalizeText(post?.updatedAt);
  if (!publishedAt) return false;
  const publishedTime = new Date(publishedAt).getTime();
  if (!Number.isFinite(publishedTime)) return false;
  return Date.now() - publishedTime <= BIJIRIS_NEW_BADGE_DAYS * 24 * 60 * 60 * 1000;
}

function getBijirisCategories(posts) {
  const categories = [];
  const seen = new Set();
  sortBijirisPosts(posts).forEach((post) => {
    const category = normalizeText(post?.category);
    if (!category || seen.has(category)) return;
    seen.add(category);
    categories.push(category);
  });
  return categories;
}

function getFilteredBijirisPosts() {
  const categories = getBijirisCategories(appState.bijirisPosts);
  if (appState.selectedBijirisCategory !== "all" && !categories.includes(appState.selectedBijirisCategory)) {
    appState.selectedBijirisCategory = "all";
  }
  let posts = sortBijirisPosts(appState.bijirisPosts);
  const searchQuery = normalizeText(appState.bijirisSearchQuery).toLocaleLowerCase();
  if (searchQuery) {
    posts = posts.filter((post) => getBijirisSearchSource(post).includes(searchQuery));
  }
  if (appState.selectedBijirisCategory !== "all") {
    posts = posts.filter((post) => normalizeText(post.category) === appState.selectedBijirisCategory);
  }
  if (appState.showFavoriteBijirisOnly) {
    const favoriteIds = getBijirisFavoriteIds();
    posts = posts.filter((post) => favoriteIds.has(post.id));
  }
  if (appState.showReadLaterBijirisOnly) {
    const readLaterIds = new Set(getBijirisReaderState().readLaterIds);
    posts = posts.filter((post) => readLaterIds.has(post.id));
  }
  return posts;
}

function renderBijirisBadges(post) {
  return `
    ${!isBijirisRead(post.id) ? `<span class="badge info">未読</span>` : ""}
    ${post.pinned ? `<span class="badge closed">重要</span>` : ""}
    ${isBijirisPostNew(post) ? `<span class="badge warn">新着</span>` : ""}
    ${post.photos.length ? `<span class="badge draft">写真 ${post.photos.length}</span>` : ""}
    ${post.documents.length ? `<span class="badge">PDF ${post.documents.length}</span>` : ""}
  `;
}

function renderBijirisHomeNotice() {
  if (!hasCustomerSession() || !appState.bijirisPosts.length) return "";
  const unreadCount = getUnreadBijirisPosts().length;
  const newUnreadCount = getNewUnreadBijirisPosts().length;
  const readLaterCount = getBijirisReaderState().readLaterIds.length;
  if (!unreadCount && !newUnreadCount && !readLaterCount) return "";
  return `
    <article class="history-card bijiris-home-notice">
      <div class="section-head">
        <div>
          <strong>豆知識のお知らせ</strong>
          <div class="meta">
            ${
              newUnreadCount
                ? `新着 ${newUnreadCount}件 / 未読 ${unreadCount}件`
                : `未読 ${unreadCount}件`
            }
            ${readLaterCount ? ` / あとで読む ${readLaterCount}件` : ""}
          </div>
        </div>
        <button class="ghost-button" type="button" data-open-bijiris-home>豆知識を見る</button>
      </div>
    </article>
  `;
}

function renderBijirisListSwitcher() {
  const favoriteCount = sortBijirisPosts(appState.bijirisPosts).filter((post) => isBijirisFavorite(post.id)).length;
  const readLaterCount = getBijirisReaderState().readLaterIds.length;
  return `
    <div class="history-card bijiris-list-switcher">
      <div class="section-head">
        <div>
          <strong>一覧</strong>
          <div class="meta">お気に入り ${favoriteCount}件 / あとで読む ${readLaterCount}件</div>
        </div>
      </div>
      <div class="bijiris-filter-row">
        <button
          class="bijiris-filter-chip ${!appState.showFavoriteBijirisOnly && !appState.showReadLaterBijirisOnly ? "active" : ""}"
          type="button"
          data-bijiris-list-mode="all"
        >
          投稿一覧
        </button>
        <button
          class="bijiris-filter-chip ${appState.showFavoriteBijirisOnly ? "active" : ""}"
          type="button"
          data-bijiris-list-mode="favorites"
        >
          お気に入り⭐️
        </button>
        <button
          class="bijiris-filter-chip ${appState.showReadLaterBijirisOnly ? "active" : ""}"
          type="button"
          data-bijiris-list-mode="read-later"
        >
          あとで読む
        </button>
      </div>
    </div>
  `;
}

function renderBijirisToolbar() {
  const categories = getBijirisCategories(appState.bijirisPosts);
  const favoriteCount = sortBijirisPosts(appState.bijirisPosts).filter((post) => isBijirisFavorite(post.id)).length;
  const readerState = getBijirisReaderState();
  const unreadCount = getUnreadBijirisPosts().length;
  const newUnreadCount = getNewUnreadBijirisPosts().length;
  return `
    <div class="history-card bijiris-toolbar-card">
      <div class="section-head bijiris-toolbar-head">
        <div>
          <strong>探す</strong>
          <div class="meta">新着 ${newUnreadCount}件 / 未読 ${unreadCount}件 / あとで読む ${readerState.readLaterIds.length}件</div>
        </div>
      </div>
      <label>
        キーワード検索
        <input
          id="bijirisSearchInput"
          type="search"
          placeholder="例: 尿もれ / 骨盤底筋 / 産後 / 姿勢"
          value="${escapeHtml(appState.bijirisSearchQuery)}"
        />
      </label>
      <div class="bijiris-filter-row">
        <button class="bijiris-filter-chip ${appState.selectedBijirisCategory === "all" ? "active" : ""}" type="button" data-bijiris-category="all">すべて</button>
        ${categories
          .map((category) => `
            <button
              class="bijiris-filter-chip ${appState.selectedBijirisCategory === category ? "active" : ""}"
              type="button"
              data-bijiris-category="${escapeHtml(category)}"
            >
              ${escapeHtml(category)}
            </button>
          `)
          .join("")}
      </div>
      <div class="bijiris-filter-row bijiris-filter-row-secondary">
        <button class="bijiris-filter-chip ${appState.showFavoriteBijirisOnly ? "active" : ""}" type="button" data-bijiris-favorite-filter>
          ${appState.showFavoriteBijirisOnly ? "お気に入りのみ表示中" : "お気に入りのみ"}
        </button>
        <button class="bijiris-filter-chip ${appState.showReadLaterBijirisOnly ? "active" : ""}" type="button" data-bijiris-read-later-filter>
          ${appState.showReadLaterBijirisOnly ? "あとで読むのみ表示中" : "あとで読むのみ"}
        </button>
        <span class="meta">お気に入り ${favoriteCount}件 / あとで読む ${readerState.readLaterIds.length}件</span>
      </div>
    </div>
  `;
}

function renderMultilineText(text) {
  const value = normalizeText(text);
  if (!value) return "";
  return value
    .split(/\n{2,}/)
    .map((block) => `<p>${escapeHtml(block).replaceAll("\n", "<br />")}</p>`)
    .join("");
}

function renderBijirisDocumentPreview(file, index, compact = false) {
  const href = getPhotoOpenHref(file);
  const thumbnail = getBijirisDocumentThumbnailSrc(file);
  return `
    <a
      class="history-card bijiris-document-card ${compact ? "compact" : ""}"
      href="${escapeHtml(href)}"
      target="_blank"
      rel="noreferrer"
    >
      <img class="bijiris-document-thumb" src="${escapeHtml(thumbnail)}" alt="${escapeHtml(file.name || `資料${index + 1}`)}" />
      <strong>${escapeHtml(file.name || `資料${index + 1}`)}</strong>
      <div class="meta">PDFを開く</div>
    </a>
  `;
}

function renderBijirisFavoriteToggle(postId, favoriteSaved) {
  return `
    <label class="bijiris-star-toggle ${favoriteSaved ? "active" : ""}" data-bijiris-action-stop>
      <input
        class="bijiris-star-input"
        type="checkbox"
        data-toggle-bijiris-favorite="${escapeHtml(postId)}"
        ${favoriteSaved ? "checked" : ""}
      />
      <span class="bijiris-star-icon" aria-hidden="true">⭐️</span>
      <span class="bijiris-star-label">お気に入り</span>
    </label>
  `;
}

function renderBijirisPostCard(post) {
  const publishedAt = post.publishedAt || post.updatedAt || post.createdAt;
  const preview = post.summary || post.body.slice(0, 90);
  const favoriteSaved = isBijirisFavorite(post.id);
  const readLaterSaved = isBijirisReadLater(post.id);
  return `
    <article
      class="history-card bijiris-post-card ${post.pinned ? "is-pinned" : ""}"
      role="button"
      tabindex="0"
      data-open-bijiris-post="${escapeHtml(post.id)}"
    >
      <div class="section-head bijiris-post-head">
        <div>
          <strong>${escapeHtml(post.title)}</strong>
          <div class="meta">${escapeHtml(post.category || "豆知識")} / ${escapeHtml(formatDate(publishedAt))}</div>
        </div>
        <div class="action-row">
          ${renderBijirisFavoriteToggle(post.id, favoriteSaved)}
          <button
            class="ghost-button bijiris-action-button ${readLaterSaved ? "active" : ""}"
            type="button"
            data-toggle-bijiris-read-later="${escapeHtml(post.id)}"
          >
            ${readLaterSaved ? "あとで読む済み" : "あとで読む"}
          </button>
        </div>
      </div>
      <div class="action-row">${renderBijirisBadges(post)}</div>
      ${preview ? `<div class="meta bijiris-post-preview">${escapeHtml(preview)}</div>` : `<div class="meta">本文は詳細で確認できます。</div>`}
    </article>
  `;
}

function renderBijirisPostDetail(post) {
  const publishedAt = post.publishedAt || post.updatedAt || post.createdAt;
  const favoriteSaved = isBijirisFavorite(post.id);
  const readLaterSaved = isBijirisReadLater(post.id);
  const relatedPosts = getRelatedBijirisPosts(post, 3);
  return `
    <div class="section-head">
      <div>
        <strong>${escapeHtml(post.title)}</strong>
        <div class="meta">${escapeHtml(post.category || "豆知識")} / ${escapeHtml(formatDate(publishedAt))}</div>
      </div>
      <div class="action-row bijiris-detail-actions">
        ${renderBijirisBadges(post)}
        ${renderBijirisFavoriteToggle(post.id, favoriteSaved)}
        <button
          class="ghost-button bijiris-action-button ${readLaterSaved ? "active" : ""}"
          type="button"
          data-toggle-bijiris-read-later="${escapeHtml(post.id)}"
        >
          ${readLaterSaved ? "あとで読む済み" : "あとで読む"}
        </button>
        <button class="ghost-button" type="button" data-back-bijiris-list>戻る</button>
      </div>
    </div>
    ${post.summary ? `<article class="history-card bijiris-summary-card"><strong>概要</strong><div class="bijiris-richtext">${renderMultilineText(post.summary)}</div></article>` : ""}
    <article class="history-card">
      <strong>内容</strong>
      <div class="bijiris-richtext">
        ${post.body ? renderMultilineText(post.body) : `<div class="meta">本文はありません。</div>`}
      </div>
    </article>
    <article class="history-card">
      <strong>写真</strong>
      ${
        post.photos.length
          ? `
            <div class="bijiris-photo-grid">
              ${post.photos
                .map((file, index) => {
                  const preview = getPhotoPreviewSrc(file);
                  const href = getPhotoOpenHref(file);
                  return `
                    <a class="bijiris-photo-link" href="${escapeHtml(href)}" target="_blank" rel="noreferrer">
                      ${preview ? `<img class="bijiris-photo-image" src="${escapeHtml(preview)}" alt="${escapeHtml(file.name || `写真${index + 1}`)}" />` : ""}
                    </a>
                  `;
                })
                .join("")}
            </div>
          `
          : `<div class="empty">写真はありません。</div>`
      }
    </article>
    <article class="history-card">
      <strong>PDF資料</strong>
      ${
        post.documents.length
          ? `
            <div class="bijiris-document-list">
              ${post.documents.map((file, index) => renderBijirisDocumentPreview(file, index)).join("")}
            </div>
          `
          : `<div class="empty">PDF資料はありません。</div>`
      }
    </article>
    ${
      relatedPosts.length
        ? `
          <article class="history-card">
            <strong>関連する豆知識</strong>
            <div class="meta">近いテーマの記事を表示しています。</div>
            <div class="bijiris-panel">
              ${relatedPosts.map(renderBijirisPostCard).join("")}
            </div>
          </article>
        `
        : ""
    }
  `;
}

function attachBijirisPostActions() {
  bijirisPanel.querySelectorAll("[data-bijiris-action-stop]").forEach((node) => {
    node.addEventListener("click", (event) => {
      event.stopPropagation();
    });
  });
  bijirisPanel.querySelectorAll("[data-toggle-bijiris-favorite]").forEach((button) => {
    button.addEventListener("change", (event) => {
      event.preventDefault();
      event.stopPropagation();
      toggleBijirisFavorite(button.dataset.toggleBijirisFavorite || "");
    });
    button.addEventListener("click", (event) => {
      event.stopPropagation();
    });
  });
  bijirisPanel.querySelectorAll("[data-toggle-bijiris-read-later]").forEach((button) => {
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      toggleBijirisReadLater(button.dataset.toggleBijirisReadLater || "");
    });
  });
  bijirisPanel.querySelectorAll("[data-open-bijiris-post]").forEach((button) => {
    button.addEventListener("click", () => {
      appState.selectedBijirisPostId = button.dataset.openBijirisPost || "";
      recordBijirisView(appState.selectedBijirisPostId);
      renderSurveys();
      renderBijirisPosts();
      window.scrollTo({ top: 0, behavior: "auto" });
    });
    button.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" && event.key !== " ") return;
      event.preventDefault();
      button.click();
    });
  });
  bijirisPanel.querySelectorAll("[data-bijiris-category]").forEach((button) => {
    button.addEventListener("click", () => {
      appState.selectedBijirisCategory = button.dataset.bijirisCategory || "all";
      renderBijirisPosts();
    });
  });
  bijirisPanel.querySelectorAll("[data-bijiris-list-mode]").forEach((button) => {
    button.addEventListener("click", () => {
      const mode = button.dataset.bijirisListMode || "all";
      appState.showFavoriteBijirisOnly = mode === "favorites";
      appState.showReadLaterBijirisOnly = mode === "read-later";
      renderBijirisPosts();
    });
  });
  bijirisPanel.querySelector("#bijirisSearchInput")?.addEventListener("input", (event) => {
    const nextQuery = normalizeText(event.currentTarget.value);
    appState.bijirisSearchQuery = nextQuery;
    renderBijirisPosts();
    const input = bijirisPanel.querySelector("#bijirisSearchInput");
    input?.focus();
    if (typeof input?.setSelectionRange === "function") {
      input.setSelectionRange(nextQuery.length, nextQuery.length);
    }
  });
  bijirisPanel.querySelector("[data-bijiris-favorite-filter]")?.addEventListener("click", () => {
    appState.showFavoriteBijirisOnly = !appState.showFavoriteBijirisOnly;
    renderBijirisPosts();
  });
  bijirisPanel.querySelector("[data-bijiris-read-later-filter]")?.addEventListener("click", () => {
    appState.showReadLaterBijirisOnly = !appState.showReadLaterBijirisOnly;
    renderBijirisPosts();
  });
}

function renderBijirisPosts() {
  if (!bijirisPanel) return;
  updateBijirisNotificationUi();
  if (!hasCustomerSession()) {
    bijirisPanel.innerHTML = `<div class="empty">先にログインしてください。</div>`;
    return;
  }
  if (appState.bijirisLoading && !appState.bijirisPosts.length) {
    bijirisPanel.innerHTML = `<div class="empty">読み込み中です。</div>`;
    return;
  }
  if (appState.bijirisLoadError && !appState.bijirisPosts.length) {
    bijirisPanel.innerHTML = `<div class="empty">${escapeHtml(appState.bijirisLoadError)}</div>`;
    return;
  }
  const selectedPost = getSelectedBijirisPost();
  if (appState.selectedBijirisPostId && !selectedPost) {
    appState.selectedBijirisPostId = "";
  }
  if (selectedPost) {
    recordBijirisView(selectedPost.id);
    bijirisPanel.innerHTML = `
      ${appState.bijirisLoadError ? `<div class="meta">更新に失敗したため前回取得内容を表示しています。</div>` : ""}
      ${renderBijirisPostDetail(selectedPost)}
    `;
    bijirisPanel.querySelector("[data-back-bijiris-list]")?.addEventListener("click", () => {
      appState.selectedBijirisPostId = "";
      renderBijirisPosts();
    });
    attachBijirisPostActions();
    return;
  }
  let visiblePosts = sortBijirisPosts(appState.bijirisPosts);
  if (appState.showFavoriteBijirisOnly) {
    visiblePosts = visiblePosts.filter((post) => isBijirisFavorite(post.id));
  } else if (appState.showReadLaterBijirisOnly) {
    visiblePosts = visiblePosts.filter((post) => isBijirisReadLater(post.id));
  }
  const emptyMessage = appState.showFavoriteBijirisOnly
    ? "お気に入りはまだありません。"
    : appState.showReadLaterBijirisOnly
      ? "あとで読むはまだありません。"
      : "まだ豆知識の投稿はありません。";
  bijirisPanel.innerHTML = `
    ${appState.bijirisLoadError ? `<div class="meta">更新に失敗したため前回取得内容を表示しています。</div>` : ""}
    ${renderBijirisListSwitcher()}
    ${visiblePosts.length ? `<div class="bijiris-panel">${visiblePosts.map(renderBijirisPostCard).join("")}</div>` : `<div class="empty">${emptyMessage}</div>`}
  `;
  attachBijirisPostActions();
}

async function loadBijirisPosts() {
  appState.bijirisLoading = true;
  appState.bijirisLoadError = "";
  renderBijirisPosts();
  try {
    const result = await api.request("/api/public/bijiris-posts");
    appState.bijirisPosts = Array.isArray(result.posts)
      ? sortBijirisPosts(result.posts.map(normalizeBijirisPost).filter((post) => post.id))
      : [];
    migrateBijirisReaderStateTokens();
    appState.bijirisLoading = false;
    appState.bijirisLoadError = "";
    renderSurveys();
    renderBijirisPosts();
    applyLaunchRouteIfPossible();
  } catch (error) {
    appState.bijirisLoading = false;
    appState.bijirisLoadError = error.message || "豆知識を読み込めませんでした。";
    reportClientError("customer.loadBijirisPosts", error);
    renderSurveys();
    renderBijirisPosts();
  }
}

async function loadHistory() {
  if (!hasCustomerSession()) {
    appState.history = [];
    appState.measurements = [];
    appState.historyLoading = false;
    appState.historyLoadError = "";
    renderHomeTicketStatus();
    historyList.innerHTML = `<div class="empty">先にログインしてください。</div>`;
    renderMeasurements();
    return;
  }

  appState.historyLoading = true;
  appState.historyLoadError = "";
  renderHomeTicketStatus();
  historyList.innerHTML = `<div class="empty">読み込み中です。</div>`;
  try {
    const params = buildHistorySearchParams(appState.customer);
    const result = await api.request(
      `/api/public/responses?${params.toString()}`,
    );
    appState.history = Array.isArray(result.responses) ? result.responses : [];
    appState.measurements = Array.isArray(result.measurements)
      ? result.measurements.map(normalizeMeasurementRecord)
      : [];
    syncCustomerProfileFromServer(result.customerProfile);
    syncActiveTicketCardOverrideFromServer(result.customerProfile);
    syncActiveTicketCardOverrideWithHistory();
    appState.historyLoading = false;
    appState.historyLoadError = "";
    renderHomeTicketStatus();
    renderSurveys();
    renderHistory();
    renderMeasurements();
    if (document.querySelector("#page-bijiris")?.classList.contains("active")) {
      renderBijirisPosts();
    }
    if (
      appState.selectedSurveyId === SESSION_SURVEY_ID &&
      getSessionTypeSelection(SESSION_SURVEY_ID) === "回数券" &&
      getSessionTicketPlanSelection(SESSION_SURVEY_ID) &&
      !getSessionTicketRoundSelection(SESSION_SURVEY_ID)
    ) {
      renderAnswerPanel();
    }
  } catch (error) {
    appState.history = [];
    appState.measurements = [];
    appState.historyLoading = false;
    appState.historyLoadError = error.message || "履歴を読み込めませんでした。";
    renderHomeTicketStatus();
    reportClientError("customer.loadHistory", error, { name: appState.customer.name });
    historyList.innerHTML = `<div class="empty">${escapeHtml(error.message || "履歴を読み込めませんでした。")}</div>`;
    renderMeasurements();
    if (
      appState.selectedSurveyId === SESSION_SURVEY_ID &&
      getSessionTypeSelection(SESSION_SURVEY_ID) === "回数券" &&
      getSessionTicketPlanSelection(SESSION_SURVEY_ID) &&
      !getSessionTicketRoundSelection(SESSION_SURVEY_ID)
    ) {
      renderAnswerPanel();
    }
  }
}

function syncCustomerForms() {
  if (customerLoginForm) {
    customerLoginForm.elements.name.value = appState.customer.name || "";
    customerLoginForm.elements.nameKana.value = appState.customer.nameKana || "";
  }
  if (customerForm) {
    customerForm.elements.name.value = appState.customer.name || "";
    customerForm.elements.nameKana.value = appState.customer.nameKana || "";
  }
  renderCustomerMemberInfo();
}

function renderCustomerMemberInfo() {
  if (!customerMemberInfo) return;
  const memberNumber = getCustomerMemberNumber();
  customerMemberInfo.innerHTML = memberNumber
    ? `
        <strong>会員番号</strong>
        <div>${escapeHtml(memberNumber)}</div>
      `
    : `
        <strong>会員番号</strong>
        <div class="meta">会員登録後に自動発行されます。</div>
      `;
}

function readCustomerProfileFromForm(form) {
  const formData = new FormData(form);
  return normalizeCustomerProfile({
    name: formData.get("name"),
    memberNumber: appState.customer.memberNumber,
    nameKana: formData.get("nameKana"),
    historyMatchMode: appState.customer.historyMatchMode,
    measurementTargets: appState.customer.measurementTargets,
  });
}

function validateCustomerProfile(profile) {
  if (!profile.name) {
    throw new Error("お名前を入力してください。");
  }
  if (!profile.nameKana) {
    throw new Error("フリガナを入力してください。");
  }
  if (!isKatakanaName(profile.nameKana)) {
    throw new Error("フリガナはカタカナで入力してください。");
  }
}

async function applyCustomerSession(profile, options = {}) {
  const { recovery = false } = options;
  validateCustomerProfile(profile);
  const nextCustomer = {
    ...profile,
    historyMatchMode: recovery ? "name" : "device",
  };

  if (recovery) {
    const result = await api.request(`/api/public/responses?${buildHistorySearchParams(nextCustomer, {
      recovery: true,
    }).toString()}`);
    const responses = Array.isArray(result.responses) ? result.responses : [];
    if (!responses.length) {
      throw new Error("一致する回答履歴が見つかりませんでした。");
    }
    Object.assign(nextCustomer, mergeCustomerProfile(nextCustomer, result.customerProfile));
    appState.history = responses;
    appState.historyLoading = false;
    appState.historyLoadError = "";
  }

  appState.customer = nextCustomer;
  appState.pushServerSignature = "";
  saveLocal(CUSTOMER_KEY, appState.customer);
  syncCustomerForms();
  await loadHistory();
  await loadBijirisPosts();
  await initializePushNotifications();
  if (applyLaunchRouteIfPossible()) {
    renderHomeTicketStatus();
    renderHistory();
    showToast(recovery ? "名前一致でアカウントを復旧しました。" : "会員登録しました。");
    return;
  }
  setPage("home");
  renderHomeTicketStatus();
  renderHistory();
  showToast(recovery ? "名前一致でアカウントを復旧しました。" : "会員登録しました。");
}

function selectSurvey(surveyId) {
  if (!hasCustomerSession()) {
    setPage("login");
    showToast("先にログインしてください。");
    return;
  }
  appState.selectedSurveyId = surveyId;
  appState.panelMode = "form";
  appState.confirmPayload = null;
  appState.editingResponseId = "";
  appState.lastSubmittedResponse = null;
  appState.lastSubmissionWasEdit = false;
  setPage("survey");
  window.scrollTo({ top: 0, behavior: "auto" });
  renderSurveys();
  renderAnswerPanel();
}

function renderSurveys() {
  updatePushUi();
  updateBijirisNotificationUi();
  if (!appState.surveys.length) {
    surveyList.innerHTML = `${renderPendingNotice()}${renderBijirisHomeNotice()}<div class="empty">現在回答できるアンケートはありません。</div>`;
    document.querySelector("[data-open-bijiris-home]")?.addEventListener("click", () => {
      setPage("bijiris");
    });
    attachCommonButtons();
    return;
  }

  surveyList.innerHTML = `
    ${renderPendingNotice()}
    ${renderBijirisHomeNotice()}
    ${appState.surveys
      .map((survey) => {
        const availability = getSurveyAvailability(survey);
        const draft = getSurveyDraft(survey.id);
        const hasDraft = hasDraftContent(draft);
        const draftSavedAt = getDraftSavedAt(survey.id);
        return `
          <button
            class="survey-card ${survey.id === appState.selectedSurveyId ? "active" : ""} ${availability.open ? "" : "disabled"} ${availability.nearDeadline ? "deadline-soon" : ""} ${hasDraft ? "has-draft" : ""}"
            type="button"
            data-survey-id="${survey.id}"
          >
            <strong>${escapeHtml(survey.title)}</strong>
            <span>${escapeHtml(survey.description)}</span>
            <div class="survey-card-meta">
              <span class="badge ${availability.open ? "open" : "closed"}">${escapeHtml(availability.label)}</span>
              ${hasDraft ? `<span class="badge draft">下書きあり</span>` : ""}
              <span class="meta">${escapeHtml(availability.detail)}</span>
            </div>
            ${
              hasDraft && draftSavedAt
                ? `<div class="meta">最終保存: ${escapeHtml(formatDate(draftSavedAt))}</div>`
                : ""
            }
          </button>
        `;
      })
      .join("")}
  `;

  document.querySelectorAll("[data-survey-id]").forEach((button) => {
    button.addEventListener("click", () => {
      selectSurvey(button.dataset.surveyId);
    });
  });
  document.querySelector("[data-open-bijiris-home]")?.addEventListener("click", () => {
    setPage("bijiris");
  });

  attachCommonButtons();
}

function renderProgressBar(progress) {
  return `
    <div class="progress-card">
      <div class="progress-head">
        <strong>入力状況</strong>
        <span>${progress.answeredRequired}/${progress.requiredTotal} 必須回答</span>
      </div>
      <div class="progress-bar">
        <div class="progress-fill" style="width: ${progress.percent}%"></div>
      </div>
    </div>
  `;
}

function renderDraftStatusCard(surveyId) {
  const draft = getSurveyDraft(surveyId);
  const hasDraft = hasDraftContent(draft);
  const savedAt = getDraftSavedAt(surveyId);
  return `
    <article class="draft-status-card">
      <div class="draft-status-head">
        <strong>下書き保存</strong>
        <span class="badge draft">自動保存</span>
      </div>
      <div class="meta" data-draft-status-text>
        ${
          hasDraft && savedAt
            ? `この端末に保存中です。最終保存: ${escapeHtml(formatDate(savedAt))}`
            : "入力内容はこの端末に自動保存されます。"
        }
      </div>
      <div class="action-row">
        <button id="clearDraftButton" class="ghost-button" type="button" ${hasDraft ? "" : "hidden"}>
          下書きを破棄
        </button>
      </div>
    </article>
  `;
}

function refreshProgressDisplay(survey) {
  const progress = getProgress(survey, getSurveyDraft(survey.id));
  document.querySelector(".progress-fill")?.setAttribute("style", `width: ${progress.percent}%`);
  const progressHead = document.querySelector(".progress-head span");
  if (progressHead) {
    progressHead.textContent = `${progress.answeredRequired}/${progress.requiredTotal} 必須回答`;
  }
  refreshDraftStatusDisplay(survey.id);
}

function refreshDraftStatusDisplay(surveyId) {
  const textNode = document.querySelector("[data-draft-status-text]");
  const clearButton = document.querySelector("#clearDraftButton");
  if (!textNode) return;
  const draft = getSurveyDraft(surveyId);
  const hasDraft = hasDraftContent(draft);
  const savedAt = getDraftSavedAt(surveyId);
  textNode.textContent = hasDraft && savedAt
    ? `この端末に保存中です。最終保存: ${formatDate(savedAt)}`
    : "入力内容はこの端末に自動保存されます。";
  if (clearButton) {
    clearButton.hidden = !hasDraft;
  }
}

function renderSessionTypeStep(survey, surveyId) {
  const sessionTypeQuestion = getSessionTypeQuestion(survey);
  const selectedType = getSessionTypeSelection(surveyId);
  answerPanel.innerHTML = `
    ${renderPendingNotice()}
    <div class="section-head survey-toolbar">
      <div>
        <h2>${escapeHtml(survey.title)}</h2>
        <p>施術内容を選択してください。</p>
      </div>
      <button id="backToHomeButton" class="ghost-button" type="button">一覧へ戻る</button>
    </div>
    <div class="question-list">
      <div class="question-block">
        <strong>${escapeHtml(sessionTypeQuestion.label)}</strong>
        <div class="selection-grid">
          ${sessionTypeQuestion.options
            .map(
              (option) => `
                <button
                  class="selection-button ${selectedType === option ? "active" : ""}"
                  type="button"
                  data-session-type-option="${escapeHtml(option)}"
                >
                  ${escapeHtml(option)}
                </button>
              `,
            )
            .join("")}
        </div>
      </div>
    </div>
  `;

  document.querySelector("#backToHomeButton").addEventListener("click", () => setPage("home"));
  document.querySelectorAll("[data-session-type-option]").forEach((button) => {
    button.addEventListener("click", () => {
      const nextType = normalizeText(button.dataset.sessionTypeOption);
      if (!nextType) return;
      updateDraftValue(surveyId, SESSION_TYPE_QUESTION_ID, nextType);
      if (nextType !== "回数券") {
        clearSessionTicketProgressSelections(surveyId);
        updateDraftValue(surveyId, SESSION_TICKET_PLAN_QUESTION_ID, "");
      }
      renderAnswerPanel();
    });
  });
  attachCommonButtons();
}

function renderSessionTicketPlanStep(survey, surveyId) {
  const ticketPlanQuestion = getSessionTicketPlanQuestion(survey);
  const selectedPlan = getSessionTicketPlanSelection(surveyId);
  answerPanel.innerHTML = `
    ${renderPendingNotice()}
    <div class="section-head survey-toolbar">
      <div>
        <h2>${escapeHtml(survey.title)}</h2>
        <p>回数券の種類を選択してください。</p>
      </div>
      <button id="backToSessionTypeButton" class="ghost-button" type="button">戻る</button>
    </div>
    <div class="question-list">
      <div class="question-block prefilled-answer">
        <strong>施術内容</strong>
        <div>${escapeHtml(getSessionTypeSelection(surveyId))}</div>
      </div>
      <div class="question-block">
        <strong>${escapeHtml(ticketPlanQuestion.label)}</strong>
        <div class="selection-grid">
          ${ticketPlanQuestion.options
            .map(
              (option) => `
                <button
                  class="selection-button ${selectedPlan === option ? "active" : ""}"
                  type="button"
                  data-session-ticket-plan-option="${escapeHtml(option)}"
                >
                  ${escapeHtml(option)}
                </button>
              `,
            )
            .join("")}
        </div>
      </div>
    </div>
  `;

  document.querySelector("#backToSessionTypeButton").addEventListener("click", () => {
    clearSessionSelections(surveyId);
    renderAnswerPanel();
  });
  document.querySelectorAll("[data-session-ticket-plan-option]").forEach((button) => {
    button.addEventListener("click", () => {
      const nextPlan = normalizeText(button.dataset.sessionTicketPlanOption);
      if (!nextPlan) return;
      updateDraftValue(surveyId, SESSION_TICKET_PLAN_QUESTION_ID, nextPlan);
      clearSessionTicketProgressSelections(surveyId);
      renderAnswerPanel();
    });
  });
  attachCommonButtons();
}

function renderSessionTicketSheetStep(survey, surveyId) {
  const ticketSheetQuestion = getSessionTicketSheetQuestion(survey);
  const selectedPlan = getSessionTicketPlanSelection(surveyId);
  const selectedValue = getSessionTicketSheetSelection(surveyId);
  answerPanel.innerHTML = `
    ${renderPendingNotice()}
    <div class="section-head survey-toolbar">
      <div>
        <h2>${escapeHtml(survey.title)}</h2>
        <p>${escapeHtml(selectedPlan)} の何枚目かを選択してください。</p>
      </div>
      <button id="backToSessionPlanButton" class="ghost-button" type="button">戻る</button>
    </div>
    <form id="sessionTicketSheetForm" class="question-list">
      <div class="question-block prefilled-answer">
        <strong>施術内容</strong>
        <div>${escapeHtml(getSessionTypeSelection(surveyId))}</div>
        <strong>回数券の種類</strong>
        <div>${escapeHtml(selectedPlan)}</div>
      </div>
      <fieldset class="question-block">
        <legend><span>${escapeHtml(ticketSheetQuestion.label)}</span></legend>
        <div class="choice-row">
          ${ticketSheetQuestion.options
            .map(
              (option) => `
                <label>
                  <input type="radio" name="sessionTicketSheet" value="${escapeHtml(option)}" ${selectedValue === option ? "checked" : ""} />
                  ${escapeHtml(option)}
                </label>
              `,
            )
            .join("")}
        </div>
      </fieldset>
      <button class="primary-button" type="submit">次へ</button>
    </form>
  `;

  document.querySelector("#backToSessionPlanButton").addEventListener("click", () => {
    updateDraftValue(surveyId, SESSION_TICKET_PLAN_QUESTION_ID, "");
    clearSessionTicketProgressSelections(surveyId);
    renderAnswerPanel();
  });
  document.querySelector("#sessionTicketSheetForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const ticketSheet = normalizeText(formData.get("sessionTicketSheet"));
    if (!ticketSheet) {
      showToast("回数券の何枚目かを選択してください。");
      return;
    }
    updateDraftValue(surveyId, SESSION_TICKET_SHEET_QUESTION_ID, ticketSheet);
    updateDraftValue(surveyId, SESSION_TICKET_ROUND_QUESTION_ID, "");
    renderAnswerPanel();
  });
  attachCommonButtons();
}

function renderSessionTicketRoundStep(survey, surveyId) {
  const ticketRoundQuestion = getSessionTicketRoundQuestion(survey);
  const selectedPlan = getSessionTicketPlanSelection(surveyId);
  const selectedSheet = getSessionTicketSheetSelection(surveyId);
  const selectedValue = getSessionTicketRoundSelection(surveyId);
  const roundOptions = getTicketRoundOptions(selectedPlan);
  const expectedRoundLabel = getExpectedSessionTicketRoundLabel(selectedPlan, selectedSheet, {
    excludeResponseId: appState.editingResponseId,
  });

  answerPanel.innerHTML = `
    ${renderPendingNotice()}
    <div class="section-head survey-toolbar">
      <div>
        <h2>${escapeHtml(survey.title)}</h2>
        <p>${escapeHtml(selectedPlan)} の何回目かを選択してください。</p>
      </div>
      <button id="backToSessionPlanButton" class="ghost-button" type="button">戻る</button>
    </div>
    <form id="sessionTicketRoundForm" class="question-list">
      <div class="question-block prefilled-answer">
        <strong>施術内容</strong>
        <div>${escapeHtml(getSessionTypeSelection(surveyId))}</div>
        <strong>回数券の種類</strong>
        <div>${escapeHtml(selectedPlan)}</div>
        <strong>ホーム画面の枚数記録</strong>
        <div>${escapeHtml(selectedSheet)}</div>
        ${
          expectedRoundLabel
            ? `
              <strong>スタンプカード上の想定回数</strong>
              <div>${escapeHtml(expectedRoundLabel)}</div>
            `
            : ""
        }
      </div>
      <fieldset class="question-block">
        <legend><span>${escapeHtml(ticketRoundQuestion.label)}</span></legend>
        <div class="choice-row">
          ${roundOptions
            .map(
              (option) => `
                <label>
                  <input type="radio" name="sessionTicketRound" value="${escapeHtml(option)}" ${selectedValue === option ? "checked" : ""} />
                  ${escapeHtml(option)}
                </label>
              `,
            )
            .join("")}
        </div>
      </fieldset>
      <button class="primary-button" type="submit">質問へ進む</button>
    </form>
  `;

  document.querySelector("#backToSessionPlanButton").addEventListener("click", () => {
    updateDraftValue(surveyId, SESSION_TICKET_PLAN_QUESTION_ID, "");
    updateDraftValue(surveyId, SESSION_TICKET_SHEET_QUESTION_ID, "");
    updateDraftValue(surveyId, SESSION_TICKET_ROUND_QUESTION_ID, "");
    renderAnswerPanel();
  });
  document.querySelectorAll('input[name="sessionTicketRound"]').forEach((input) => {
    input.addEventListener("change", () => {
      const message = getSessionTicketRoundMismatchMessage(
        selectedPlan,
        selectedSheet,
        normalizeText(input.value),
        { excludeResponseId: appState.editingResponseId },
      );
      if (message) {
        showToast(message);
      }
    });
  });
  document.querySelector("#sessionTicketRoundForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const ticketRound = normalizeText(formData.get("sessionTicketRound"));
    if (!ticketRound) {
      showToast("回数券の何回目ですか？を選択してください。");
      return;
    }
    const mismatchMessage = getSessionTicketRoundMismatchMessage(selectedPlan, selectedSheet, ticketRound, {
      excludeResponseId: appState.editingResponseId,
    });
    if (mismatchMessage) {
      showToast(mismatchMessage);
      return;
    }
    updateDraftValue(surveyId, SESSION_TICKET_ROUND_QUESTION_ID, ticketRound);
    renderAnswerPanel();
  });
  attachCommonButtons();
}

function renderTicketStepCount(survey, surveyId) {
  const ticketQuestion = getTicketEndCountQuestion(survey);
  const selectedValue = getTicketEndCountSelection(surveyId);
  answerPanel.innerHTML = `
    ${renderPendingNotice()}
    <div class="section-head survey-toolbar">
      <div>
        <h2>${escapeHtml(survey.title)}</h2>
        <p>最初に回数券の種類を選択してください。</p>
      </div>
      <button id="backToHomeButton" class="ghost-button" type="button">一覧へ戻る</button>
    </div>
    <form id="ticketCountForm" class="question-list">
      <fieldset class="question-block">
        <legend><span>${escapeHtml(ticketQuestion.label)}</span></legend>
        <div class="choice-row">
          ${ticketQuestion.options
            .map(
              (option) => `
                <label>
                  <input type="radio" name="ticketCount" value="${escapeHtml(option)}" ${selectedValue === option ? "checked" : ""} />
                  ${escapeHtml(option)}
                </label>
              `,
            )
            .join("")}
        </div>
      </fieldset>
      <button class="primary-button" type="submit">次へ</button>
    </form>
  `;

  document.querySelector("#backToHomeButton").addEventListener("click", () => setPage("home"));
  document.querySelector("#ticketCountForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const ticketCount = normalizeText(formData.get("ticketCount"));
    if (!ticketCount) {
      showToast("回数券の種類を選択してください。");
      return;
    }
    updateDraftValue(surveyId, TICKET_END_COUNT_QUESTION_ID, ticketCount);
    updateDraftValue(surveyId, TICKET_END_SHEET_QUESTION_ID, "");
    updateDraftValue(surveyId, TICKET_END_ROUND_QUESTION_ID, "");
    renderAnswerPanel();
  });
  attachCommonButtons();
}

function renderTicketStepSheet(survey, surveyId) {
  const ticketSheetQuestion = getTicketEndSheetQuestion(survey);
  const selectedTicketCount = getTicketEndCountSelection(surveyId);
  const selectedValue = getTicketEndSheetSelection(surveyId);
  answerPanel.innerHTML = `
    ${renderPendingNotice()}
    <div class="section-head survey-toolbar">
      <div>
        <h2>${escapeHtml(survey.title)}</h2>
        <p>${escapeHtml(selectedTicketCount)} の何枚目かを選択してください。</p>
      </div>
      <button id="backToTicketCountButton" class="ghost-button" type="button">戻る</button>
    </div>
    <form id="ticketSheetForm" class="question-list">
      <div class="question-block prefilled-answer">
        <strong>回数券の種類</strong>
        <div>${escapeHtml(selectedTicketCount)}</div>
      </div>
      <fieldset class="question-block">
        <legend><span>${escapeHtml(ticketSheetQuestion.label)}</span></legend>
        <div class="choice-row">
          ${ticketSheetQuestion.options
            .map(
              (option) => `
                <label>
                  <input type="radio" name="ticketSheet" value="${escapeHtml(option)}" ${selectedValue === option ? "checked" : ""} />
                  ${escapeHtml(option)}
                </label>
              `,
            )
            .join("")}
        </div>
      </fieldset>
      <button class="primary-button" type="submit">次へ</button>
    </form>
  `;

  document.querySelector("#backToTicketCountButton").addEventListener("click", () => {
    clearTicketEndSelections(surveyId);
    renderAnswerPanel();
  });
  document.querySelector("#ticketSheetForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const ticketSheet = normalizeText(formData.get("ticketSheet"));
    if (!ticketSheet) {
      showToast("回数券の何枚目かを選択してください。");
      return;
    }
    updateDraftValue(surveyId, TICKET_END_SHEET_QUESTION_ID, ticketSheet);
    updateDraftValue(surveyId, TICKET_END_ROUND_QUESTION_ID, "");
    renderAnswerPanel();
  });
  attachCommonButtons();
}

function renderTicketStepRound(survey, surveyId) {
  const ticketRoundQuestion = getTicketEndRoundQuestion(survey);
  const selectedTicketCount = getTicketEndCountSelection(surveyId);
  const selectedTicketSheet = getTicketEndSheetSelection(surveyId);
  const selectedValue = getTicketEndRoundSelection(surveyId);
  const roundOptions = getTicketRoundOptions(selectedTicketCount);

  answerPanel.innerHTML = `
    ${renderPendingNotice()}
    <div class="section-head survey-toolbar">
      <div>
        <h2>${escapeHtml(survey.title)}</h2>
        <p>${escapeHtml(selectedTicketCount)} の ${escapeHtml(selectedTicketSheet)} の何回目かを選択してください。</p>
      </div>
      <button id="backToTicketSheetButton" class="ghost-button" type="button">戻る</button>
    </div>
    <form id="ticketRoundForm" class="question-list">
      <div class="question-block prefilled-answer">
        <strong>回数券の種類</strong>
        <div>${escapeHtml(selectedTicketCount)}</div>
        <strong>回数券の何枚目</strong>
        <div>${escapeHtml(selectedTicketSheet)}</div>
      </div>
      <fieldset class="question-block">
        <legend><span>${escapeHtml(ticketRoundQuestion.label)}</span></legend>
        <div class="choice-row">
          ${roundOptions
            .map(
              (option) => `
                <label>
                  <input type="radio" name="ticketRound" value="${escapeHtml(option)}" ${selectedValue === option ? "checked" : ""} />
                  ${escapeHtml(option)}
                </label>
              `,
            )
            .join("")}
        </div>
      </fieldset>
      <button class="primary-button" type="submit">質問へ進む</button>
    </form>
  `;

  document.querySelector("#backToTicketSheetButton").addEventListener("click", () => {
    updateDraftValue(surveyId, TICKET_END_SHEET_QUESTION_ID, "");
    updateDraftValue(surveyId, TICKET_END_ROUND_QUESTION_ID, "");
    renderAnswerPanel();
  });
  document.querySelector("#ticketRoundForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const ticketRound = normalizeText(formData.get("ticketRound"));
    if (!ticketRound) {
      showToast("回数券の何回目ですか？を選択してください。");
      return;
    }
    updateDraftValue(surveyId, TICKET_END_ROUND_QUESTION_ID, ticketRound);
    renderAnswerPanel();
  });
  attachCommonButtons();
}

function renderQuestion(question, index, surveyId) {
  const draft = getSurveyDraft(surveyId);
  const name = `question-${question.id}`;
  const selectedCheckboxValues = Array.isArray(draft.values[question.id]) ? draft.values[question.id] : [];
  const label = renderQuestionLabel(question, index, surveyId);

  if (question.type === "textarea") {
    return `
      <label class="question-block" data-question-wrap="${question.id}">
        ${label}
        <textarea name="${name}" data-question-id="${question.id}">${escapeHtml(draft.values[question.id] || "")}</textarea>
      </label>
    `;
  }

  if (question.type === "rating") {
    return `
      <div class="question-block" data-question-wrap="${question.id}" role="group">
        <div class="question-legend">${label}</div>
        <div class="rating-row">
          ${[1, 2, 3, 4, 5]
            .map(
              (value) => `
                <label>
                  <input type="radio" name="${name}" value="${value}" data-question-id="${question.id}" ${String(draft.values[question.id] || "") === String(value) ? "checked" : ""} />
                  <span class="option-text">${value}</span>
                </label>
              `,
            )
            .join("")}
        </div>
      </div>
    `;
  }

  if (question.type === "choice") {
    return `
      <div class="question-block" data-question-wrap="${question.id}" role="group">
        <div class="question-legend">${label}</div>
        <div class="choice-row">
          ${question.options
            .map(
              (option) => `
                <label>
                  <input type="radio" name="${name}" value="${escapeHtml(option)}" data-question-id="${question.id}" ${String(draft.values[question.id] || "") === option ? "checked" : ""} />
                  <span class="option-text">${escapeHtml(option)}</span>
                </label>
              `,
            )
            .join("")}
        </div>
      </div>
    `;
  }

  if (question.type === "checkbox") {
    const selected = new Set(selectedCheckboxValues);
    if (question.id === SESSION_CONCERN_QUESTION_ID) {
      const activeCategoryId = getConcernActiveCategory(surveyId, question.id, selectedCheckboxValues);
      return `
        <div class="question-block" data-question-wrap="${question.id}" role="group">
          <div class="question-legend">${label}</div>
          <div class="question-caption">気になるカテゴリを選んでから、該当する詳細項目にチェックしてください。</div>
          <div class="concern-category-list">
            ${SESSION_CONCERN_CATEGORIES.map((category) => {
              const isActive = category.id === activeCategoryId;
              return `
                <div class="concern-category-card ${isActive ? "active" : ""}">
                  <button
                    class="selection-button ${isActive ? "active" : ""}"
                    type="button"
                    data-concern-category="${escapeHtml(category.id)}"
                    data-question-id="${question.id}"
                  >
                    ${escapeHtml(category.label)}
                  </button>
                  ${
                    isActive
                      ? `
                        <div class="checkbox-row concern-option-list">
                          ${category.options
                            .map(
                              (option) => `
                                <label>
                                  <input type="checkbox" name="${name}" value="${escapeHtml(option)}" data-question-id="${question.id}" ${selected.has(option) ? "checked" : ""} />
                                  <span class="option-text">${escapeHtml(option)}</span>
                                </label>
                              `,
                            )
                            .join("")}
                        </div>
                      `
                      : ""
                  }
                </div>
              `;
            }).join("")}
          </div>
          <div class="checkbox-row concern-option-list">
            <div class="checkbox-option-card">
              <label>
                <input
                  type="checkbox"
                  name="${name}"
                  value="${escapeHtml(SESSION_CONCERN_OTHER_OPTION)}"
                  data-question-id="${question.id}"
                  ${selected.has(SESSION_CONCERN_OTHER_OPTION) ? "checked" : ""}
                />
                <span class="option-text">${escapeHtml(SESSION_CONCERN_OTHER_OPTION)}</span>
              </label>
              ${selected.has(SESSION_CONCERN_OTHER_OPTION)
                ? renderInlineOtherTextarea(
                    SESSION_CONCERN_OTHER_QUESTION_ID,
                    surveyId,
                    "その他（長文）の内容を入力してください",
                  )
                : ""}
            </div>
          </div>
          ${activeCategoryId ? "" : `<div class="empty">カテゴリを選択すると詳細項目を表示します。</div>`}
        </div>
      `;
    }
    return `
      <div class="question-block" data-question-wrap="${question.id}" role="group">
        <div class="question-legend">${label}</div>
        <div class="checkbox-row">
          ${question.options
            .map(
              (option) => option === "その他（自由記述）"
                ? `
                    <div class="checkbox-option-card">
                      <label>
                        <input type="checkbox" name="${name}" value="${escapeHtml(option)}" data-question-id="${question.id}" ${selected.has(option) ? "checked" : ""} />
                        <span class="option-text">${escapeHtml(option)}</span>
                      </label>
                      ${selected.has(option)
                        ? renderInlineOtherTextarea(
                            SESSION_LIFE_CHANGES_OTHER_QUESTION_ID,
                            surveyId,
                            "その他（自由記述）の内容を入力してください",
                          )
                        : ""}
                    </div>
                  `
                : `
                    <label>
                      <input type="checkbox" name="${name}" value="${escapeHtml(option)}" data-question-id="${question.id}" ${selected.has(option) ? "checked" : ""} />
                      <span class="option-text">${escapeHtml(option)}</span>
                    </label>
                  `,
            )
            .join("")}
        </div>
      </div>
    `;
  }

  if (question.type === "photo") {
    const files = getDraftPhotos(surveyId, question.id);
    const maxFiles = getPhotoQuestionFileLimit(question);
    const requiredPhotoCount = getPhotoQuestionRequiredCount(question, surveyId);
    return `
      <div class="question-block" data-question-wrap="${question.id}">
        <div class="question-label">${label}</div>
        <input type="file" name="${name}" accept="image/*" multiple hidden data-photo-input="${question.id}" />
        <div class="photo-actions">
          <button class="secondary-button" type="button" data-photo-pick="${question.id}" data-photo-mode="append">
            ${files.length ? "写真を追加" : "写真を選ぶ"}
          </button>
          ${files.length ? `<button class="ghost-button" type="button" data-photo-pick="${question.id}" data-photo-mode="replace">撮り直す</button>` : ""}
        </div>
        <span class="field-help">スマホ内の写真を選択してください。${requiredPhotoCount > 1 ? `${requiredPhotoCount}枚必須、` : ""}最大 ${maxFiles} 枚まで添付できます。</span>
        ${
          files.length
            ? `
              <div class="photo-list editable">
                ${files
                  .map((file, fileIndex) => {
                    const preview = getPhotoPreviewSrc(file);
                    return `
                      <div class="photo-thumb editable-thumb">
                        ${preview ? `<img src="${escapeHtml(preview)}" alt="${escapeHtml(file.name || "photo")}" />` : ""}
                        <span>${escapeHtml(file.name || `写真${fileIndex + 1}`)}</span>
                        ${formatPhotoCapturedAt(file.capturedAt) ? `<span class="meta">撮影日: ${escapeHtml(formatPhotoCapturedAt(file.capturedAt))}</span>` : ""}
                        <button class="ghost-button small-button" type="button" data-photo-remove="${question.id}" data-photo-index="${fileIndex}">
                          削除
                        </button>
                      </div>
                    `;
                  })
                  .join("")}
              </div>
            `
            : `<div class="empty">まだ写真は追加されていません。</div>`
        }
      </div>
    `;
  }

  return `
    <label class="question-block" data-question-wrap="${question.id}">
      ${label}
      <input type="text" name="${name}" data-question-id="${question.id}" value="${escapeHtml(draft.values[question.id] || "")}" />
    </label>
  `;
}

function renderSummaryValue(answer) {
  if (Array.isArray(answer.files) && answer.files.length) {
    return `
      <div class="photo-list">
        ${answer.files
          .map((file) => {
            const preview = getPhotoPreviewSrc(file);
            return `
              <div class="photo-thumb">
                ${preview ? `<img src="${escapeHtml(preview)}" alt="${escapeHtml(file.name || "photo")}" />` : ""}
                <span>${escapeHtml(file.name || "写真")}</span>
                ${formatPhotoCapturedAt(file.capturedAt) ? `<span class="meta">撮影日: ${escapeHtml(formatPhotoCapturedAt(file.capturedAt))}</span>` : ""}
              </div>
            `;
          })
          .join("")}
      </div>
    `;
  }
  return escapeHtml(
    Array.isArray(answer.value) ? answer.value.join(", ") : answer.value || "未回答",
  );
}

function renderConfirmPanel(survey) {
  const prepared = appState.confirmPayload;
  if (!prepared) {
    appState.panelMode = "form";
    renderAnswerPanel();
    return;
  }

  answerPanel.innerHTML = `
    <div class="section-head survey-toolbar">
      <div>
        <h2>${escapeHtml(survey.title)}</h2>
        <p>内容を確認して送信してください。</p>
      </div>
      <span class="badge">${appState.editingResponseId ? "回答修正" : "送信前確認"}</span>
    </div>
    <div class="confirm-list">
      ${prepared.summary
        .map(
          (answer) => `
            <article class="question-block">
              <strong>${escapeHtml(answer.label)}</strong>
              <div>${renderSummaryValue(answer)}</div>
            </article>
          `,
        )
        .join("")}
    </div>
    <div class="action-row">
      <button id="backToFormButton" class="secondary-button" type="button">修正する</button>
      <button id="confirmSubmitButton" class="primary-button" type="button">${appState.editingResponseId ? "更新する" : "送信する"}</button>
    </div>
  `;

  document.querySelector("#backToFormButton").addEventListener("click", () => {
    appState.panelMode = "form";
    renderAnswerPanel();
  });
  document.querySelector("#confirmSubmitButton").addEventListener("click", () => {
    void submitPreparedAnswer();
  });
  attachCommonButtons();
}

function renderCompletionPanel(survey) {
  const submittedAt = appState.lastSubmittedResponse?.submittedAt;
  answerPanel.innerHTML = `
    <section class="completion-card">
      <div class="eyebrow">${appState.lastSubmissionWasEdit ? "Updated" : "Completed"}</div>
      <h2>${escapeHtml(survey.title)}</h2>
      <p>${escapeHtml(
        appState.lastSubmissionWasEdit
          ? "回答を更新しました。"
          : survey.completionMessage || "ご回答ありがとうございました。"
      )}</p>
      ${
        submittedAt
          ? `<div class="meta">送信日時: ${formatDate(submittedAt)}</div>`
          : ""
      }
      <div class="action-row">
        <button id="completionHistoryButton" class="secondary-button" type="button">履歴を見る</button>
        <button id="completionHomeButton" class="primary-button" type="button">ホームへ戻る</button>
      </div>
    </section>
  `;

  document.querySelector("#completionHistoryButton").addEventListener("click", () => {
    setPage("history");
  });
  document.querySelector("#completionHomeButton").addEventListener("click", () => {
    appState.panelMode = "form";
    appState.lastSubmittedResponse = null;
    appState.editingResponseId = "";
    appState.lastSubmissionWasEdit = false;
    setPage("home");
    renderSurveys();
    renderAnswerPanel();
  });
  attachCommonButtons();
}

function renderFormPanel(survey) {
  const surveyId = survey.id;
  const draft = getSurveyDraft(surveyId);
  const availability = getSurveyAvailability(survey);
  const visibleQuestions = getVisibleQuestions(survey, draft);
  const renderableQuestions = visibleQuestions.filter((question) => !isInlineManagedTextareaQuestion(question));
  const progress = getProgress(survey, draft);

  answerPanel.innerHTML = `
    ${renderPendingNotice()}
    <div class="section-head survey-toolbar">
      <div>
        <h2>${escapeHtml(survey.title)}</h2>
        <p>${escapeHtml(survey.description)}</p>
      </div>
      <button id="backToHomeButton" class="ghost-button" type="button">一覧へ戻る</button>
    </div>
    <div class="notice-row">
      <span class="badge ${availability.open ? "open" : "closed"} ${availability.nearDeadline ? "warn" : ""}">${escapeHtml(availability.label)}</span>
      <span class="meta">${escapeHtml(availability.detail)}</span>
    </div>
    ${
      survey.introMessage
        ? `
          <article class="policy-card intro-card">
            <strong>ご案内</strong>
            <p>${escapeHtml(survey.introMessage)}</p>
          </article>
        `
        : ""
    }
    ${
      availability.nearDeadline
        ? `
          <article class="policy-card deadline-card">
            <strong>回答期限が近づいています</strong>
            <p>${escapeHtml(availability.detail)}</p>
          </article>
        `
        : ""
    }
    ${renderProgressBar(progress)}
    ${renderDraftStatusCard(surveyId)}
    ${
      isSessionSurvey(survey)
        ? `
          <div class="question-block prefilled-answer">
            <strong>施術内容</strong>
            <div>${escapeHtml(getSessionTypeSelection(surveyId))}</div>
            ${
              getSessionTypeSelection(surveyId) === "回数券"
                ? `
                  <strong>回数券の種類</strong>
                  <div>${escapeHtml(getSessionTicketPlanSelection(surveyId))}</div>
                  <strong>ホーム画面の枚数記録</strong>
                  <div>${escapeHtml(getSessionTicketSheetSelection(surveyId))}</div>
                  <strong>回数券の何回目</strong>
                  <div>${escapeHtml(getSessionTicketRoundSelection(surveyId))}</div>
                `
                : ""
            }
            <button id="changeSessionTypeButton" class="ghost-button" type="button">変更する</button>
          </div>
        `
        : ""
    }
    ${
      !availability.open
        ? `<div class="empty">このアンケートは現在回答を受け付けていません。</div>`
        : ""
    }
    <form id="answerForm" class="question-list">
      ${renderableQuestions.map((question, index) => renderQuestion(question, index, surveyId)).join("")}
      ${
        appState.publicInfo.requireConsent
          ? `
            <div class="question-block consent-block" data-question-wrap="${getAgreementDraftKey(surveyId)}">
              <span>同意確認</span>
              <label class="inline-check">
                <input
                  type="checkbox"
                  name="agreement"
                  data-agreement-input="true"
                  ${isAgreementAccepted(surveyId) ? "checked" : ""}
                />
                <span>${escapeHtml(appState.publicInfo.consentText || "回答内容と添付写真の利用に同意します。")}</span>
              </label>
            </div>
          `
          : ""
      }
      ${
        renderableQuestions.length
          ? `<button class="primary-button" type="submit">${appState.editingResponseId ? "確認へ進む" : "内容を確認する"}</button>`
          : `<div class="empty">表示できる質問がありません。</div>`
      }
    </form>
    <article class="policy-card">
      <strong>取り扱いについて</strong>
      <p>${escapeHtml(appState.publicInfo.dataPolicyText || "ご回答内容は院内管理のために利用します。")}</p>
      <div class="meta">Version ${escapeHtml(appState.publicInfo.version || APP_VERSION)}</div>
    </article>
  `;

  document.querySelector("#backToHomeButton").addEventListener("click", () => setPage("home"));
  document.querySelector("#clearDraftButton")?.addEventListener("click", () => {
    clearSurveyDraft(surveyId);
    appState.confirmPayload = null;
    if (appState.editingResponseId) {
      appState.editingResponseId = "";
    }
    showToast("下書きを削除しました。");
    renderSurveys();
    renderAnswerPanel();
  });
  document.querySelector("#changeSessionTypeButton")?.addEventListener("click", () => {
    clearSessionSelections(surveyId);
    renderAnswerPanel();
  });

  if (availability.open) {
    attachAnswerFormHandlers(document.querySelector("#answerForm"), survey);
  }
  attachCommonButtons();
}

function renderAnswerPanel() {
  const survey = getSelectedSurvey();
  if (!survey) {
    answerPanel.innerHTML = `
      ${renderPendingNotice()}
      <div class="empty">ホームに戻って、回答したいアンケートを選択してください。</div>
    `;
    attachCommonButtons();
    return;
  }

  const surveyId = survey.id;
  const draft = getSurveyDraft(surveyId);

  if (appState.panelMode === "confirm") {
    renderConfirmPanel(survey);
    return;
  }
  if (appState.panelMode === "complete" && appState.lastSubmittedResponse?.surveyId === surveyId) {
    renderCompletionPanel(survey);
    return;
  }

  if (isSessionSurvey(survey) && !getSessionTypeSelection(surveyId)) {
    renderSessionTypeStep(survey, surveyId);
    return;
  }
  if (
    isSessionSurvey(survey) &&
    getSessionTypeSelection(surveyId) === "回数券" &&
    !getSessionTicketPlanSelection(surveyId)
  ) {
    renderSessionTicketPlanStep(survey, surveyId);
    return;
  }
  if (
    isSessionSurvey(survey) &&
    getSessionTypeSelection(surveyId) === "回数券" &&
    !getSessionTicketRoundSelection(surveyId)
  ) {
    const selectedSheet = ensureSessionTicketSheetSelection(surveyId);
    if (!selectedSheet && hasCustomerSession() && appState.historyLoading) {
      answerPanel.innerHTML = `
        ${renderPendingNotice()}
        <div class="section-head survey-toolbar">
          <div>
            <h2>${escapeHtml(survey.title)}</h2>
            <p>ホーム画面の枚数記録を確認しています。</p>
          </div>
          <button id="backToSessionPlanLoadingButton" class="ghost-button" type="button">戻る</button>
        </div>
        <div class="empty">回数券の現在の枚数を読み込み中です。</div>
      `;
      document.querySelector("#backToSessionPlanLoadingButton").addEventListener("click", () => {
        updateDraftValue(surveyId, SESSION_TICKET_PLAN_QUESTION_ID, "");
        updateDraftValue(surveyId, SESSION_TICKET_SHEET_QUESTION_ID, "");
        updateDraftValue(surveyId, SESSION_TICKET_ROUND_QUESTION_ID, "");
        renderAnswerPanel();
      });
      attachCommonButtons();
      return;
    }
    renderSessionTicketRoundStep(survey, surveyId);
    return;
  }

  renderFormPanel(survey, draft);
}

function attachAnswerFormHandlers(form, survey) {
  if (!form) return;

  form.addEventListener("input", (event) => {
    const input = event.target;
    if (input.dataset.agreementInput) {
      setAgreementAccepted(survey.id, input.checked);
      refreshDraftStatusDisplay(survey.id);
      return;
    }
    const questionId = input.dataset.questionId;
    if (!questionId) return;

    if (input.type === "checkbox") {
      const values = Array.from(form.querySelectorAll(`[data-question-id="${questionId}"]:checked`)).map(
        (node) => node.value,
      );
      updateDraftValue(survey.id, questionId, values);
      if (doesQuestionAffectVisibility(survey, questionId)) {
        renderAnswerPanel();
        return;
      }
      refreshProgressDisplay(survey);
      return;
    }

    if (input.type === "radio") {
      updateDraftValue(survey.id, questionId, input.value);
      renderAnswerPanel();
      return;
    }

    updateDraftValue(survey.id, questionId, input.value);
    refreshProgressDisplay(survey);
  });

  form.querySelectorAll("[data-photo-pick]").forEach((button) => {
    button.addEventListener("click", () => {
      const questionId = button.dataset.photoPick;
      const input = form.querySelector(`[data-photo-input="${questionId}"]`);
      if (!input) return;
      input.dataset.photoMode = button.dataset.photoMode || "append";
      input.click();
    });
  });

  form.querySelectorAll("[data-concern-category]").forEach((button) => {
    button.addEventListener("click", () => {
      const questionId = button.dataset.questionId;
      const categoryId = button.dataset.concernCategory;
      if (!questionId || !categoryId) return;
      setConcernActiveCategory(survey.id, questionId, categoryId);
      renderAnswerPanel();
    });
  });

  form.querySelectorAll("[data-photo-input]").forEach((input) => {
    input.addEventListener("change", async () => {
      const questionId = input.dataset.photoInput;
      const files = Array.from(input.files || []);
      if (!files.length) return;

      const question = survey.questions.find((item) => item.id === questionId);
      const maxFiles = getPhotoQuestionFileLimit(question);
      const existing = getDraftPhotos(survey.id, questionId);
      const nextFiles = input.dataset.photoMode === "replace" ? [] : existing.slice();

      try {
        for (const file of files.slice(0, maxFiles)) {
          nextFiles.push(await preparePhotoFile(file));
        }
        updateDraftPhotos(survey.id, questionId, nextFiles.slice(0, maxFiles));
        renderAnswerPanel();
      } catch (error) {
        reportClientError("customer.photo.prepare", error, { questionId });
        showToast(error.message || "写真を追加できませんでした。");
      } finally {
        input.value = "";
      }
    });
  });

  form.querySelectorAll("[data-photo-remove]").forEach((button) => {
    button.addEventListener("click", () => {
      const questionId = button.dataset.photoRemove;
      const index = Number(button.dataset.photoIndex);
      const files = getDraftPhotos(survey.id, questionId).slice();
      files.splice(index, 1);
      updateDraftPhotos(survey.id, questionId, files);
      renderAnswerPanel();
    });
  });

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    try {
      const prepared = prepareSubmissionFromDraft(survey, getSurveyDraft(survey.id));
      appState.confirmPayload = prepared;
      appState.panelMode = "confirm";
      renderAnswerPanel();
    } catch (error) {
      reportClientError("customer.prepareSubmission", error, { surveyId: survey.id });
      scrollToQuestion(error.questionId);
      showToast(error.message || "回答内容を確認できませんでした。");
    }
  });
}

function prepareSubmissionFromDraft(survey, draft) {
  const answerMap = buildDraftAnswerMap(survey, draft);
  const answers = [];
  const summary = [];

  if (isSessionSurvey(survey) && getSessionTypeSelection(survey.id) === "回数券") {
    const mismatchMessage = getSessionTicketRoundMismatchMessage(
      getSessionTicketPlanSelection(survey.id),
      getSessionTicketSheetSelection(survey.id),
      getSessionTicketRoundSelection(survey.id),
      { excludeResponseId: appState.editingResponseId },
    );
    if (mismatchMessage) {
      throw makeValidationError(SESSION_TICKET_ROUND_QUESTION_ID, mismatchMessage);
    }
  }

  for (const question of survey.questions) {
    const visible = getSubmissionQuestionVisibility(question, answerMap, survey.id);
    const required = isQuestionRequired(question, survey.id);

    if (question.type === "photo") {
      const files = visible ? getDraftPhotos(survey.id, question.id) : [];
      const requiredPhotoCount = getPhotoQuestionRequiredCount(question, survey.id);
      const maxFiles = getPhotoQuestionFileLimit(question);
      if (visible && files.length > maxFiles) {
        throw makeValidationError(question.id, `写真は${maxFiles}枚まで添付できます。`);
      }
      if (visible && required && files.length < requiredPhotoCount) {
        throw makeValidationError(
          question.id,
          requiredPhotoCount === 1 ? "未回答の質問があります。" : `写真を${requiredPhotoCount}枚添付してください。`,
        );
      }
      answers.push({ questionId: question.id, files: cloneData(files) });
      if (visible) {
        summary.push({
          questionId: question.id,
          label: getQuestionLabel(question, survey.id),
          files,
        });
      }
      continue;
    }

    let value = draft.values[question.id];
    if (question.type === "checkbox") {
      value = Array.isArray(value) ? value.map(normalizeText).filter(Boolean) : [];
      if (visible && required && !value.length) {
        throw makeValidationError(question.id, "未回答の質問があります。");
      }
      if (visible && value.some((item) => !question.options.includes(item))) {
        throw makeValidationError(question.id, "選択肢から回答してください。");
      }
      answers.push({ questionId: question.id, value: visible ? value : [] });
      if (visible) {
        summary.push({
          questionId: question.id,
          label: getQuestionLabel(question, survey.id),
          value,
        });
      }
      continue;
    }

    value = normalizeText(value);
    if (visible && required && !value) {
      throw makeValidationError(question.id, "未回答の質問があります。");
    }
    if (visible && question.type === "rating" && value && !["1", "2", "3", "4", "5"].includes(value)) {
      throw makeValidationError(question.id, "5段階評価の選択が正しくありません。");
    }
    if (visible && question.type === "choice" && value && !question.options.includes(value)) {
      throw makeValidationError(question.id, "選択肢から回答してください。");
    }
    answers.push({ questionId: question.id, value: visible ? value : "" });
    if (visible) {
      summary.push({
        questionId: question.id,
        label: getQuestionLabel(question, survey.id),
        value,
      });
    }
  }

  if (appState.publicInfo.requireConsent && !isAgreementAccepted(survey.id)) {
    throw makeValidationError(
      getAgreementDraftKey(survey.id),
      "同意確認にチェックを入れてください。",
    );
  }

  return { answers, summary };
}

async function submitPreparedAnswer() {
  const survey = getSelectedSurvey();
  if (!survey || !appState.confirmPayload) return;
  if (!hasCustomerSession()) {
    showToast("先にログインしてください。");
    setPage("login");
    return;
  }

  const button = document.querySelector("#confirmSubmitButton");
  if (button) {
    button.disabled = true;
    button.textContent = appState.editingResponseId ? "更新中です" : "送信中です";
  }

  const body = {
    surveyId: survey.id,
    customer: appState.customer,
    answers: appState.confirmPayload.answers,
  };
  const path = appState.editingResponseId
    ? `/api/public/responses/${encodeURIComponent(appState.editingResponseId)}`
    : "/api/public/responses";
  const method = appState.editingResponseId ? "PUT" : "POST";

  try {
    const result = await api.request(path, {
      method,
      body,
    });
    clearPendingSubmission();
    clearSurveyDraft(survey.id);
    appState.confirmPayload = null;
    appState.panelMode = "complete";
    appState.lastSubmittedResponse = result.response || null;
    appState.lastSubmissionWasEdit = method === "PUT";
    appState.editingResponseId = "";
    await loadHistory();
    renderSurveys();
    renderAnswerPanel();
    showToast(method === "PUT" ? "回答を更新しました。" : "回答を送信しました。");
  } catch (error) {
    reportClientError("customer.submit", error, { surveyId: survey.id, method });
    storePendingSubmission({
      savedAt: new Date().toISOString(),
      surveyId: survey.id,
      surveyTitle: survey.title,
      method,
      path,
      body,
    });
    showToast(error.message || "回答を送信できませんでした。");
    renderSurveys();
    renderAnswerPanel();
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = appState.editingResponseId ? "更新する" : "送信する";
    }
  }
}

function canEditResponse(response) {
  return Date.now() - new Date(response.submittedAt).getTime() <= RESPONSE_EDIT_WINDOW_MS;
}

function responseToDraft(response) {
  const draft = { values: {}, photos: {} };
  (Array.isArray(response.answers) ? response.answers : []).forEach((answer) => {
    if (Array.isArray(answer.files) && answer.files.length) {
      draft.photos[answer.questionId] = cloneData(answer.files);
      return;
    }
    const survey = appState.surveys.find((item) => item.id === response.surveyId);
    const question = survey?.questions.find((item) => item.id === answer.questionId);
    if (question?.type === "checkbox") {
      draft.values[answer.questionId] = String(answer.value || "")
        .split(",")
        .map((item) => item.trim())
        .filter(Boolean);
      return;
    }
    draft.values[answer.questionId] = answer.value || "";
  });
  return draft;
}

function startResponseEdit(responseId) {
  const response = appState.history.find((item) => item.id === responseId);
  if (!response) return;
  const survey = appState.surveys.find((item) => item.id === response.surveyId);
  if (!survey) {
    showToast("アンケート定義を読み込んでから再度お試しください。");
    return;
  }
  setSurveyDraft(survey.id, responseToDraft(response));
  appState.selectedSurveyId = survey.id;
  appState.panelMode = "form";
  appState.confirmPayload = null;
  appState.editingResponseId = response.id;
  appState.lastSubmittedResponse = null;
  appState.lastSubmissionWasEdit = false;
  renderSurveys();
  renderAnswerPanel();
  setPage("survey");
}

function renderAnswerValue(answer) {
  if (Array.isArray(answer.files) && answer.files.length) {
    return `
      <div class="photo-list">
        ${answer.files
          .map((file) => {
            const href = getPhotoOpenHref(file);
            const preview = getPhotoPreviewSrc(file);
            return `
              <a class="photo-thumb" href="${escapeHtml(href)}" target="_blank" rel="noopener">
                ${preview ? `<img src="${escapeHtml(preview)}" alt="${escapeHtml(file.name || "photo")}" />` : ""}
                <span>${escapeHtml(file.name || "写真")}</span>
                ${formatPhotoCapturedAt(file.capturedAt) ? `<span class="meta">撮影日: ${escapeHtml(formatPhotoCapturedAt(file.capturedAt))}</span>` : ""}
              </a>
            `;
          })
          .join("")}
      </div>
    `;
  }
  return escapeHtml(answer.value || "未回答");
}

function renderTicketStampList(ticketInfo) {
  if (!ticketInfo.length) return "";
  return `
    <div class="ticket-stamp-list">
      ${ticketInfo
        .map(
          (item) => `
            <span class="ticket-stamp">
              <span class="ticket-stamp-label">${escapeHtml(item.label)}</span>
              <span>${escapeHtml(item.value)}</span>
            </span>
          `,
        )
        .join("")}
    </div>
  `;
}

function buildTicketInfoFromValues(ticketPlan, ticketSheetLabel, ticketRoundLabel) {
  return [
    {
      label: "回数券",
      value: normalizeText(ticketPlan),
    },
    {
      label: "何枚目",
      value: normalizeText(ticketSheetLabel),
    },
    {
      label: "何回目",
      value: normalizeText(ticketRoundLabel),
    },
  ].filter((item) => item.value);
}

function getAnswerValueFromQuestionIds(answerMap, questionIds) {
  for (const questionId of questionIds) {
    const value = normalizeText(answerMap.get(questionId)?.value);
    if (value) return value;
  }
  return "";
}

function getResponseTicketInfo(response) {
  const answerMap = new Map((response.answers || []).map((answer) => [answer.questionId, answer]));
  return [
    {
      label: "回数券",
      value: getAnswerValueFromQuestionIds(answerMap, [
        SESSION_TICKET_PLAN_QUESTION_ID,
        LEGACY_TICKET_INFO_QUESTION_IDS.size,
      ]),
    },
    {
      label: "何枚目",
      value: getAnswerValueFromQuestionIds(answerMap, [
        SESSION_TICKET_SHEET_QUESTION_ID,
        LEGACY_TICKET_INFO_QUESTION_IDS.sheet,
      ]),
    },
    {
      label: "何回目",
      value: getAnswerValueFromQuestionIds(answerMap, [
        SESSION_TICKET_ROUND_QUESTION_ID,
        LEGACY_TICKET_INFO_QUESTION_IDS.round,
      ]),
    },
  ].filter((item) => item.value);
}

function parseTicketCount(ticketValue) {
  if (ticketValue.includes("10")) return 10;
  if (ticketValue.includes("6")) return 6;
  const matched = ticketValue.match(/\d+/);
  return matched ? Number(matched[0]) : 0;
}

function parseTicketStep(stepValue) {
  const matched = normalizeText(stepValue).match(/\d+/);
  return matched ? Number(matched[0]) : 0;
}

function parseTicketSheet(sheetValue) {
  const matched = normalizeText(sheetValue).match(/\d+/);
  return matched ? Number(matched[0]) : 0;
}

function getNextTicketSheetLabel(sheetValue) {
  const currentSheet = parseTicketSheet(sheetValue);
  return currentSheet ? `${currentSheet + 1}枚目` : "";
}

async function acquireNextTicketSheet(ticketPlan) {
  const nextSheetLabel = getNextTicketSheetLabelForPlan(ticketPlan);
  const nextSheetNumber = parseTicketSheet(nextSheetLabel);
  if (!normalizeText(ticketPlan) || !nextSheetNumber) {
    showToast("新しいスタンプカードを取得できませんでした。");
    return;
  }
  setActiveTicketCardOverride(ticketPlan, nextSheetNumber);
  renderHomeTicketStatus();
  try {
    const result = await api.request("/api/public/customer-profile/ticket-card", {
      method: "POST",
      body: {
        customer: appState.customer,
        ticketCard: {
          plan: ticketPlan,
          sheetNumber: nextSheetNumber,
          round: 0,
        },
      },
    });
    syncCustomerProfileFromServer(result.customerProfile);
    syncActiveTicketCardOverrideFromServer(result.customerProfile);
    showToast(`${ticketPlan} ${nextSheetLabel} を取得しました。`);
  } catch (error) {
    reportClientError("customer.acquireTicketSheet", error, {
      ticketPlan,
      sheetNumber: nextSheetNumber,
    });
    showToast(`${ticketPlan} ${nextSheetLabel} を端末に保存しました。通信後に更新してください。`);
  }
}

function getLatestTicketResponse() {
  return getVisibleHistoryResponses()
    .filter((response) => getResponseTicketInfo(response).length)
    .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))[0] || null;
}

function getActiveTicketCardState() {
  const response = getLatestTicketResponse();
  if (!response) return null;

  const ticketMap = new Map(getResponseTicketInfo(response).map((item) => [item.label, item.value]));
  const ticketPlan = normalizeText(ticketMap.get("回数券") || "");
  const currentSheetLabel = normalizeText(ticketMap.get("何枚目") || "");
  const currentRound = parseTicketStep(ticketMap.get("何回目") || "");
  const ticketCount = parseTicketCount(ticketPlan);
  const activeOverride = getActiveTicketCardOverride();
  const isCurrentCardComplete = Boolean(ticketPlan && ticketCount && currentRound >= ticketCount);

  if (
    activeOverride &&
    isCurrentCardComplete
  ) {
    return {
      response,
      ticketPlan: activeOverride.plan,
      ticketSheetLabel: `${activeOverride.sheetNumber}枚目`,
      currentRound: 0,
      ticketCount: parseTicketCount(activeOverride.plan),
      ticketInfo: buildTicketInfoFromValues(activeOverride.plan, `${activeOverride.sheetNumber}枚目`, "0回目"),
      showAcquireButtons: false,
      acquirePlans: [],
    };
  }

  return {
    response,
    ticketPlan,
    ticketSheetLabel: currentSheetLabel,
    currentRound,
    ticketCount,
    ticketInfo: buildTicketInfoFromValues(ticketPlan, currentSheetLabel, ticketMap.get("何回目") || ""),
    showAcquireButtons: isCurrentCardComplete,
    acquirePlans: isCurrentCardComplete ? ["6回券", "10回券"] : [],
  };
}

function getPhotoPreviewSrc(file) {
  return normalizeText(file?.thumbnailUrl || file?.previewUrl || file?.dataUrl || file?.url || "");
}

function getPhotoOpenHref(file) {
  return normalizeText(
    file?.previewUrl || file?.url || file?.downloadUrl || file?.thumbnailUrl || file?.dataUrl || "#",
  );
}

function renderTicketStampProgress(ticketCount, currentRound) {
  if (!ticketCount) return "";
  const normalizedCount = Math.max(0, Number(ticketCount) || 0);
  const normalizedRound = Math.max(0, Math.min(normalizedCount, Number(currentRound) || 0));
  return `
    <div class="ticket-progress" data-ticket-count="${normalizedCount}">
      ${Array.from({ length: normalizedCount }, (_, index) => {
        const step = index + 1;
        return `
          <span class="stamp-dot ${step <= normalizedRound ? "active" : ""}" aria-label="${step}回目">
            <span class="stamp-dot-step">${step}</span>
            <span class="stamp-dot-unit">回</span>
          </span>
        `;
      }).join("")}
    </div>
  `;
}

function renderHomeTicketStatus() {
  if (!homeTicketStatus) return;
  if (!hasCustomerSession()) {
    homeTicketStatus.innerHTML = `<div class="empty">ログインすると回数券スタンプを表示します。</div>`;
    return;
  }
  if (appState.historyLoading) {
    homeTicketStatus.innerHTML = `<div class="empty">回数券スタンプを読み込み中です。</div>`;
    return;
  }
  if (appState.historyLoadError) {
    homeTicketStatus.innerHTML = `<div class="empty">${escapeHtml(appState.historyLoadError)}</div>`;
    return;
  }

  const activeTicketCard = getActiveTicketCardState();
  if (!activeTicketCard) {
    homeTicketStatus.innerHTML = `<div class="empty">回数券情報のある回答がまだありません。</div>`;
    return;
  }

  homeTicketStatus.innerHTML = `
    <article class="ticket-home-card">
      <div class="ticket-home-head">
        <div>
          <strong>${escapeHtml(getCustomerDisplayName())}</strong>
          <div class="meta">会員番号: ${escapeHtml(getCustomerMemberNumber() || "未発行")}</div>
          <div class="meta">最新更新: ${formatDate(activeTicketCard.response.submittedAt)}</div>
        </div>
        <span class="badge open">${escapeHtml(`${activeTicketCard.currentRound}回目`)}</span>
      </div>
      ${renderTicketStampList(activeTicketCard.ticketInfo)}
      ${
        activeTicketCard.ticketCount
          ? `
            <div class="ticket-progress-card">
              <div class="ticket-progress-head">
                <strong>${escapeHtml(activeTicketCard.ticketSheetLabel || "-")}</strong>
                <span>${activeTicketCard.currentRound} / ${activeTicketCard.ticketCount}</span>
              </div>
              ${renderTicketStampProgress(activeTicketCard.ticketCount, activeTicketCard.currentRound)}
              ${
                activeTicketCard.showAcquireButtons
                  ? `
                    <div class="ticket-next-sheet-action">
                      <div class="meta">新しいスタンプカードを取得</div>
                      <div class="action-row">
                        ${activeTicketCard.acquirePlans
                          .map(
                            (plan) => `
                              <button
                                class="secondary-button"
                                type="button"
                                data-acquire-ticket-plan="${escapeHtml(plan)}"
                              >
                                ${escapeHtml(plan)}を取得
                              </button>
                            `,
                          )
                          .join("")}
                      </div>
                    </div>
                  `
                  : ""
              }
            </div>
          `
          : ""
      }
    </article>
  `;

  homeTicketStatus.querySelectorAll("[data-acquire-ticket-plan]").forEach((button) => {
    button.addEventListener("click", () => {
      void acquireNextTicketSheet(button.dataset.acquireTicketPlan || "");
    });
  });
}

function renderResponseTicketInfo(response) {
  const ticketInfo = getResponseTicketInfo(response);
  if (!ticketInfo.length) return "";
  return `
    <article class="ticket-info-card">
      <strong>回答者情報</strong>
      ${renderTicketStampList(ticketInfo)}
    </article>
  `;
}

function groupHistoryBySurvey() {
  const groups = new Map();
  getVisibleHistoryResponses()
    .slice()
    .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))
    .forEach((response) => {
      const key = normalizeText(response.surveyId || response.surveyTitle || response.id);
      const current = groups.get(key) || {
        surveyId: response.surveyId || key,
        surveyTitle: response.surveyTitle || "アンケート",
        latestAt: response.submittedAt,
        responses: [],
      };
      current.responses.push(response);
      if (new Date(response.submittedAt) > new Date(current.latestAt)) {
        current.latestAt = response.submittedAt;
      }
      groups.set(key, current);
    });
  return Array.from(groups.values()).sort((a, b) => new Date(b.latestAt) - new Date(a.latestAt));
}

function renderHistorySurveyGroups(groups) {
  historyList.innerHTML = `
    ${renderPendingNotice()}
    <div class="history-group-list">
      ${groups
        .map(
          (group) => `
            <button class="history-card history-nav-card" type="button" data-history-survey="${escapeHtml(group.surveyId)}">
              <strong>${escapeHtml(group.surveyTitle)}</strong>
              <div class="meta">回答数: ${group.responses.length}件</div>
              <div class="meta">最新回答: ${formatDate(group.latestAt)}</div>
            </button>
          `,
        )
        .join("")}
    </div>
  `;

  historyList.querySelectorAll("[data-history-survey]").forEach((button) => {
    button.addEventListener("click", () => {
      appState.historySurveyId = button.dataset.historySurvey || "";
      appState.historyResponseId = "";
      renderHistory();
    });
  });
}

function renderHistoryResponseList(group) {
  historyList.innerHTML = `
    ${renderPendingNotice()}
    <div class="section-head">
      <div>
        <strong>${escapeHtml(group.surveyTitle)}</strong>
        <div class="meta">最新順で表示しています。</div>
      </div>
      <button class="ghost-button" type="button" data-history-back="groups">戻る</button>
    </div>
    <div class="history-group-list">
      ${group.responses
        .slice()
        .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))
        .map((response) => {
          const ticketInfo = getResponseTicketInfo(response);
          return `
            <button class="history-card history-nav-card" type="button" data-history-response="${escapeHtml(response.id)}">
              <div class="meta">${formatDate(response.submittedAt)}</div>
              ${
                ticketInfo.length
                  ? `
                    ${renderTicketStampList(ticketInfo)}
                  `
                  : `<div class="meta">この回答に回数券情報はありません。</div>`
              }
            </button>
          `;
        })
        .join("")}
    </div>
  `;

  historyList.querySelector('[data-history-back="groups"]')?.addEventListener("click", () => {
    appState.historySurveyId = "";
    appState.historyResponseId = "";
    renderHistory();
  });
  historyList.querySelectorAll("[data-history-response]").forEach((button) => {
    button.addEventListener("click", () => {
      appState.historyResponseId = button.dataset.historyResponse || "";
      renderHistory();
    });
  });
}

function renderHistoryDetail(group, response) {
  historyList.innerHTML = `
    ${renderPendingNotice()}
    <div class="section-head">
      <div>
        <strong>${escapeHtml(group.surveyTitle)}</strong>
        <div class="meta">${formatDate(response.submittedAt)}</div>
      </div>
      <button class="ghost-button" type="button" data-history-back="responses">戻る</button>
    </div>
    <article class="history-card">
      ${renderResponseTicketInfo(response)}
      ${response.answers
        .map(
          (answer) => `
            <div class="history-answer">
              <strong>${escapeHtml(answer.label)}</strong><br />
              ${renderAnswerValue(answer)}
            </div>
          `,
        )
        .join("")}
      ${
        canEditResponse(response)
          ? `
            <div class="action-row">
              <div class="meta">回答後24時間以内は修正できます。</div>
              <button class="secondary-button" type="button" data-edit-response="${escapeHtml(response.id)}">修正する</button>
            </div>
          `
          : `<div class="meta">修正可能期間を過ぎています。</div>`
      }
    </article>
  `;

  historyList.querySelector('[data-history-back="responses"]')?.addEventListener("click", () => {
    appState.historyResponseId = "";
    renderHistory();
  });
  historyList.querySelectorAll("[data-edit-response]").forEach((button) => {
    button.addEventListener("click", () => startResponseEdit(button.dataset.editResponse));
  });
}

function renderHistory() {
  if (!appState.history.length) {
    historyList.innerHTML = `<div class="empty">まだ回答履歴はありません。</div>`;
    return;
  }
  const groups = groupHistoryBySurvey();
  if (!groups.some((group) => group.surveyId === appState.historySurveyId)) {
    appState.historySurveyId = "";
    appState.historyResponseId = "";
  }

  if (!appState.historySurveyId) {
    renderHistorySurveyGroups(groups);
    attachCommonButtons();
    return;
  }

  const group = groups.find((item) => item.surveyId === appState.historySurveyId);
  if (!group) {
    appState.historySurveyId = "";
    appState.historyResponseId = "";
    renderHistorySurveyGroups(groups);
    attachCommonButtons();
    return;
  }

  if (!group.responses.some((response) => response.id === appState.historyResponseId)) {
    appState.historyResponseId = "";
  }

  if (!appState.historyResponseId) {
    renderHistoryResponseList(group);
    attachCommonButtons();
    return;
  }

  const response = group.responses.find((item) => item.id === appState.historyResponseId);
  if (!response) {
    appState.historyResponseId = "";
    renderHistoryResponseList(group);
    attachCommonButtons();
    return;
  }

  renderHistoryDetail(group, response);
  attachCommonButtons();
}

function getCustomerMeasurements() {
  return appState.measurements
    .filter((measurement) => {
      if (!measurement.customerName) return true;
      return measurement.customerName === appState.customer.name;
    })
    .sort((a, b) => {
      const measuredDiff = new Date(b.measuredAt) - new Date(a.measuredAt);
      if (measuredDiff !== 0) return measuredDiff;
      return new Date(b.updatedAt || b.createdAt || 0) - new Date(a.updatedAt || a.createdAt || 0);
    });
}

function filterMeasurementsByPeriod(measurements, period) {
  if (!Array.isArray(measurements) || !measurements.length || period === "all") return measurements.slice();
  const latest = measurements
    .map((measurement) => new Date(measurement.measuredAt))
    .filter((date) => Number.isFinite(date.getTime()))
    .sort((a, b) => b - a)[0];
  if (!latest) return measurements.slice();
  const threshold = new Date(latest);
  if (period === "1m") threshold.setMonth(threshold.getMonth() - 1);
  else if (period === "6m") threshold.setMonth(threshold.getMonth() - 6);
  else if (period === "1y") threshold.setFullYear(threshold.getFullYear() - 1);
  return measurements.filter((measurement) => new Date(measurement.measuredAt) >= threshold);
}

function buildMeasurementHistoryRows(measurements) {
  const chronological = measurements
    .slice()
    .sort((a, b) => {
      const measuredDiff = new Date(a.measuredAt) - new Date(b.measuredAt);
      if (measuredDiff !== 0) return measuredDiff;
      return new Date(a.createdAt || 0) - new Date(b.createdAt || 0);
    });
  const first = chronological[0] || null;
  const rowsById = new Map();
  chronological.forEach((measurement, index) => {
    const previous = index > 0 ? chronological[index - 1] : null;
    const metrics = {};
    MEASUREMENT_METRICS.forEach((metric) => {
      const value = normalizeMeasurementValue(measurement[metric.key]);
      const previousValue = normalizeMeasurementValue(previous?.[metric.key]);
      const firstValue = normalizeMeasurementValue(first?.[metric.key]);
      metrics[metric.key] = {
        value,
        previousDelta:
          value === "" || previousValue === "" ? "" : roundMeasurementDelta(value - previousValue),
        firstDelta: value === "" || firstValue === "" ? "" : roundMeasurementDelta(value - firstValue),
      };
    });
    const thighRight = normalizeMeasurementValue(measurement.thighRight);
    const thighLeft = normalizeMeasurementValue(measurement.thighLeft);
    rowsById.set(measurement.id, {
      ...measurement,
      metrics,
      leftRightGap:
        thighRight === "" || thighLeft === "" ? "" : roundMeasurementDelta(Math.abs(thighRight - thighLeft)),
    });
  });
  return measurements.map((measurement) => rowsById.get(measurement.id)).filter(Boolean);
}

function isMeasurementPhotoAnswer(answer, response) {
  const questionId = normalizeText(answer?.questionId);
  const label = normalizeText(answer?.label);
  if (!Array.isArray(answer?.files) || !answer.files.length) return false;
  if (
    SESSION_MONITOR_PHOTOS_QUESTION_IDS.includes(questionId) ||
    SESSION_TICKET_END_PHOTOS_QUESTION_IDS.includes(questionId)
  ) {
    return true;
  }
  if (
    [
      LEGACY_SESSION_MONITOR_PHOTOS_QUESTION_ID,
      LEGACY_SESSION_TICKET_END_PHOTOS_QUESTION_ID,
      "q_bijiris_session_ticket_photos",
      "q_ticket_end_photo_last",
    ].includes(questionId)
  ) {
    return true;
  }
  if (response?.surveyId === TICKET_END_SURVEY_ID && questionId === "q_ticket_end_photo_last") {
    return true;
  }
  return /(モニター|計測写真|回数券終了時)/.test(label);
}

function getMeasurementPhotoGroupKind(answer) {
  const questionId = normalizeText(answer?.questionId);
  const label = normalizeText(answer?.label);
  if (
    SESSION_MONITOR_PHOTOS_QUESTION_IDS.includes(questionId) ||
    questionId === LEGACY_SESSION_MONITOR_PHOTOS_QUESTION_ID ||
    label.includes("モニター")
  ) {
    return "monitor";
  }
  return "measurement";
}

function normalizeMeasurementPhotoStageLabel(label) {
  const normalized = normalizeText(label);
  if (!normalized) return "計測写真";
  if (normalized.includes("モニター")) return "モニター時写真";
  if (normalized.includes("回数券終了")) return "回数券終了時";
  const measuredMatch = normalized.match(/計測写真[（(]?(.+?)[）)]?$/);
  if (measuredMatch?.[1]) {
    const detail = measuredMatch[1].replace(/^[（(]|[）)]$/g, "").trim();
    return detail ? `計測写真(${detail})` : normalized;
  }
  const stripped = normalized.replace(/写真\s*\d*枚?/g, "").replace(/\s+/g, " ").trim();
  return stripped || normalized;
}

function buildMeasurementPhotoGroups() {
  const responses = getVisibleHistoryResponses()
    .slice()
    .sort((a, b) => {
      const submittedDiff = new Date(a.submittedAt) - new Date(b.submittedAt);
      if (submittedDiff !== 0) return submittedDiff;
      return normalizeText(a.id).localeCompare(normalizeText(b.id), "ja");
    });
  const groups = [];
  let trialCount = 0;
  let singleCount = 0;

  responses.forEach((response, responseIndex) => {
    const answerMap = new Map((response.answers || []).map((answer) => [normalizeText(answer.questionId), answer]));
    const ticketMap = new Map(getResponseTicketInfo(response).map((item) => [item.label, item.value]));
    const sessionType = normalizeText(answerMap.get(SESSION_TYPE_QUESTION_ID)?.value);

    (response.answers || []).forEach((answer, answerIndex) => {
      if (!isMeasurementPhotoAnswer(answer, response)) return;
      const files = Array.isArray(answer.files) ? answer.files.filter(Boolean) : [];
      if (!files.length) return;

      const questionId = normalizeText(answer.questionId);
      const kind = getMeasurementPhotoGroupKind(answer);
      const ticketSheetLabel = normalizeText(ticketMap.get("何枚目") || "");
      let stageLabel = "";
      if (kind === "monitor") {
        stageLabel = "モニター時写真";
      } else if (ticketSheetLabel) {
        stageLabel = `回数券${ticketSheetLabel}終了時`;
      } else if (sessionType === "トライアル") {
        trialCount += 1;
        stageLabel = `トライアル${trialCount}回目`;
      } else if (sessionType === "単発") {
        singleCount += 1;
        stageLabel = `単発${singleCount}回目`;
      } else if (sessionType === "初回お試し") {
        stageLabel = "初回お試し";
      } else {
        stageLabel = normalizeMeasurementPhotoStageLabel(answer.label);
      }

      groups.push({
        id: `${normalizeText(response.id) || `response_${responseIndex}`}:${questionId || `photo_${answerIndex}`}`,
        response,
        questionId,
        label: normalizeText(answer.label) || "写真",
        files,
        kind,
        stageLabel,
        submittedAt: normalizeText(response.submittedAt),
        ticketInfo: getResponseTicketInfo(response),
      });
    });
  });

  return groups;
}

function getMeasurementPhotoCurrentLabel(group) {
  const count = Array.isArray(group?.files) ? group.files.length : 0;
  return count > 0 ? `計測写真${count}枚` : "計測写真";
}

function buildMeasurementPhotoTimelineEntries(photoGroups) {
  return photoGroups
    .slice()
    .sort((a, b) => {
      const submittedDiff = new Date(a.submittedAt) - new Date(b.submittedAt);
      if (submittedDiff !== 0) return submittedDiff;
      return normalizeText(a.id).localeCompare(normalizeText(b.id), "ja");
    })
    .map((group, index) => ({
      id: group.id,
      group,
      orderLabel: `${index + 1}`,
      submittedAt: group.submittedAt,
      title: group.stageLabel,
      photoLabel: getMeasurementPhotoCurrentLabel(group),
    }));
}

function renderMeasurementPhotoGroupCard(group, title) {
  if (!group) return "";
  return `
    <article class="history-card">
      <strong>${escapeHtml(title)}</strong>
      <div class="meta">${escapeHtml(formatDate(group.submittedAt))}</div>
      ${renderTicketStampList(group.ticketInfo || [])}
      ${renderAnswerValue({ files: group.files })}
    </article>
  `;
}

function renderMeasurementSummaryCards(measurements) {
  if (!measurements.length) {
    return `<div class="empty">まだ計測記録はありません。</div>`;
  }
  const latest = measurements[0];
  const first = measurements[measurements.length - 1];
  const thighGap =
    latest.thighRight !== "" && latest.thighLeft !== ""
      ? roundMeasurementDelta(Math.abs(latest.thighRight - latest.thighLeft))
      : "";
  return `
    <div class="measurement-summary-grid">
      <article class="measurement-summary-card">
        <div class="measurement-summary-label">最新測定日</div>
        <strong>${escapeHtml(formatDateOnly(latest.measuredAt))}</strong>
        <div class="meta">履歴 ${measurements.length}件</div>
      </article>
      <article class="measurement-summary-card">
        <div class="measurement-summary-label">最新 WHR</div>
        <strong>${escapeHtml(formatWhr(latest.whr))}</strong>
        <div class="meta">初回 ${escapeHtml(formatWhr(first.whr))}</div>
      </article>
      <article class="measurement-summary-card">
        <div class="measurement-summary-label">太もも左右差</div>
        <strong>${thighGap === "" ? "-" : `${escapeHtml(thighGap.toFixed(1))}cm`}</strong>
        <div class="meta">左右のバランス確認用</div>
      </article>
      ${MEASUREMENT_METRICS.map((metric) => {
        const latestValue = latest[metric.key];
        const firstValue = first[metric.key];
        const firstDelta =
          normalizeMeasurementValue(latestValue) === "" || normalizeMeasurementValue(firstValue) === ""
            ? ""
            : roundMeasurementDelta(normalizeMeasurementValue(latestValue) - normalizeMeasurementValue(firstValue));
        return `
          <article class="measurement-summary-card">
            <div class="measurement-summary-label">${escapeHtml(metric.label)}</div>
            <strong>${escapeHtml(formatMeasurementValue(latestValue))}</strong>
            <div class="meta">初回比 ${formatMeasurementDelta(firstDelta)}</div>
          </article>
        `;
      }).join("")}
    </div>
  `;
}

function renderMeasurementLineChart(title, measurements, metrics, options = {}) {
  const visibleMetrics = metrics.filter((metric) => appState.measurementMetricVisibility[metric.key] !== false);
  if (!measurements.length) {
    return `<div class="empty">まだグラフを表示できる計測記録がありません。</div>`;
  }
  if (!visibleMetrics.length) {
    return `<div class="empty">表示する項目を選択してください。</div>`;
  }

  const chronological = measurements
    .slice()
    .sort((a, b) => new Date(a.measuredAt) - new Date(b.measuredAt));
  const values = [];
  chronological.forEach((measurement) => {
    visibleMetrics.forEach((metric) => {
      const value = normalizeMeasurementValue(measurement[metric.key]);
      if (value !== "") values.push(Number(value));
    });
  });
  if (!values.length) {
    return `<div class="empty">表示できる数値がありません。</div>`;
  }

  const chartWidth = 720;
  const chartHeight = options.compact ? 220 : 280;
  const padding = { top: 20, right: 20, bottom: 44, left: 52 };
  const plotWidth = chartWidth - padding.left - padding.right;
  const plotHeight = chartHeight - padding.top - padding.bottom;
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);
  const pad = Math.max((maxValue - minValue) * 0.15, options.pad || 1);
  const domainMin = Math.max(0, Math.floor((minValue - pad) * 10) / 10);
  const domainMax = Math.ceil((maxValue + pad) * 10) / 10;
  const domainRange = domainMax - domainMin || 1;
  const xForIndex = (index) =>
    chronological.length === 1
      ? padding.left + plotWidth / 2
      : padding.left + (plotWidth * index) / (chronological.length - 1);
  const yForValue = (value) => padding.top + ((domainMax - value) / domainRange) * plotHeight;

  const grid = Array.from({ length: 5 }, (_, index) => {
    const value = domainMin + (domainRange * index) / 4;
    const y = yForValue(value);
    return `
      <g>
        <line x1="${padding.left}" y1="${y}" x2="${chartWidth - padding.right}" y2="${y}" class="measurement-grid-line" />
        <text x="${padding.left - 8}" y="${y + 4}" text-anchor="end" class="measurement-axis-label">
          ${escapeHtml(options.valueFormatter ? options.valueFormatter(value) : value.toFixed(1))}
        </text>
      </g>
    `;
  }).join("");

  const series = visibleMetrics
    .map((metric) => {
      const points = chronological
        .map((measurement, index) => {
          const value = normalizeMeasurementValue(measurement[metric.key]);
          if (value === "") return "";
          return `${xForIndex(index)},${yForValue(value)}`;
        })
        .filter(Boolean);
      if (!points.length) return "";
      const markers = chronological
        .map((measurement, index) => {
          const value = normalizeMeasurementValue(measurement[metric.key]);
          if (value === "") return "";
          return `
            <circle cx="${xForIndex(index)}" cy="${yForValue(value)}" r="4" fill="${metric.color}">
              <title>${escapeHtml(metric.label)} ${escapeHtml(formatDateOnly(measurement.measuredAt))}: ${escapeHtml(options.valueFormatter ? options.valueFormatter(value) : `${Number(value).toFixed(1)}${options.unit || ""}`)}</title>
            </circle>
          `;
        })
        .join("");
      return `
        <g>
          <polyline
            fill="none"
            stroke="${metric.color}"
            stroke-width="3"
            stroke-linecap="round"
            stroke-linejoin="round"
            points="${points.join(" ")}"
          />
          ${markers}
        </g>
      `;
    })
    .join("");

  const xLabels = chronological
    .map((measurement, index) => `
      <text x="${xForIndex(index)}" y="${chartHeight - 14}" text-anchor="middle" class="measurement-axis-label">
        ${escapeHtml(formatDateOnly(measurement.measuredAt).slice(5))}
      </text>
    `)
    .join("");

  return `
    <div class="measurement-legend">
      ${visibleMetrics
        .map((metric) => `
          <span class="measurement-legend-item">
            <span class="measurement-legend-dot" style="background:${metric.color}"></span>
            ${escapeHtml(metric.label)}
          </span>
        `)
        .join("")}
    </div>
    <div class="measurement-chart-shell">
      <svg class="measurement-chart" viewBox="0 0 ${chartWidth} ${chartHeight}" role="img" aria-label="${escapeHtml(title)}">
        ${grid}
        ${series}
        ${xLabels}
      </svg>
    </div>
  `;
}

function renderMeasurementHistoryList(measurements) {
  const rows = buildMeasurementHistoryRows(measurements);
  if (!rows.length) {
    return `<div class="empty">まだ計測履歴はありません。</div>`;
  }
  return `
    <div class="measurement-history-table-wrap">
      <table class="measurement-history-table">
        <thead>
          <tr>
            <th>測定日</th>
            <th>WHR</th>
            <th>左右差</th>
            ${MEASUREMENT_METRICS.map((metric) => `<th>${escapeHtml(metric.label)}</th>`).join("")}
          </tr>
        </thead>
        <tbody>
          ${rows
            .map((row) => `
              <tr>
                <td class="measurement-date-cell">
                  <strong>${escapeHtml(formatDateOnly(row.measuredAt))}</strong>
                </td>
                <td>
                  <div class="measurement-table-main">${escapeHtml(formatWhr(row.whr))}</div>
                </td>
                <td>
                  <div class="measurement-table-main">
                    ${row.leftRightGap === "" ? "-" : `${escapeHtml(row.leftRightGap.toFixed(1))}cm`}
                  </div>
                </td>
                ${MEASUREMENT_METRICS.map((metric) => {
                  const metricRow = row.metrics[metric.key];
                  return `
                    <td>
                      <div class="measurement-table-main">
                        ${escapeHtml(formatMeasurementValue(metricRow.value))}
                      </div>
                      <div class="measurement-table-sub">前回 ${formatMeasurementDelta(metricRow.previousDelta)}</div>
                      <div class="measurement-table-sub">初回 ${formatMeasurementDelta(metricRow.firstDelta)}</div>
                    </td>
                  `;
                }).join("")}
              </tr>
            `)
            .join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderMeasurementPhotoTimeline(entries) {
  if (!entries.length) {
    return `<div class="empty">まだ表示できる計測写真はありません。</div>`;
  }
  if (
    appState.selectedMeasurementPhotoComparisonId &&
    !entries.some((entry) => entry.id === appState.selectedMeasurementPhotoComparisonId)
  ) {
    appState.selectedMeasurementPhotoComparisonId = "";
  }
  const selectedEntry =
    entries.find((entry) => entry.id === appState.selectedMeasurementPhotoComparisonId) || null;

  return `
    <div class="measurement-photo-entry-list">
      ${entries
        .map(
          (entry) => `
            <button
              class="measurement-photo-entry ${entry.id === selectedEntry?.id ? "is-active" : ""}"
              type="button"
              data-measurement-photo-entry="${escapeHtml(entry.id)}"
            >
              <div class="measurement-photo-entry-order">#${escapeHtml(entry.orderLabel)}</div>
              <strong>${escapeHtml(entry.title)}</strong>
              <div class="meta">${escapeHtml(entry.photoLabel)}</div>
              <div class="meta">${escapeHtml(formatDate(entry.submittedAt))}</div>
            </button>
          `,
        )
        .join("")}
    </div>
    ${
      selectedEntry
        ? `<div class="measurement-photo-detail">
            ${renderMeasurementPhotoGroupCard(
              selectedEntry.group,
              `${selectedEntry.group.stageLabel}の${getMeasurementPhotoCurrentLabel(selectedEntry.group)}`,
            )}
          </div>`
        : `<div class="empty">上の一覧から見たい計測写真を選択してください。</div>`
    }
  `;
}

function renderMeasurements() {
  if (!measurementPanel) return;
  if (!hasCustomerSession()) {
    measurementPanel.innerHTML = `<div class="empty">先にログインしてください。</div>`;
    return;
  }
  if (appState.historyLoading && !appState.measurements.length && !appState.history.length) {
    measurementPanel.innerHTML = `<div class="empty">読み込み中です。</div>`;
    return;
  }
  if (appState.historyLoadError && !appState.measurements.length && !appState.history.length) {
    measurementPanel.innerHTML = `<div class="empty">${escapeHtml(appState.historyLoadError)}</div>`;
    return;
  }

  const allMeasurements = getCustomerMeasurements();
  const filteredMeasurements = filterMeasurementsByPeriod(allMeasurements, appState.selectedMeasurementPeriod);
  const photoTimelineEntries = buildMeasurementPhotoTimelineEntries(buildMeasurementPhotoGroups());

  measurementPanel.innerHTML = `
    <div class="measurement-section">
      ${renderMeasurementSummaryCards(allMeasurements)}
      <article class="history-card">
        <div class="section-head">
          <div>
            <strong>推移グラフ</strong>
            <div class="meta">期間を切り替えて、部位ごとの変化を確認できます。</div>
          </div>
          <label>
            表示期間
            <select id="measurementPeriodFilter">
              ${MEASUREMENT_PERIOD_OPTIONS.map((option) => `
                <option value="${escapeHtml(option.value)}" ${appState.selectedMeasurementPeriod === option.value ? "selected" : ""}>
                  ${escapeHtml(option.label)}
                </option>
              `).join("")}
            </select>
          </label>
        </div>
        <div class="measurement-toggle-row">
          ${MEASUREMENT_METRICS.map((metric) => `
            <label class="measurement-toggle-chip">
              <input
                type="checkbox"
                data-toggle-measurement-metric="${escapeHtml(metric.key)}"
                ${appState.measurementMetricVisibility[metric.key] !== false ? "checked" : ""}
              />
              ${escapeHtml(metric.label)}
            </label>
          `).join("")}
          <label class="measurement-toggle-chip">
            <input
              type="checkbox"
              data-toggle-measurement-metric="whr"
              ${appState.measurementMetricVisibility.whr !== false ? "checked" : ""}
            />
            WHR
          </label>
        </div>
        ${renderMeasurementLineChart("部位別の推移", filteredMeasurements, MEASUREMENT_METRICS, {
          unit: "cm",
        })}
        ${
          appState.measurementMetricVisibility.whr !== false
            ? renderMeasurementLineChart(
                "WHR の推移",
                filteredMeasurements,
                [{ key: "whr", label: "WHR", color: "#7a5b9f", unit: "" }],
                {
                  compact: true,
                  pad: 0.03,
                  valueFormatter: (value) => Number(value).toFixed(3),
                },
              )
            : ""
        }
      </article>

      <article class="history-card">
        <strong>計測履歴</strong>
        <div class="meta">新しい記録が上に表示されます。表一覧で前回比と初回比を確認できます。</div>
        ${renderMeasurementHistoryList(filteredMeasurements)}
      </article>

      <article class="history-card">
        <strong>計測写真一覧</strong>
        <div class="meta">アンケートでアップロードした計測写真を、モニター時・回数券終了時・トライアル回ごとの時系列一覧から選んで確認できます。</div>
        ${renderMeasurementPhotoTimeline(photoTimelineEntries)}
      </article>
    </div>
  `;

  measurementPanel.querySelector("#measurementPeriodFilter")?.addEventListener("change", (event) => {
    appState.selectedMeasurementPeriod = event.currentTarget.value || "6m";
    renderMeasurements();
  });
  measurementPanel.querySelectorAll("[data-toggle-measurement-metric]").forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      appState.measurementMetricVisibility = {
        ...appState.measurementMetricVisibility,
        [checkbox.dataset.toggleMeasurementMetric]: checkbox.checked,
      };
      renderMeasurements();
    });
  });
  measurementPanel.querySelectorAll("[data-measurement-photo-entry]").forEach((button) => {
    button.addEventListener("click", () => {
      appState.selectedMeasurementPhotoComparisonId = button.dataset.measurementPhotoEntry || "";
      renderMeasurements();
    });
  });
}

async function retryPendingSubmission() {
  if (!appState.pendingSubmission) return;
  try {
    const result = await api.request(appState.pendingSubmission.path, {
      method: appState.pendingSubmission.method,
      body: appState.pendingSubmission.body,
    });
    clearPendingSubmission();
    appState.lastSubmittedResponse = result.response || null;
    if (appState.pendingSubmission.surveyId) {
      clearSurveyDraft(appState.pendingSubmission.surveyId);
      appState.selectedSurveyId = appState.pendingSubmission.surveyId;
      appState.panelMode = "complete";
      appState.lastSubmissionWasEdit = appState.pendingSubmission.method === "PUT";
    }
    await loadHistory();
    renderSurveys();
    renderAnswerPanel();
    showToast("未送信の回答を再送信しました。");
  } catch (error) {
    reportClientError("customer.retryPending", error, {
      surveyId: appState.pendingSubmission?.surveyId,
    });
    showToast(error.message || "再送信できませんでした。");
  }
}

function attachCommonButtons() {
  document.querySelectorAll("[data-retry-pending]").forEach((button) => {
    button.addEventListener("click", () => {
      void retryPendingSubmission();
    });
  });
}

async function runAppUpdate() {
  if (appUpdateButton) {
    appUpdateButton.disabled = true;
    appUpdateButton.textContent = "更新中";
  }
  try {
    if ("serviceWorker" in navigator) {
      const registration = await navigator.serviceWorker.getRegistration();
      await registration?.update().catch(() => {});
      registration?.waiting?.postMessage({ type: "SKIP_WAITING" });
    }
    if ("caches" in window) {
      const cacheKeys = await caches.keys();
      await Promise.all(cacheKeys.filter((key) => key.startsWith(CACHE_PREFIX)).map((key) => caches.delete(key)));
    }
    showToast("アプリを更新しています。");
    const url = new URL(window.location.href);
    url.searchParams.set("appRefresh", String(Date.now()));
    window.location.replace(url.toString());
  } catch (error) {
    reportClientError("customer.appUpdate", error);
    showToast(error.message || "アプリを更新できませんでした。");
  } finally {
    if (appUpdateButton) {
      appUpdateButton.disabled = false;
      appUpdateButton.textContent = "更新";
    }
  }
}

async function pruneAppCaches() {
  if (!("caches" in window)) return;
  const cacheKeys = await caches.keys();
  const staleKeys = cacheKeys.filter(
    (key) => key.startsWith(CACHE_PREFIX) && key !== ACTIVE_CACHE_NAME,
  );
  await Promise.all(staleKeys.map((key) => caches.delete(key)));
}

async function runAutoCacheMaintenance({ force = false } = {}) {
  try {
    const lastRun = Number(localStorage.getItem(AUTO_CACHE_MAINTENANCE_KEY) || "0");
    const now = Date.now();
    if (!force && now - lastRun < AUTO_CACHE_MAINTENANCE_INTERVAL_MS) return;
    localStorage.setItem(AUTO_CACHE_MAINTENANCE_KEY, String(now));

    if ("serviceWorker" in navigator) {
      const registration = await navigator.serviceWorker.getRegistration();
      await registration?.update().catch(() => {});
      registration?.waiting?.postMessage({ type: "SKIP_WAITING" });
    }
    await pruneAppCaches();
  } catch (error) {
    reportClientError("customer.autoCacheMaintenance", error);
  }
}

function scheduleAutoCacheMaintenance(force = false) {
  const task = () => {
    void runAutoCacheMaintenance({ force });
  };
  if ("requestIdleCallback" in window) {
    window.requestIdleCallback(task, { timeout: 10000 });
    return;
  }
  window.setTimeout(task, 1200);
}

function setupAutoCacheMaintenance() {
  window.addEventListener("load", () => {
    scheduleAutoCacheMaintenance();
  });
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      scheduleAutoCacheMaintenance();
    }
  });
  window.setInterval(() => {
    scheduleAutoCacheMaintenance(true);
  }, AUTO_CACHE_MAINTENANCE_INTERVAL_MS);
}

function setupInstall() {
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker
        .register("./sw.js?v=20260416-09", { updateViaCache: "none" })
        .then((registration) => {
          const activateWaiting = () => {
            registration.waiting?.postMessage({ type: "SKIP_WAITING" });
          };
          activateWaiting();
          registration.addEventListener("updatefound", () => {
            registration.installing?.addEventListener("statechange", () => {
              if (registration.waiting) activateWaiting();
            });
          });
          return registration.update().catch(() => {});
        })
        .catch(() => {});
    });
  }

  window.addEventListener("beforeinstallprompt", (event) => {
    event.preventDefault();
    appState.installPrompt = event;
    installButton.hidden = false;
  });

  installButton.addEventListener("click", async () => {
    if (!appState.installPrompt) return;
    appState.installPrompt.prompt();
    await appState.installPrompt.userChoice;
    appState.installPrompt = null;
    installButton.hidden = true;
  });
}

customerLoginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!canRegisterFromThisContext()) {
    showToast("初回会員登録はホーム画面に追加したアプリから行ってください。");
    return;
  }
  try {
    await applyCustomerSession(readCustomerProfileFromForm(event.currentTarget));
  } catch (error) {
    reportClientError("customer.login", error);
    showToast(error.message || "ログインできませんでした。");
  }
});

recoverAccountButton?.addEventListener("click", async () => {
  try {
    await applyCustomerSession(readCustomerProfileFromForm(customerLoginForm), {
      recovery: true,
    });
  } catch (error) {
    reportClientError("customer.recover", error);
    showToast(error.message || "アカウントを復旧できませんでした。");
  }
});

customerForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  try {
    const profile = readCustomerProfileFromForm(event.currentTarget);
    const previousName = appState.customer.name;
    appState.customer = {
      ...profile,
      historyMatchMode: appState.customer.historyMatchMode,
    };
    appState.pushServerSignature = "";
    saveLocal(CUSTOMER_KEY, appState.customer);
    syncCustomerForms();
    showToast("会員情報を保存しました。");
    if (previousName !== appState.customer.name) {
      appState.historySurveyId = "";
      appState.historyResponseId = "";
    }
    await loadHistory();
    await initializePushNotifications();
    renderHistory();
  } catch (error) {
    reportClientError("customer.profile.update", error);
    showToast(error.message || "お客様情報を保存できませんでした。");
  }
});

document.querySelectorAll("[data-page]").forEach((button) => {
  button.addEventListener("click", () => {
    if (!hasCustomerSession()) {
      setPage("login");
      return;
    }
    setPage(button.dataset.page);
  });
});

document.querySelector("#refreshButton").addEventListener("click", () => {
  void loadSurveys();
  void loadBijirisPosts();
  if (hasCustomerSession()) void loadHistory();
});

document.querySelector("#historyRefreshButton").addEventListener("click", () => {
  void loadHistory();
});
measurementRefreshButton?.addEventListener("click", () => {
  void loadHistory();
});
bijirisRefreshButton?.addEventListener("click", () => {
  void loadBijirisPosts();
});
pushToggleButton?.addEventListener("click", () => {
  void togglePushNotifications();
});
appUpdateButton?.addEventListener("click", () => {
  void runAppUpdate();
});

setupInstall();
setupAutoCacheMaintenance();
window.addEventListener("beforeunload", (event) => {
  const hasAnyDraft = Object.keys(appState.drafts || {}).some((surveyId) =>
    hasDraftContent(getSurveyDraft(surveyId)),
  );
  if (!appState.pendingSubmission && !hasAnyDraft) return;
  event.preventDefault();
  event.returnValue = "";
});
window.addEventListener("error", (event) => {
  reportClientError("customer.window", event.error || event.message || "window error");
});
window.addEventListener("unhandledrejection", (event) => {
  reportClientError("customer.promise", event.reason || "unhandled rejection");
});
syncCustomerForms();
renderRegistrationGuide();
renderHomeTicketStatus();
updatePushUi();
renderSurveys();
renderAnswerPanel();
renderMeasurements();
renderBijirisPosts();
void initializePushNotifications();
void loadSurveys();
void loadBijirisPosts();
setPage(hasCustomerSession() ? "home" : "login");
