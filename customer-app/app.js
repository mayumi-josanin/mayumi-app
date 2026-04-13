const CUSTOMER_KEY = "mayumi_survey_customer";
const DRAFTS_KEY = "mayumi_survey_drafts";
const PENDING_KEY = "mayumi_survey_pending_submission";
const TICKET_CARD_OVERRIDE_KEY = "mayumi_survey_ticket_card_overrides";
const PHOTO_FILE_LIMIT = 6;
const PHOTO_MAX_SIZE = 1400;
const PHOTO_JPEG_QUALITY = 0.74;
const RESPONSE_EDIT_WINDOW_MS = 24 * 60 * 60 * 1000;
const APP_VERSION = "20260413-21";
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
const SESSION_LIFE_CHANGES_QUESTION_ID = "q_bijiris_session_life_changes";
const SESSION_LIFE_CHANGES_OTHER_QUESTION_ID = "q_bijiris_session_life_changes_other";
const TICKET_END_SURVEY_ID = "survey_bijiris_ticket_end";
const LEGACY_TICKET_INFO_QUESTION_IDS = {
  size: "q_ticket_end_ticket_size",
  sheet: "q_ticket_end_ticket_sheet",
  round: "q_ticket_end_ticket_round",
};
const TICKET_END_COUNT_QUESTION_ID = LEGACY_TICKET_INFO_QUESTION_IDS.size;
const TICKET_END_SHEET_QUESTION_ID = LEGACY_TICKET_INFO_QUESTION_IDS.sheet;
const TICKET_END_ROUND_QUESTION_ID = LEGACY_TICKET_INFO_QUESTION_IDS.round;
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
  surveys: [],
  history: [],
  publicInfo: {
    dataPolicyText: "",
    requireConsent: true,
    consentText: "",
    version: "",
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
  concernCategoryByQuestion: {},
};

const api = window.MayumiSurveyApi;
const toast = document.querySelector("#toast");
const surveyList = document.querySelector("#surveyList");
const answerPanel = document.querySelector("#answerPanel");
const historyList = document.querySelector("#historyList");
const homeTicketStatus = document.querySelector("#homeTicketStatus");
const customerLoginForm = document.querySelector("#customerLoginForm");
const customerForm = document.querySelector("#customerForm");
const appUpdateButton = document.querySelector("#appUpdateButton");
const installButton = document.querySelector("#installButton");
const registrationLead = document.querySelector("#registrationLead");
const customerRegisterButton = document.querySelector("#customerRegisterButton");
const recoverAccountButton = document.querySelector("#recoverAccountButton");
const bottomNav = document.querySelector("#bottomNav");

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

function normalizeText(value) {
  return String(value ?? "").trim();
}

function normalizeKana(value) {
  return String(value ?? "").replace(/\s+/g, "").trim();
}

function normalizeCustomerProfile(value) {
  return {
    name: normalizeText(value?.name),
    nameKana: normalizeKana(value?.nameKana),
    historyMatchMode: value?.historyMatchMode === "name" ? "name" : "device",
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

function normalizeServerCustomerProfile(value) {
  return {
    name: normalizeText(value?.name),
    nameKana: normalizeKana(value?.nameKana),
  };
}

function mergeCustomerProfile(baseProfile, serverProfile) {
  const base = normalizeCustomerProfile(baseProfile);
  const server = normalizeServerCustomerProfile(serverProfile);
  if (!server.name && !server.nameKana) return base;
  return normalizeCustomerProfile({
    name: server.name || base.name,
    nameKana: server.nameKana || base.nameKana,
    historyMatchMode: base.historyMatchMode,
  });
}

function syncCustomerProfileFromServer(serverProfile) {
  const merged = mergeCustomerProfile(appState.customer, serverProfile);
  if (
    merged.name === appState.customer.name &&
    merged.nameKana === appState.customer.nameKana &&
    merged.historyMatchMode === appState.customer.historyMatchMode
  ) {
    return false;
  }
  appState.customer = merged;
  saveLocal(CUSTOMER_KEY, appState.customer);
  syncCustomerForms();
  return true;
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
  return {
    values: typeof draft?.values === "object" && draft.values ? draft.values : {},
    photos: typeof draft?.photos === "object" && draft.photos ? draft.photos : {},
  };
}

function getSurveyDraft(surveyId) {
  return normalizeDraft(appState.drafts[surveyId]);
}

function setSurveyDraft(surveyId, draft) {
  appState.drafts[surveyId] = normalizeDraft(draft);
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
  if ((page === "history" || page === "home") && hasCustomerSession()) {
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
    };
    renderSurveys();
    renderAnswerPanel();
  } catch (error) {
    reportClientError("customer.loadSurveys", error);
    const fallbackSurveys = getFallbackSurveys();
    if (fallbackSurveys.length) {
      appState.surveys = fallbackSurveys;
      renderSurveys();
      renderAnswerPanel();
      showToast("アンケート取得に失敗したため、保存済みの一覧を表示しています。");
      return;
    }
    surveyList.innerHTML = `<div class="empty">${escapeHtml(error.message || "アンケートを読み込めませんでした。")}</div>`;
  }
}

async function loadHistory() {
  if (!hasCustomerSession()) {
    appState.history = [];
    appState.historyLoading = false;
    appState.historyLoadError = "";
    renderHomeTicketStatus();
    historyList.innerHTML = `<div class="empty">先にログインしてください。</div>`;
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
    syncCustomerProfileFromServer(result.customerProfile);
    syncActiveTicketCardOverrideWithHistory();
    appState.historyLoading = false;
    appState.historyLoadError = "";
    renderHomeTicketStatus();
    renderHistory();
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
    appState.historyLoading = false;
    appState.historyLoadError = error.message || "履歴を読み込めませんでした。";
    renderHomeTicketStatus();
    reportClientError("customer.loadHistory", error, { name: appState.customer.name });
    historyList.innerHTML = `<div class="empty">${escapeHtml(error.message || "履歴を読み込めませんでした。")}</div>`;
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
}

function readCustomerProfileFromForm(form) {
  const formData = new FormData(form);
  return normalizeCustomerProfile({
    name: formData.get("name"),
    nameKana: formData.get("nameKana"),
    historyMatchMode: appState.customer.historyMatchMode,
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
  saveLocal(CUSTOMER_KEY, appState.customer);
  syncCustomerForms();
  await loadHistory();
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
  if (!appState.surveys.length) {
    surveyList.innerHTML = `${renderPendingNotice()}<div class="empty">現在回答できるアンケートはありません。</div>`;
    attachCommonButtons();
    return;
  }

  surveyList.innerHTML = `
    ${renderPendingNotice()}
    ${appState.surveys
      .map((survey) => {
        const availability = getSurveyAvailability(survey);
        return `
          <button
            class="survey-card ${survey.id === appState.selectedSurveyId ? "active" : ""} ${availability.open ? "" : "disabled"} ${availability.nearDeadline ? "deadline-soon" : ""}"
            type="button"
            data-survey-id="${survey.id}"
          >
            <strong>${escapeHtml(survey.title)}</strong>
            <span>${escapeHtml(survey.description)}</span>
            <div class="survey-card-meta">
              <span class="badge ${availability.open ? "open" : "closed"}">${escapeHtml(availability.label)}</span>
              <span class="meta">${escapeHtml(availability.detail)}</span>
            </div>
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

function refreshProgressDisplay(survey) {
  const progress = getProgress(survey, getSurveyDraft(survey.id));
  document.querySelector(".progress-fill")?.setAttribute("style", `width: ${progress.percent}%`);
  const progressHead = document.querySelector(".progress-head span");
  if (progressHead) {
    progressHead.textContent = `${progress.answeredRequired}/${progress.requiredTotal} 必須回答`;
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
  document.querySelector("#sessionTicketRoundForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const ticketRound = normalizeText(formData.get("sessionTicketRound"));
    if (!ticketRound) {
      showToast("回数券の何回目ですか？を選択してください。");
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
  const required = isQuestionRequired(question, surveyId);
  const selectedCheckboxValues = Array.isArray(draft.values[question.id]) ? draft.values[question.id] : [];
  const label = `
    <span>${index + 1}. ${escapeHtml(getQuestionLabel(question, surveyId))}</span>
    ${required ? "" : `<span class="meta">任意</span>`}
  `;

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
      <fieldset class="question-block" data-question-wrap="${question.id}">
        <legend>${label}</legend>
        <div class="rating-row">
          ${[1, 2, 3, 4, 5]
            .map(
              (value) => `
                <label>
                  <input type="radio" name="${name}" value="${value}" data-question-id="${question.id}" ${String(draft.values[question.id] || "") === String(value) ? "checked" : ""} />
                  ${value}
                </label>
              `,
            )
            .join("")}
        </div>
      </fieldset>
    `;
  }

  if (question.type === "choice") {
    return `
      <fieldset class="question-block" data-question-wrap="${question.id}">
        <legend>${label}</legend>
        <div class="choice-row">
          ${question.options
            .map(
              (option) => `
                <label>
                  <input type="radio" name="${name}" value="${escapeHtml(option)}" data-question-id="${question.id}" ${String(draft.values[question.id] || "") === option ? "checked" : ""} />
                  ${escapeHtml(option)}
                </label>
              `,
            )
            .join("")}
        </div>
      </fieldset>
    `;
  }

  if (question.type === "checkbox") {
    const selected = new Set(selectedCheckboxValues);
    if (question.id === SESSION_CONCERN_QUESTION_ID) {
      const activeCategoryId = getConcernActiveCategory(surveyId, question.id, selectedCheckboxValues);
      return `
        <fieldset class="question-block" data-question-wrap="${question.id}">
          <legend>${label}</legend>
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
                                  ${escapeHtml(option)}
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
          ${activeCategoryId ? "" : `<div class="empty">カテゴリを選択すると詳細項目を表示します。</div>`}
        </fieldset>
      `;
    }
    return `
      <fieldset class="question-block" data-question-wrap="${question.id}">
        <legend>${label}</legend>
        <div class="checkbox-row">
          ${question.options
            .map(
              (option) => `
                <label>
                  <input type="checkbox" name="${name}" value="${escapeHtml(option)}" data-question-id="${question.id}" ${selected.has(option) ? "checked" : ""} />
                  ${escapeHtml(option)}
                </label>
              `,
            )
            .join("")}
        </div>
      </fieldset>
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
  const response = appState.lastSubmittedResponse;
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
        response?.id
          ? `<div class="receipt-box"><strong>受付番号</strong><span>${escapeHtml(response.id)}</span></div>`
          : ""
      }
      ${
        response?.submittedAt
          ? `<div class="meta">送信日時: ${formatDate(response.submittedAt)}</div>`
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
      ${visibleQuestions.map((question, index) => renderQuestion(question, index, surveyId)).join("")}
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
        visibleQuestions.length
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

function acquireNextTicketSheet(ticketPlan) {
  const nextSheetLabel = getNextTicketSheetLabelForPlan(ticketPlan);
  const nextSheetNumber = parseTicketSheet(nextSheetLabel);
  if (!normalizeText(ticketPlan) || !nextSheetNumber) {
    showToast("新しいスタンプカードを取得できませんでした。");
    return;
  }
  setActiveTicketCardOverride(ticketPlan, nextSheetNumber);
  renderHomeTicketStatus();
  showToast(`${ticketPlan} ${nextSheetLabel} を取得しました。`);
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
    file?.downloadUrl || file?.previewUrl || file?.thumbnailUrl || file?.dataUrl || file?.url || "#",
  );
}

function renderTicketStampProgress(ticketCount, currentRound) {
  if (!ticketCount) return "";
  const normalizedRound = Math.max(0, Number(currentRound) || 0);
  return `
    <div class="ticket-progress">
      ${Array.from({ length: ticketCount }, (_, index) => {
        const step = index + 1;
        return `
          <span class="stamp-dot ${step <= normalizedRound ? "active" : ""}" aria-label="${step}回目">
            ${step}
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
      acquireNextTicketSheet(button.dataset.acquireTicketPlan || "");
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
      await Promise.all(
        cacheKeys
          .filter((key) => key.startsWith("mayumi-customer-survey-"))
          .map((key) => caches.delete(key)),
      );
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

function setupInstall() {
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker
        .register("./sw.js?v=20260413-21", { updateViaCache: "none" })
        .then((registration) => registration.update().catch(() => {}))
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
    saveLocal(CUSTOMER_KEY, appState.customer);
    syncCustomerForms();
    showToast("会員情報を保存しました。");
    if (previousName !== appState.customer.name) {
      appState.historySurveyId = "";
      appState.historyResponseId = "";
    }
    await loadHistory();
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
  if (hasCustomerSession()) void loadHistory();
});

document.querySelector("#historyRefreshButton").addEventListener("click", () => {
  void loadHistory();
});
appUpdateButton?.addEventListener("click", () => {
  void runAppUpdate();
});

setupInstall();
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
renderSurveys();
renderAnswerPanel();
void loadSurveys();
setPage(hasCustomerSession() ? "home" : "login");
