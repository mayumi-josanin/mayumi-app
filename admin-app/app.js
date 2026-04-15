const TOKEN_KEY = "mayumi_survey_admin_token";
const STATUS_LABELS = {
  new: "未対応",
  checked: "確認済み",
  done: "対応済み",
  trash: "ゴミ箱",
};
const TICKET_INFO_QUESTION_IDS = {
  size: ["q_bijiris_session_ticket_plan", "q_ticket_end_ticket_size"],
  sheet: ["q_bijiris_session_ticket_sheet", "q_ticket_end_ticket_sheet"],
  round: ["q_bijiris_session_ticket_round", "q_ticket_end_ticket_round"],
};
const SESSION_CONCERN_QUESTION_ID = "q_bijiris_session_concern";
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
const CUSTOMER_TICKET_PLAN_OPTIONS = ["", "6回券", "10回券"];
const CUSTOMER_TICKET_SHEET_OPTIONS = Array.from({ length: 20 }, (_, index) => `${index + 1}枚目`);
const CUSTOMER_TICKET_ROUND_OPTIONS = Array.from({ length: 10 }, (_, index) => `${index + 1}回目`);
const MEASUREMENT_METRICS = [
  { key: "waist", label: "ウエスト", shortLabel: "W", color: "#c95f50", unit: "cm" },
  { key: "hip", label: "ヒップ", shortLabel: "H", color: "#4e8c73", unit: "cm" },
  { key: "thighRight", label: "太もも右", shortLabel: "右", color: "#c78a2c", unit: "cm" },
  { key: "thighLeft", label: "太もも左", shortLabel: "左", color: "#6e7fba", unit: "cm" },
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
const BIJIRIS_POST_CATEGORY_OPTIONS = [
  "豆知識",
  "アドバイス",
  "セルフケア",
  "お知らせ",
  "よくある質問",
];
const BIJIRIS_POST_TEMPLATES = [
  {
    id: "selfcare_breathing",
    label: "セルフケア: 呼吸",
    category: "セルフケア",
    title: "骨盤底筋を意識しやすくする呼吸のコツ",
    summary: "毎日の呼吸で骨盤底筋を意識しやすくするポイントを短くまとめます。",
    body: [
      "深く息を吐くときに、お腹の奥がやさしく締まる感覚を意識してみてください。",
      "肩を上げずに、みぞおちから下がゆるやかに動く呼吸にすると続けやすくなります。",
      "無理に力を入れすぎず、息を吐き切ることを優先すると感覚をつかみやすくなります。",
    ].join("\n\n"),
  },
  {
    id: "advice_posture",
    label: "アドバイス: 姿勢",
    category: "アドバイス",
    title: "日常で姿勢を整えるときの見直しポイント",
    summary: "立つ・座る・歩く場面で意識しやすい姿勢のポイントをまとめます。",
    body: [
      "座る時は腰だけで支えず、坐骨で座る意識を持つと下腹の力が入りやすくなります。",
      "立つ時は胸を張りすぎず、あごを軽く引いて重心を真ん中に戻してみてください。",
      "長時間同じ姿勢が続く時は、1時間に1回ほど骨盤を前後にゆるく動かすのがおすすめです。",
    ].join("\n\n"),
  },
  {
    id: "faq_ticket",
    label: "よくある質問: 回数券",
    category: "よくある質問",
    title: "回数券で通う時によくあるご質問",
    summary: "続け方や変化の見方について、よくいただく質問をまとめます。",
    body: [
      "体感や変化の出方には個人差がありますが、記録を見返すと小さな変化に気づきやすくなります。",
      "写真やアンケートを残しておくと、初回との違いを比較しやすくなります。",
      "気になることがある時は、その都度アンケートや施術時にご相談ください。",
    ].join("\n\n"),
  },
  {
    id: "notice_campaign",
    label: "お知らせ",
    category: "お知らせ",
    title: "今月のお知らせ",
    summary: "営業日やご案内内容を共有するためのテンプレートです。",
    body: [
      "今月のご案内内容をここに入力してください。",
      "営業時間やご予約に関する変更がある場合もこの欄に追記できます。",
      "必要に応じて写真やPDF資料も添付してください。",
    ].join("\n\n"),
  },
];

const api = window.MayumiSurveyApi;
const state = {
  token: localStorage.getItem(TOKEN_KEY) || "",
  surveys: [],
  responses: [],
  measurements: [],
  bijirisPosts: [],
  adminInfo: null,
  customerProfiles: {},
  preferences: null,
  logs: { auditLogs: [], errorLogs: [] },
  customerMemos: {},
  adminUsers: [],
  selectedSurveyEditorId: "",
  selectedAnalyticsSurveyId: "",
  selectedResponseSurveyId: "",
  selectedCustomerName: "",
  selectedResponseId: "",
  selectedResponseIds: [],
  selectedCustomerViewName: "",
  selectedCustomerViewSurveyId: "",
  selectedCustomerViewResponseId: "",
  selectedCustomerTimelineOpen: false,
  selectedMeasurementEditId: "",
  selectedBijirisPostId: "",
  selectedBijirisView: "list",
  bijirisEditorDraft: null,
  selectedMeasurementPeriod: "6m",
  measurementMetricVisibility: { ...DEFAULT_MEASUREMENT_VISIBILITY },
  installPrompt: null,
};

const loginView = document.querySelector("#loginView");
const adminView = document.querySelector("#adminView");
const toast = document.querySelector("#toast");
const loginForm = document.querySelector("#loginForm");
const credentialForm = document.querySelector("#credentialForm");
const appUpdateButton = document.querySelector("#appUpdateButton");
const installButton = document.querySelector("#installButton");
const loginSubmitButton = document.querySelector("#loginSubmitButton");

function getVisibleSurveyIdSet() {
  return new Set(
    state.surveys
      .map((survey) => String(survey?.id || "").trim())
      .filter(Boolean),
  );
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
  if (!value) return "-";
  return new Intl.DateTimeFormat("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(new Date(value));
}

function parseTicketLabelNumber(value) {
  const matched = String(value || "").match(/\d+/);
  return matched ? Number(matched[0]) : 0;
}

function normalizeActiveTicketCard(value) {
  if (!value || typeof value !== "object") return null;
  const plan = String(value.plan || "").trim();
  const sheetNumber = Math.floor(Number(value.sheetNumber) || parseTicketLabelNumber(value.sheetLabel));
  const round = Math.max(0, Math.floor(Number(value.round) || parseTicketLabelNumber(value.roundLabel)));
  if (!plan || sheetNumber <= 0) return null;
  return {
    plan,
    sheetNumber,
    round,
  };
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

function normalizeCustomerProfile(value) {
  return {
    name: String(value?.name || "").trim(),
    memberNumber: String(value?.memberNumber || "").trim().toUpperCase(),
    nameKana: String(value?.nameKana || "").trim(),
    activeTicketCard: normalizeActiveTicketCard(value?.activeTicketCard),
    measurementTargets: normalizeMeasurementTargets(value?.measurementTargets),
    pushStatus: normalizePushStatus(value?.pushStatus),
    updatedAt: String(value?.updatedAt || "").trim(),
  };
}

function normalizePushStatus(value) {
  if (!value || typeof value !== "object") return null;
  const hasEnabled = Object.prototype.hasOwnProperty.call(value, "enabled");
  const hasSupported = Object.prototype.hasOwnProperty.call(value, "supported");
  const permission = String(value?.permission || "").trim().toLowerCase();
  const normalizedPermission = ["granted", "denied", "default", "unsupported"].includes(permission) ? permission : "";
  if (!hasEnabled && !hasSupported && !normalizedPermission) return null;
  return {
    enabled: value.enabled === true,
    supported: value.supported === true,
    permission: normalizedPermission || (value.supported === false ? "unsupported" : ""),
    updatedAt: String(value?.updatedAt || "").trim(),
  };
}

function describePushStatus(pushStatus) {
  const status = normalizePushStatus(pushStatus);
  if (!status) {
    return { label: "未設定", detail: "顧客側でまだ通知設定が同期されていません。", badgeClass: "draft" };
  }
  if (status.supported !== true) {
    return { label: "未対応", detail: "この端末または表示方法では通知を利用できません。", badgeClass: "draft" };
  }
  if (status.enabled) {
    return { label: "オン", detail: "顧客アプリで通知受信が有効です。", badgeClass: "open" };
  }
  if (status.permission === "denied") {
    return { label: "拒否", detail: "端末側で通知が拒否されています。", badgeClass: "closed" };
  }
  return { label: "オフ", detail: "顧客アプリで通知はオフです。", badgeClass: "draft" };
}

function renderCustomerPushStatus(pushStatus) {
  const described = describePushStatus(pushStatus);
  const status = normalizePushStatus(pushStatus);
  return `
    <div class="meta">
      通知状態:
      <span class="badge ${described.badgeClass}">${escapeHtml(described.label)}</span>
      ${escapeHtml(described.detail)}
      ${status?.updatedAt ? ` / 最終更新: ${escapeHtml(formatDate(status.updatedAt))}` : ""}
    </div>
  `;
}

function normalizeMeasurementRecord(value) {
  const measuredAt = String(value?.measuredAt || "").trim();
  const createdAt = String(value?.createdAt || "").trim();
  const updatedAt = String(value?.updatedAt || "").trim();
  const waist = normalizeMeasurementValue(value?.waist);
  const hip = normalizeMeasurementValue(value?.hip);
  const thighRight = normalizeMeasurementValue(value?.thighRight);
  const thighLeft = normalizeMeasurementValue(value?.thighLeft);
  return {
    id: String(value?.id || "").trim(),
    customerName: String(value?.customerName || "").trim(),
    memberNumber: String(value?.memberNumber || "").trim().toUpperCase(),
    measuredAt,
    waist,
    hip,
    thighRight,
    thighLeft,
    whr: hip === "" || !(Number(hip) > 0) || waist === "" ? "" : Math.round((waist / hip) * 1000) / 1000,
    memo: String(value?.memo || "").trim(),
    createdAt,
    updatedAt,
    target: normalizeMeasurementTargets(value?.target),
  };
}

function normalizeBijirisPostFile(file, kind = "photo") {
  if (!file || typeof file !== "object") return null;
  const normalizedKind = String(file.kind || kind || "").trim() || kind;
  const name = String(file.name || "").trim();
  const fileId = String(file.fileId || "").trim();
  const url = String(file.url || "").trim();
  const previewUrl = String(file.previewUrl || "").trim();
  const downloadUrl = String(file.downloadUrl || "").trim();
  const thumbnailUrl = String(file.thumbnailUrl || "").trim();
  const dataUrl = typeof file.dataUrl === "string" ? file.dataUrl : "";
  const base64Data = typeof file.base64Data === "string" ? file.base64Data : "";
  if (!name && !fileId && !url && !previewUrl && !downloadUrl && !dataUrl && !base64Data) return null;
  return {
    kind: normalizedKind,
    name: name || (normalizedKind === "pdf" ? "資料" : "写真"),
    type: String(file.type || file.mimeType || "").trim(),
    fileId,
    url,
    previewUrl,
    downloadUrl,
    thumbnailUrl,
    dataUrl,
    base64Data,
  };
}

function normalizeBijirisPost(post) {
  return {
    id: String(post?.id || "").trim(),
    title: String(post?.title || "").trim(),
    category: String(post?.category || "").trim(),
    summary: String(post?.summary || "").trim(),
    body: String(post?.body || "").trim(),
    status: ["published", "draft", "archived"].includes(String(post?.status || "").trim())
      ? String(post.status).trim()
      : "draft",
    pinned: post?.pinned === true,
    createdAt: String(post?.createdAt || "").trim(),
    updatedAt: String(post?.updatedAt || "").trim(),
    publishedAt: String(post?.publishedAt || "").trim(),
    notifyCustomers: post?.notifyCustomers === true,
    notificationTitle: String(post?.notificationTitle || "").trim(),
    notificationBody: String(post?.notificationBody || "").trim(),
    photos: (Array.isArray(post?.photos) ? post.photos : [])
      .map((file) => normalizeBijirisPostFile(file, "photo"))
      .filter(Boolean),
    documents: (Array.isArray(post?.documents) ? post.documents : [])
      .map((file) => normalizeBijirisPostFile(file, "pdf"))
      .filter(Boolean),
  };
}

function indexCustomerProfiles(list) {
  const indexed = {};
  (Array.isArray(list) ? list : []).forEach((item) => {
    const profile = normalizeCustomerProfile(item);
    if (!profile.name) return;
    indexed[profile.name] = profile;
  });
  return indexed;
}

function getCustomerProfileByName(customerName) {
  return state.customerProfiles[String(customerName || "").trim()] || null;
}

function getCustomerMemberNumber(customerName) {
  return String(getCustomerProfileByName(customerName)?.memberNumber || "").trim().toUpperCase();
}

function getCustomerNameWithMember(customerName) {
  const memberNumber = getCustomerMemberNumber(customerName);
  return memberNumber ? `${memberNumber} / ${customerName}` : customerName;
}

function buildTicketInfoFromActiveTicketCard(ticketCard) {
  const normalized = normalizeActiveTicketCard(ticketCard);
  if (!normalized) return [];
  return [
    { label: "回数券", value: normalized.plan },
    { label: "何枚目", value: `${normalized.sheetNumber}枚目` },
    { label: "何回目", value: `${normalized.round}回目` },
  ];
}

function normalizeStatus(status) {
  return STATUS_LABELS[status] ? status : "new";
}

function normalizeSurveyStatus(status) {
  if (status === "draft") return "draft";
  if (status === "archived") return "archived";
  return "published";
}

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("show");
  setTimeout(() => toast.classList.remove("show"), 2600);
}

function reportClientError(source, error, detail = {}) {
  void api.logError?.(source, error, detail);
}

function setPage(page) {
  document.querySelectorAll(".page").forEach((node) => {
    node.classList.toggle("active", node.id === `page-${page}`);
  });
  document.querySelectorAll("[data-page]").forEach((button) => {
    button.classList.toggle("active", button.dataset.page === page);
  });
}

function setLoggedIn(loggedIn) {
  loginView.hidden = loggedIn;
  adminView.hidden = !loggedIn;
  if (loggedIn) {
    setPage("dashboard");
    loginForm.reset();
  }
  loginSubmitButton.textContent = "ログイン";
}

function normalizeMemoRecord(record) {
  if (!record) return { latestMemo: "", entries: [] };
  if (typeof record === "string") {
    return {
      latestMemo: record,
      entries: record ? [{ at: "", memo: record }] : [],
    };
  }
  return {
    latestMemo: String(record.latestMemo || record.memo || "").trim(),
    entries: Array.isArray(record.entries) ? record.entries : [],
  };
}

function getCustomerMemoRecord(customerName) {
  return normalizeMemoRecord(state.customerMemos?.[customerName]);
}

function getFallbackSurveys() {
  const makeSurveys = window.MayumiDefaultSurveys;
  return typeof makeSurveys === "function" ? makeSurveys(new Date().toISOString()) : [];
}

async function loadAdminData() {
  const [
    adminInfoResult,
    surveysResult,
    responsesResult,
    measurementsResult,
    bijirisPostsResult,
    preferencesResult,
    logsResult,
    customerMemosResult,
  ] = await Promise.all([
    api.request("/api/admin/info", { token: state.token }),
    api.request("/api/admin/surveys", { token: state.token }),
    api.request("/api/admin/responses", { token: state.token }),
    api.request("/api/admin/measurements", { token: state.token }),
    api.request("/api/admin/bijiris-posts", { token: state.token }),
    api.request("/api/admin/preferences", { token: state.token }),
    api.request("/api/admin/logs", { token: state.token }),
    api.request("/api/admin/customer-memos", { token: state.token }),
  ]);
  state.adminInfo = adminInfoResult || null;
  state.customerProfiles = indexCustomerProfiles(state.adminInfo?.customerProfiles);
  state.surveys = (surveysResult.surveys || []).length
    ? surveysResult.surveys || []
    : getFallbackSurveys();
  const visibleSurveyIds = getVisibleSurveyIdSet();
  state.responses = (responsesResult.responses || [])
    .map((response) => ({
      status: "new",
      adminMemo: "",
      ...response,
    }))
    .filter((response) =>
      !visibleSurveyIds.size || visibleSurveyIds.has(String(response?.surveyId || "").trim()),
    );
  state.measurements = Array.isArray(measurementsResult?.measurements)
    ? measurementsResult.measurements.map(normalizeMeasurementRecord)
    : [];
  state.bijirisPosts = Array.isArray(bijirisPostsResult?.posts)
    ? bijirisPostsResult.posts.map(normalizeBijirisPost).filter((post) => post.id)
    : [];
  state.preferences = preferencesResult?.preferences || null;
  state.logs = logsResult || { auditLogs: [], errorLogs: [] };
  state.customerMemos = customerMemosResult?.memos || {};
  state.adminUsers = Array.isArray(state.adminInfo?.adminUsers) ? state.adminInfo.adminUsers : [];
  if (!state.selectedAnalyticsSurveyId && state.surveys[0]) {
    state.selectedAnalyticsSurveyId = state.surveys[0].id;
  }
  renderAll();
}

function renderAll() {
  renderNavigation();
  renderDashboard();
  renderBijirisManager();
  renderSurveyManager();
  renderFilters();
  renderCustomerManagement();
  renderResponses();
  renderSettings();
  document.querySelector("#adminUrlBox").textContent = window.location.href;
}

function renderNavigation() {
  const unread = state.responses.filter((response) => normalizeStatus(response.status) === "new").length;
  document.querySelectorAll('[data-page="responses"]').forEach((button) => {
    const baseLabel = button.classList.contains("bottom-nav-item") ? "👥<span>回答</span>" : "👥 回答管理";
    if (!unread) {
      button.innerHTML = baseLabel;
      return;
    }
    button.innerHTML = button.classList.contains("bottom-nav-item")
      ? `👥<span>回答(${unread})</span>`
      : `👥 回答管理 <span class="inline-count">${unread}</span>`;
  });
}

function renderDashboard() {
  const activeResponses = state.responses.filter((response) => normalizeStatus(response.status) !== "trash");
  const customers = groupByCustomerFrom(activeResponses);
  const unread = activeResponses.filter((response) => normalizeStatus(response.status) === "new").length;
  document.querySelector("#statsGrid").innerHTML = `
    <div class="stat-card"><span>回答数</span><strong>${activeResponses.length}</strong></div>
    <div class="stat-card"><span>未対応</span><strong>${unread}</strong></div>
    <div class="stat-card"><span>回答者数</span><strong>${customers.length}</strong></div>
    <div class="stat-card"><span>アンケート数</span><strong>${state.surveys.length}</strong></div>
  `;

  const latest = activeResponses
    .slice()
    .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))
    .slice(0, 5);
  document.querySelector("#latestResponses").innerHTML = latest.length
    ? latest.map(renderCompactResponse).join("")
    : `<div class="empty">まだ回答はありません。</div>`;

  const summary = state.surveys.map((survey) => ({
    title: survey.title,
    count: activeResponses.filter((response) => response.surveyId === survey.id).length,
  }));
  document.querySelector("#surveySummary").innerHTML = summary.length
    ? summary
        .map(
          (item) => `
            <article class="card">
              <strong>${escapeHtml(item.title)}</strong>
              <div class="meta">${item.count}件</div>
            </article>
          `,
        )
        .join("")
      : `<div class="empty">アンケートはありません。</div>`;

  renderSurveyAnalytics();
  renderMonthlySurveyChart();
}

function getMonthlySurveyCounts() {
  const map = new Map();
  state.responses.forEach((response) => {
    if (normalizeStatus(response.status) === "trash") return;
    const month = new Date(response.submittedAt).toISOString().slice(0, 7);
    const current = map.get(month) || {};
    current[response.surveyTitle] = (current[response.surveyTitle] || 0) + 1;
    map.set(month, current);
  });
  return Array.from(map.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .slice(-6);
}

function renderMonthlySurveyChart() {
  const container = document.querySelector("#monthlySurveyChart");
  if (!container) return;
  const months = getMonthlySurveyCounts();
  const surveys = state.surveys.map((survey) => survey.title);
  if (!months.length) {
    container.innerHTML = `<div class="empty">まだ集計できる回答がありません。</div>`;
    return;
  }
  const max = Math.max(
    1,
    ...months.flatMap(([, counts]) => surveys.map((title) => Number(counts[title] || 0))),
  );
  container.innerHTML = months
    .map(
      ([month, counts]) => `
        <article class="answer-item">
          <strong>${escapeHtml(month)}</strong>
          <div class="monthly-bar-list">
            ${surveys
              .map((title) => {
                const count = Number(counts[title] || 0);
                return `
                  <div class="monthly-bar-row">
                    <span>${escapeHtml(title)}</span>
                    <div class="monthly-bar-track"><div class="monthly-bar-fill" style="width:${Math.round((count / max) * 100)}%"></div></div>
                    <strong>${count}</strong>
                  </div>
                `;
              })
              .join("")}
          </div>
        </article>
      `,
    )
    .join("");
}

function renderCompactResponse(response) {
  const status = normalizeStatus(response.status);
  return `
    <article class="card">
      <strong>${escapeHtml(getCustomerNameWithMember(response.customerName))}</strong>
      <div class="meta">${escapeHtml(response.surveyTitle)} / ${formatDate(response.submittedAt)}</div>
      <span class="badge ${status}">${STATUS_LABELS[status]}</span>
    </article>
  `;
}

function getCheckboxAnswerValues(answer) {
  return String(answer?.value || "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function getConcernAnalyticsRows(answerMapList) {
  const selectedOptionsByResponse = answerMapList.map((answerMap) => {
    const answer = answerMap.get(SESSION_CONCERN_QUESTION_ID);
    return new Set(getCheckboxAnswerValues(answer));
  });

  return SESSION_CONCERN_CATEGORIES.map((category) => {
    const optionRows = category.options.map((option) => {
      const count = selectedOptionsByResponse.reduce(
        (total, selectedOptions) => total + (selectedOptions.has(option) ? 1 : 0),
        0,
      );
      return { label: option, count };
    });

    return {
      label: category.label,
      count: selectedOptionsByResponse.reduce(
        (total, selectedOptions) =>
          total + (category.options.some((option) => selectedOptions.has(option)) ? 1 : 0),
        0,
      ),
      rows: optionRows,
    };
  });
}

function getSurveyAnalyticsSummary(surveyId) {
  const survey = findSurveyById(surveyId);
  if (!survey) return [];
  const responses = state.responses.filter(
    (response) => response.surveyId === surveyId && normalizeStatus(response.status) !== "trash",
  );
  const answerMapList = responses.map((response) => {
    const map = new Map();
    (response.answers || []).forEach((answer) => map.set(answer.questionId, answer));
    return map;
  });

  return survey.questions.map((question) => {
    if (question.type === "photo") {
      const count = answerMapList.reduce((total, answerMap) => {
        const answer = answerMap.get(question.id);
        return total + (Array.isArray(answer?.files) && answer.files.length ? 1 : 0);
      }, 0);
      return {
        question,
        rows: [{ label: "写真添付あり", count }],
      };
    }

    if (question.type === "checkbox" || question.type === "choice" || question.type === "rating") {
      if (question.id === SESSION_CONCERN_QUESTION_ID) {
        return {
          question,
          groups: getConcernAnalyticsRows(answerMapList),
        };
      }

      const counter = new Map();
      answerMapList.forEach((answerMap) => {
        const answer = answerMap.get(question.id);
        const values =
          question.type === "checkbox"
            ? getCheckboxAnswerValues(answer)
            : [String(answer?.value || "").trim()].filter(Boolean);
        values.forEach((value) => {
          counter.set(value, (counter.get(value) || 0) + 1);
        });
      });

      const rows = Array.from(counter.entries())
        .sort((a, b) => b[1] - a[1])
        .map(([label, count]) => ({ label, count }));
      return {
        question,
        rows: rows.length ? rows : [{ label: "未回答", count: 0 }],
      };
    }

    const filled = answerMapList.reduce((total, answerMap) => {
      return total + (String(answerMap.get(question.id)?.value || "").trim() ? 1 : 0);
    }, 0);
    return {
      question,
      rows: [{ label: "入力あり", count: filled }],
    };
  });
}

function renderSurveyAnalytics() {
  const analyticsFilter = document.querySelector("#analyticsSurveyFilter");
  const analyticsBody = document.querySelector("#surveyAnalytics");
  if (!analyticsFilter || !analyticsBody) return;

  const current = state.selectedAnalyticsSurveyId || analyticsFilter.value;
  analyticsFilter.innerHTML = state.surveys
    .map(
      (survey) => `
        <option value="${survey.id}" ${current === survey.id ? "selected" : ""}>${escapeHtml(survey.title)}</option>
      `,
    )
    .join("");
  state.selectedAnalyticsSurveyId = analyticsFilter.value || current || "";

  const survey = findSurveyById(state.selectedAnalyticsSurveyId);
  if (!survey) {
    analyticsBody.innerHTML = `<div class="empty">アンケートを選択してください。</div>`;
    return;
  }

  const analytics = getSurveyAnalyticsSummary(survey.id);
  analyticsBody.innerHTML = analytics
    .map(
      (item) => `
        <article class="answer-item">
          <strong>${escapeHtml(item.question.label)}</strong>
          ${
            Array.isArray(item.groups)
              ? `
                <div class="analytics-group-list">
                  ${item.groups
                    .map(
                      (group) => `
                        <section class="analytics-group-card">
                          <div class="analytics-group-head">
                            <strong>${escapeHtml(group.label)}</strong>
                            <span>${group.count}件</span>
                          </div>
                          <div class="analytics-list">
                            ${group.rows
                              .map(
                                (row) => `
                                  <div class="analytics-row">
                                    <span>${escapeHtml(row.label)}</span>
                                    <strong>${row.count}件</strong>
                                  </div>
                                `,
                              )
                              .join("")}
                          </div>
                        </section>
                      `,
                    )
                    .join("")}
                </div>
              `
              : `
                <div class="analytics-list">
                  ${item.rows
                    .map(
                      (row) => `
                        <div class="analytics-row">
                          <span>${escapeHtml(row.label)}</span>
                          <strong>${row.count}件</strong>
                        </div>
                      `,
                    )
                    .join("")}
                </div>
              `
          }
        </article>
      `,
    )
    .join("");

  analyticsFilter.onchange = () => {
    state.selectedAnalyticsSurveyId = analyticsFilter.value;
    renderSurveyAnalytics();
  };
}

function renderAnswerValue(answer) {
  const photoFiles = getPhotoFilesFromAnswer(answer);
  if (photoFiles.length) {
    return `
      <div class="photo-list">
        ${photoFiles
          .map((file) => {
            const preview = getPhotoPreviewSrc(file);
            return `
              <button
                class="photo-thumb lightbox-trigger"
                type="button"
                data-lightbox-src="${escapeHtml(getPhotoLightboxSrc(file))}"
                data-lightbox-title="${escapeHtml(file.name || "写真")}"
              >
                ${preview ? `<img src="${escapeHtml(preview)}" alt="${escapeHtml(file.name)}" />` : ""}
                <span>${escapeHtml(file.name || "写真")}</span>
                ${file.capturedAt ? `<span class="meta">撮影日: ${escapeHtml(formatDate(file.capturedAt))}</span>` : ""}
              </button>
            `;
          })
          .join("")}
      </div>
    `;
  }
  return escapeHtml(answer.value || "未回答");
}

function extractDriveFileId(url) {
  const value = String(url || "");
  const match = value.match(/\/d\/([^/]+)/) || value.match(/[?&]id=([^&]+)/);
  return match ? match[1] : "";
}

function buildDrivePreviewUrl(fileId) {
  return fileId ? `https://drive.google.com/uc?export=view&id=${encodeURIComponent(fileId)}` : "";
}

function buildDriveThumbnailUrl(fileId) {
  return fileId ? `https://drive.google.com/thumbnail?id=${encodeURIComponent(fileId)}&sz=w1200` : "";
}

function buildDriveDownloadUrl(fileId) {
  return fileId ? `https://drive.google.com/uc?export=download&id=${encodeURIComponent(fileId)}` : "";
}

function ensureDownloadablePhotoFile(file, index = 0) {
  if (!file) return null;
  const fallbackUrl = String(file.url || file.value || file.previewUrl || file.thumbnailUrl || "").trim();
  const driveFileId = String(
    file.fileId ||
    extractDriveFileId(file.downloadUrl || file.previewUrl || file.thumbnailUrl || file.url || file.value || ""),
  ).trim();
  return {
    ...file,
    name: String(file.name || `写真${index + 1}`).trim(),
    fileId: driveFileId || "",
    url: String(file.url || fallbackUrl).trim(),
    previewUrl: String(file.previewUrl || buildDrivePreviewUrl(driveFileId) || fallbackUrl).trim(),
    thumbnailUrl: String(file.thumbnailUrl || buildDriveThumbnailUrl(driveFileId) || fallbackUrl).trim(),
    downloadUrl: String(file.downloadUrl || buildDriveDownloadUrl(driveFileId) || fallbackUrl).trim(),
  };
}

function derivePhotoFileFromUrl(url, index = 0) {
  const normalizedUrl = String(url || "").trim();
  if (!normalizedUrl) return null;
  return ensureDownloadablePhotoFile({
    name: `写真${index + 1}`,
    url: normalizedUrl,
  }, index);
}

function getPhotoPreviewSrc(file) {
  return String(file?.thumbnailUrl || file?.previewUrl || file?.dataUrl || file?.url || "").trim();
}

function getPhotoLightboxSrc(file) {
  return String(file?.previewUrl || file?.thumbnailUrl || file?.dataUrl || file?.url || "").trim();
}

function createPdfThumbnailDataUrl(fileName) {
  const label = String(fileName || "PDF").trim().slice(0, 24);
  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="360" height="240" viewBox="0 0 360 240">
      <rect width="360" height="240" rx="18" fill="#fbf7f2"/>
      <rect x="20" y="20" width="320" height="200" rx="14" fill="#ffffff" stroke="#eadfd2"/>
      <rect x="40" y="38" width="74" height="28" rx="8" fill="#c95f50"/>
      <text x="77" y="57" text-anchor="middle" font-family="Arial, sans-serif" font-size="16" font-weight="700" fill="#ffffff">PDF</text>
      <text x="40" y="102" font-family="Arial, sans-serif" font-size="20" font-weight="700" fill="#5b493d">${escapeHtml(label)}</text>
      <text x="40" y="138" font-family="Arial, sans-serif" font-size="14" fill="#866f60">資料プレビュー</text>
    </svg>
  `;
  return `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(svg)}`;
}

function getBijirisDocumentThumbnailSrc(file) {
  return String(file?.thumbnailUrl || "").trim() || createPdfThumbnailDataUrl(file?.name);
}

function getPhotoFilesFromAnswer(answer) {
  if (Array.isArray(answer?.files) && answer.files.length) {
    return answer.files.map((file, index) => ensureDownloadablePhotoFile(file, index)).filter(Boolean);
  }
  const value = String(answer?.value || "").trim();
  if (!value || !value.includes("http")) {
    return [];
  }
  return value
    .split(",")
    .map((item) => item.trim())
    .filter((item) => /^https?:\/\//.test(item))
    .map((url, index) => derivePhotoFileFromUrl(url, index))
    .filter(Boolean);
}

function renderResponsePhotoPreview(response, limit = 3) {
  const photos = collectPhotosFromResponse(response).slice(0, limit);
  if (!photos.length) return "";
  return `
    <div class="history-photo-strip">
      ${photos
        .map((file, index) => {
          const preview = getPhotoPreviewSrc(file);
          const title = file.name || `写真${index + 1}`;
          if (!preview) {
            return `<span class="history-photo-badge">${escapeHtml(title)}</span>`;
          }
          return `
            <span class="history-photo-thumb" title="${escapeHtml(title)}">
              <img src="${escapeHtml(preview)}" alt="${escapeHtml(title)}" />
            </span>
          `;
        })
        .join("")}
      <span class="history-photo-badge">${collectPhotosFromResponse(response).length}枚</span>
    </div>
  `;
}

function getAnswerValueFromQuestionIds(answerMap, questionIds) {
  for (const questionId of questionIds) {
    const value = String(answerMap.get(questionId)?.value || "").trim();
    if (value) return value;
  }
  return "";
}

function getResponseTicketInfo(response) {
  const answerMap = new Map((response.answers || []).map((answer) => [answer.questionId, answer]));
  return [
    {
      label: "回数券",
      value: getAnswerValueFromQuestionIds(answerMap, TICKET_INFO_QUESTION_IDS.size),
    },
    {
      label: "何枚目",
      value: getAnswerValueFromQuestionIds(answerMap, TICKET_INFO_QUESTION_IDS.sheet),
    },
    {
      label: "何回目",
      value: getAnswerValueFromQuestionIds(answerMap, TICKET_INFO_QUESTION_IDS.round),
    },
  ].filter((item) => item.value);
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

function parseTicketCount(ticketValue) {
  const normalized = String(ticketValue || "");
  if (normalized.includes("10")) return 10;
  if (normalized.includes("6")) return 6;
  const matched = normalized.match(/\d+/);
  return matched ? Number(matched[0]) : 0;
}

function parseTicketStep(stepValue) {
  const matched = String(stepValue || "").match(/\d+/);
  return matched ? Number(matched[0]) : 0;
}

function getTicketProgressInfo(ticketInfo) {
  const ticketMap = new Map((ticketInfo || []).map((item) => [item.label, item.value]));
  const ticketCount = parseTicketCount(ticketMap.get("回数券") || "");
  const currentRound = parseTicketStep(ticketMap.get("何回目") || "");
  return {
    sheetLabel: ticketMap.get("何枚目") || "",
    ticketCount,
    currentRound,
  };
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

function renderTicketStampPanel(ticketInfo) {
  if (!ticketInfo.length) return "";
  const { sheetLabel, ticketCount, currentRound } = getTicketProgressInfo(ticketInfo);
  return `
    <div class="ticket-stamp-panel">
      ${renderTicketStampList(ticketInfo)}
      ${
        ticketCount
          ? `
            <div class="ticket-progress-card">
              <div class="ticket-progress-head">
                <strong>${escapeHtml(sheetLabel || "-")}</strong>
                <span>${currentRound || 0} / ${ticketCount}</span>
              </div>
              ${renderTicketStampProgress(ticketCount, currentRound)}
            </div>
          `
          : ""
      }
    </div>
  `;
}

function renderResponseTicketInfo(response) {
  const ticketInfo = getResponseTicketInfo(response);
  if (!ticketInfo.length) return "";
  return `
    <div class="answer-item ticket-info-panel">
      <strong>回答者情報</strong>
      ${renderTicketStampPanel(ticketInfo)}
    </div>
  `;
}

function groupResponsesBySurvey(responses) {
  const map = new Map();
  (Array.isArray(responses) ? responses : []).forEach((response) => {
    const key = String(response.surveyId || response.surveyTitle || response.id);
    const current = map.get(key) || {
      surveyId: response.surveyId || key,
      surveyTitle: response.surveyTitle || "アンケート",
      latestAt: response.submittedAt,
      count: 0,
      unreadCount: 0,
      responses: [],
    };
    current.responses.push(response);
    current.count += 1;
    if (normalizeStatus(response.status) === "new") current.unreadCount += 1;
    if (new Date(response.submittedAt) > new Date(current.latestAt)) {
      current.latestAt = response.submittedAt;
    }
    map.set(key, current);
  });
  return Array.from(map.values()).sort((a, b) => new Date(b.latestAt) - new Date(a.latestAt));
}

function getCurrentTicketInfoForCustomer(customerName) {
  const profileTicketInfo = buildTicketInfoFromActiveTicketCard(
    getCustomerProfileByName(customerName)?.activeTicketCard,
  );
  if (profileTicketInfo.length) {
    return profileTicketInfo;
  }
  const latestTicketResponse = state.responses
    .filter(
      (response) =>
        response.customerName === customerName &&
        normalizeStatus(response.status) !== "trash" &&
        getResponseTicketInfo(response).length,
    )
    .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))[0];
  return latestTicketResponse ? getResponseTicketInfo(latestTicketResponse) : [];
}

function getActiveResponses() {
  return state.responses.filter((response) => normalizeStatus(response.status) !== "trash");
}

function getActiveCustomerResponses(customerName) {
  return getActiveResponses()
    .filter((response) => response.customerName === customerName)
    .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt));
}

function getCustomerDirectoryEntries() {
  const map = new Map();
  Object.values(state.customerProfiles || {}).forEach((profile) => {
    if (!profile?.name) return;
    map.set(profile.name, {
      name: profile.name,
      memberNumber: profile.memberNumber || "",
      latestAt: profile.updatedAt || "",
      count: 0,
      surveys: new Map(),
      hasResponses: false,
    });
  });
  getActiveResponses().forEach((response) => {
    const existing = map.get(response.customerName) || {
      name: response.customerName,
      memberNumber: getCustomerMemberNumber(response.customerName) || response.customerMemberNumber || "",
      latestAt: response.submittedAt,
      count: 0,
      surveys: new Map(),
      hasResponses: false,
    };
    existing.memberNumber =
      existing.memberNumber || response.customerMemberNumber || getCustomerMemberNumber(response.customerName) || "";
    existing.count += 1;
    existing.hasResponses = true;
    existing.latestAt =
      !existing.latestAt || new Date(response.submittedAt) > new Date(existing.latestAt)
        ? response.submittedAt
        : existing.latestAt;
    existing.surveys.set(response.surveyTitle, (existing.surveys.get(response.surveyTitle) || 0) + 1);
    map.set(response.customerName, existing);
  });
  return Array.from(map.values()).sort((a, b) => {
    if (a.latestAt && b.latestAt) return new Date(b.latestAt) - new Date(a.latestAt);
    if (a.latestAt) return -1;
    if (b.latestAt) return 1;
    return String(a.memberNumber || "").localeCompare(String(b.memberNumber || "")) || a.name.localeCompare(b.name);
  });
}

function getFilteredCustomerDirectoryEntries() {
  const keyword = String(document.querySelector("#customerKeywordFilter")?.value || "").trim().toLowerCase();
  const ticketPlan = String(document.querySelector("#customerTicketPlanFilter")?.value || "").trim();
  const responseState = String(document.querySelector("#customerResponseStateFilter")?.value || "").trim();
  return getCustomerDirectoryEntries()
    .filter((entry) => {
      if (!keyword) return true;
      const profile = getCustomerProfileByName(entry.name);
      const haystack = [
        entry.name,
        entry.memberNumber,
        profile?.nameKana,
        getCurrentTicketInfoForCustomer(entry.name).map((item) => item.value).join(" "),
      ]
        .filter(Boolean)
        .join(" ")
        .toLowerCase();
      return haystack.includes(keyword);
    })
    .filter((entry) => {
      if (!ticketPlan) return true;
      return getCurrentTicketInfoForCustomer(entry.name).some(
        (item) => item.label === "回数券" && item.value === ticketPlan,
      );
    })
    .filter((entry) => {
      if (responseState === "with") return entry.hasResponses;
      if (responseState === "without") return !entry.hasResponses;
      return true;
    });
}

function formatResponseEntryTitle(response) {
  const ticketInfo = getResponseTicketInfo(response);
  if (ticketInfo.length) {
    return ticketInfo
      .map((item, index) => (index === 0 ? `${item.label} ${item.value}` : item.value))
      .join(" / ");
  }
  return response.surveyTitle || "回答詳細";
}

function renderCustomerSummaryCard(customerName, responses) {
  const latestResponse = responses[0] || null;
  const latestMeasurement = getCustomerMeasurements(customerName)[0] || null;
  const ticketInfo = getCurrentTicketInfoForCustomer(customerName);
  const surveyCount = new Set(responses.map((response) => response.surveyId || response.surveyTitle || response.id)).size;
  const profile = getCustomerProfileByName(customerName);
  return `
    <article class="answer-item">
      <strong>${escapeHtml(getCustomerNameWithMember(customerName))}</strong>
      <div class="meta">フリガナ: ${escapeHtml(profile?.nameKana || "-")}</div>
      ${renderCustomerPushStatus(profile?.pushStatus)}
      <div class="meta">回答数: ${responses.length}件 / アンケート種類: ${surveyCount}件</div>
      <div class="meta">最新回答: ${latestResponse ? `${escapeHtml(latestResponse.surveyTitle)} / ${formatDate(latestResponse.submittedAt)}` : "-"}</div>
      <div class="meta">最新測定: ${latestMeasurement ? `${escapeHtml(formatDateOnly(latestMeasurement.measuredAt))} / WHR ${escapeHtml(formatWhr(latestMeasurement.whr))}` : "-"}</div>
      ${
        ticketInfo.length
          ? renderTicketStampPanel(ticketInfo)
          : `<div class="meta">現在の回数券スタンプ情報はありません。</div>`
      }
    </article>
  `;
}

function getCurrentTicketDraft(customerName) {
  const profile = getCustomerProfileByName(customerName);
  const ticketMap = new Map(getCurrentTicketInfoForCustomer(customerName).map((item) => [item.label, item.value]));
  return {
    memberNumber: profile?.memberNumber || "",
    ticketPlan: ticketMap.get("回数券") || "",
    ticketSheet: ticketMap.get("何枚目") || "",
    ticketRound: ticketMap.get("何回目") || "",
  };
}

function renderCustomerEditorCard(customerName) {
  const draft = getCurrentTicketDraft(customerName);
  return `
    <article class="answer-item">
      <strong>顧客情報を編集</strong>
      <form id="customerProfileForm" class="stack" data-customer-name="${escapeHtml(customerName)}">
        <div class="customer-editor-grid">
          <label>
            お名前
            <input name="customerName" type="text" value="${escapeHtml(customerName)}" required />
          </label>
          <label>
            会員番号
            <input name="memberNumber" type="text" value="${escapeHtml(draft.memberNumber)}" placeholder="自動採番" />
          </label>
          <label>
            回数券の種類
            <select name="ticketPlan">
              ${CUSTOMER_TICKET_PLAN_OPTIONS
                .map(
                  (option) => `
                    <option value="${escapeHtml(option)}" ${draft.ticketPlan === option ? "selected" : ""}>
                      ${escapeHtml(option || "未設定")}
                    </option>
                  `,
                )
                .join("")}
            </select>
          </label>
          <label>
            回数券の何枚目
            <select name="ticketSheet">
              <option value="">未設定</option>
              ${CUSTOMER_TICKET_SHEET_OPTIONS
                .map(
                  (option) => `
                    <option value="${escapeHtml(option)}" ${draft.ticketSheet === option ? "selected" : ""}>
                      ${escapeHtml(option)}
                    </option>
                  `,
                )
                .join("")}
            </select>
          </label>
          <label>
            回数券の何回目
            <select name="ticketRound">
              <option value="">未設定</option>
              ${CUSTOMER_TICKET_ROUND_OPTIONS
                .map(
                  (option) => `
                    <option value="${escapeHtml(option)}" ${draft.ticketRound === option ? "selected" : ""}>
                      ${escapeHtml(option)}
                    </option>
                  `,
                )
                .join("")}
            </select>
          </label>
        </div>
        <div class="action-row">
          <button class="secondary-button" type="submit">顧客情報を保存</button>
          <button class="secondary-button danger-button" type="button" id="deleteCustomerButton">顧客を削除</button>
        </div>
      </form>
    </article>
  `;
}

function getMeasurementMetric(metricKey) {
  return MEASUREMENT_METRICS.find((metric) => metric.key === metricKey) || null;
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
  const sign = normalized > 0 ? "+" : normalized < 0 ? "" : "+/-";
  const text =
    normalized === 0
      ? `0.0${unit}`
      : `${sign}${Math.abs(normalized).toFixed(1)}${unit}`;
  return `<span class="delta-badge ${className}">${escapeHtml(text)}</span>`;
}

function formatMeasurementGapToTarget(value, target) {
  const normalizedValue = normalizeMeasurementValue(value);
  const normalizedTarget = normalizeMeasurementValue(target);
  if (normalizedValue === "" || normalizedTarget === "") return '<span class="delta-badge neutral">未設定</span>';
  const diff = roundMeasurementDelta(normalizedValue - normalizedTarget);
  if (diff <= 0) {
    return `<span class="delta-badge achieved">達成</span>`;
  }
  return `<span class="delta-badge remaining">残り${escapeHtml(diff.toFixed(1))}cm</span>`;
}

function formatWhr(value) {
  if (value === "" || value === null || value === undefined || !Number.isFinite(Number(value))) return "-";
  return Number(value).toFixed(3);
}

function getTodayDateInputValue() {
  const now = new Date();
  const local = new Date(now.getTime() - now.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 10);
}

function getCustomerMeasurements(customerName) {
  return state.measurements
    .filter((measurement) => measurement.customerName === customerName)
    .sort((a, b) => {
      const measuredDiff = new Date(b.measuredAt) - new Date(a.measuredAt);
      if (measuredDiff !== 0) return measuredDiff;
      return new Date(b.updatedAt || b.createdAt || 0) - new Date(a.updatedAt || a.createdAt || 0);
    });
}

function getMeasurementTargetsForCustomer(customerName) {
  return getCustomerProfileByName(customerName)?.measurementTargets || null;
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

function buildMeasurementHistoryRows(measurements, targets) {
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
      const targetValue = normalizeMeasurementValue(targets?.[metric.key]);
      metrics[metric.key] = {
        value,
        previousDelta:
          value === "" || previousValue === "" ? "" : roundMeasurementDelta(value - previousValue),
        firstDelta: value === "" || firstValue === "" ? "" : roundMeasurementDelta(value - firstValue),
        targetGap: value === "" || targetValue === "" ? "" : roundMeasurementDelta(value - targetValue),
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

function renderMeasurementSummaryCards(measurements, targets) {
  const latest = measurements[0] || null;
  const first = measurements[measurements.length - 1] || null;
  const thighGap =
    latest && latest.thighRight !== "" && latest.thighLeft !== ""
      ? roundMeasurementDelta(Math.abs(latest.thighRight - latest.thighLeft))
      : "";
  return `
    <div class="measurement-summary-grid">
      <article class="measurement-summary-card">
        <div class="measurement-summary-label">最新測定日</div>
        <strong>${escapeHtml(formatDateOnly(latest?.measuredAt || ""))}</strong>
        <div class="meta">履歴 ${measurements.length}件</div>
      </article>
      <article class="measurement-summary-card">
        <div class="measurement-summary-label">最新 WHR</div>
        <strong>${escapeHtml(formatWhr(latest?.whr))}</strong>
        <div class="meta">初回 ${escapeHtml(formatWhr(first?.whr))}</div>
      </article>
      <article class="measurement-summary-card">
        <div class="measurement-summary-label">太もも左右差</div>
        <strong>${thighGap === "" ? "-" : `${escapeHtml(thighGap.toFixed(1))}cm`}</strong>
        <div class="meta">左右のバランス確認用</div>
      </article>
      ${MEASUREMENT_METRICS.map((metric) => {
        const value = latest?.[metric.key];
        const target = targets?.[metric.key];
        const firstValue = first?.[metric.key];
        const firstDelta =
          normalizeMeasurementValue(value) === "" || normalizeMeasurementValue(firstValue) === ""
            ? ""
            : roundMeasurementDelta(normalizeMeasurementValue(value) - normalizeMeasurementValue(firstValue));
        return `
          <article class="measurement-summary-card metric-${escapeHtml(metric.key)}">
            <div class="measurement-summary-label">${escapeHtml(metric.label)}</div>
            <strong>${escapeHtml(formatMeasurementValue(value))}</strong>
            <div class="meta">初回比 ${formatMeasurementDelta(firstDelta)}</div>
            <div class="meta-inline">${formatMeasurementGapToTarget(value, target)}</div>
          </article>
        `;
      }).join("")}
    </div>
  `;
}

function renderMeasurementTargetsForm(customerName) {
  const targets = getMeasurementTargetsForCustomer(customerName) || {};
  return `
    <article class="answer-item">
      <strong>目標値</strong>
      <div class="meta">各部位の目標数値を設定すると、グラフに目標ラインを表示します。</div>
      <form id="measurementTargetForm" class="stack" data-customer-name="${escapeHtml(customerName)}">
        <div class="measurement-target-grid">
          ${MEASUREMENT_METRICS.map((metric) => `
            <label>
              ${escapeHtml(metric.label)}
              <input
                name="target_${escapeHtml(metric.key)}"
                type="number"
                step="0.1"
                min="0"
                inputmode="decimal"
                value="${targets?.[metric.key] === "" || targets?.[metric.key] === undefined ? "" : escapeHtml(String(targets[metric.key]))}"
                placeholder="目標 cm"
              />
            </label>
          `).join("")}
        </div>
        <div class="action-row">
          <button class="secondary-button" type="submit">目標値を保存</button>
        </div>
      </form>
    </article>
  `;
}

function renderMeasurementForm(customerName) {
  const editingMeasurement = getCustomerMeasurements(customerName).find(
    (measurement) => measurement.id === state.selectedMeasurementEditId,
  ) || null;
  return `
    <article class="answer-item">
      <strong>${editingMeasurement ? "測定データを編集" : "測定データを追加"}</strong>
      <form id="measurementForm" class="stack" data-customer-name="${escapeHtml(customerName)}">
        <input name="measurementId" type="hidden" value="${escapeHtml(editingMeasurement?.id || "")}" />
        <div class="measurement-target-grid">
          <label>
            測定日
            <input name="measuredAt" type="date" required value="${escapeHtml(editingMeasurement?.measuredAt || getTodayDateInputValue())}" />
          </label>
          <label>
            ウエスト (cm)
            <input name="waist" type="number" step="0.1" min="0" inputmode="decimal" value="${editingMeasurement?.waist === "" ? "" : escapeHtml(String(editingMeasurement?.waist || ""))}" />
          </label>
          <label>
            ヒップ (cm)
            <input name="hip" type="number" step="0.1" min="0" inputmode="decimal" value="${editingMeasurement?.hip === "" ? "" : escapeHtml(String(editingMeasurement?.hip || ""))}" />
          </label>
          <label>
            太もも 右 (cm)
            <input name="thighRight" type="number" step="0.1" min="0" inputmode="decimal" value="${editingMeasurement?.thighRight === "" ? "" : escapeHtml(String(editingMeasurement?.thighRight || ""))}" />
          </label>
          <label>
            太もも 左 (cm)
            <input name="thighLeft" type="number" step="0.1" min="0" inputmode="decimal" value="${editingMeasurement?.thighLeft === "" ? "" : escapeHtml(String(editingMeasurement?.thighLeft || ""))}" />
          </label>
        </div>
        <label>
          分析メモ・コメント
          <textarea name="memo" rows="3" placeholder="変化の背景や次回の見立てを記録">${escapeHtml(editingMeasurement?.memo || "")}</textarea>
        </label>
        <div class="action-row">
          <button class="secondary-button" type="submit">${editingMeasurement ? "測定を更新" : "測定を保存"}</button>
          ${
            editingMeasurement
              ? '<button class="secondary-button" type="button" data-cancel-measurement-edit>編集をやめる</button>'
              : ""
          }
        </div>
      </form>
    </article>
  `;
}

function renderMeasurementLineChart(title, measurements, metrics, options = {}) {
  const visibleMetrics = metrics.filter((metric) => state.measurementMetricVisibility[metric.key] !== false);
  if (!measurements.length) {
    return `
      <article class="answer-item">
        <strong>${escapeHtml(title)}</strong>
        <div class="empty">まだグラフを表示できる測定データがありません。</div>
      </article>
    `;
  }
  if (!visibleMetrics.length) {
    return `
      <article class="answer-item">
        <strong>${escapeHtml(title)}</strong>
        <div class="empty">表示する項目を選択してください。</div>
      </article>
    `;
  }

  const chronological = measurements
    .slice()
    .sort((a, b) => new Date(a.measuredAt) - new Date(b.measuredAt));
  const targetMap = options.targets || {};
  const values = [];
  chronological.forEach((measurement) => {
    visibleMetrics.forEach((metric) => {
      const value = normalizeMeasurementValue(measurement[metric.key]);
      if (value !== "") values.push(Number(value));
    });
  });
  visibleMetrics.forEach((metric) => {
    const target = normalizeMeasurementValue(targetMap?.[metric.key]);
    if (target !== "") values.push(Number(target));
  });
  if (!values.length) {
    return `
      <article class="answer-item">
        <strong>${escapeHtml(title)}</strong>
        <div class="empty">表示できる数値がまだありません。</div>
      </article>
    `;
  }

  const chartWidth = 760;
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

  const tickCount = 4;
  const grid = Array.from({ length: tickCount + 1 }, (_, index) => {
    const value = domainMin + (domainRange * index) / tickCount;
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

  const targetLines = visibleMetrics
    .map((metric) => {
      const target = normalizeMeasurementValue(targetMap?.[metric.key]);
      if (target === "") return "";
      const y = yForValue(target);
      return `
        <g>
          <line
            x1="${padding.left}"
            y1="${y}"
            x2="${chartWidth - padding.right}"
            y2="${y}"
            stroke="${metric.color}"
            stroke-width="1.5"
            stroke-dasharray="6 6"
            opacity="0.6"
          />
          <text x="${chartWidth - padding.right}" y="${y - 6}" text-anchor="end" class="measurement-target-label" fill="${metric.color}">
            ${escapeHtml(metric.label)} 目標 ${escapeHtml(target.toFixed(1))}${escapeHtml(options.unit || "")}
          </text>
        </g>
      `;
    })
    .join("");

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
      <text
        x="${xForIndex(index)}"
        y="${chartHeight - 14}"
        text-anchor="middle"
        class="measurement-axis-label"
      >
        ${escapeHtml(formatDateOnly(measurement.measuredAt).slice(5))}
      </text>
    `)
    .join("");

  const legend = visibleMetrics
    .map((metric) => `
      <label class="measurement-legend-item">
        <span class="measurement-legend-dot" style="background:${metric.color}"></span>
        ${escapeHtml(metric.label)}
      </label>
    `)
    .join("");

  return `
    <article class="answer-item">
      <strong>${escapeHtml(title)}</strong>
      <div class="measurement-legend">${legend}</div>
      <div class="measurement-chart-shell">
        <svg class="measurement-chart" viewBox="0 0 ${chartWidth} ${chartHeight}" role="img" aria-label="${escapeHtml(title)}">
          ${grid}
          ${targetLines}
          ${series}
          ${xLabels}
        </svg>
      </div>
    </article>
  `;
}

function renderMeasurementHistoryTable(measurements, targets) {
  const rows = buildMeasurementHistoryRows(measurements, targets);
  if (!rows.length) {
    return `
      <article class="answer-item">
        <strong>測定履歴</strong>
        <div class="empty">まだ測定履歴はありません。</div>
      </article>
    `;
  }
  return `
    <article class="answer-item">
      <strong>測定履歴</strong>
      <div class="meta">新しい測定が上に表示されます。前回比・初回比・目標との差を自動計算しています。</div>
      <div class="measurement-table-wrap">
        <table class="measurement-table">
          <thead>
            <tr>
              <th>測定日</th>
              ${MEASUREMENT_METRICS.map((metric) => `<th>${escapeHtml(metric.label)}</th>`).join("")}
              <th>WHR</th>
              <th>左右差</th>
              <th>メモ</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            ${rows
              .map((row) => `
                <tr>
                  <td>
                    <strong>${escapeHtml(formatDateOnly(row.measuredAt))}</strong>
                  </td>
                  ${MEASUREMENT_METRICS.map((metric) => {
                    const metricRow = row.metrics[metric.key];
                    return `
                      <td>
                        <div class="measurement-cell-main">${escapeHtml(formatMeasurementValue(metricRow.value))}</div>
                        <div class="measurement-cell-sub">前回 ${formatMeasurementDelta(metricRow.previousDelta)}</div>
                        <div class="measurement-cell-sub">初回 ${formatMeasurementDelta(metricRow.firstDelta)}</div>
                        <div class="measurement-cell-sub">目標 ${formatMeasurementGapToTarget(metricRow.value, targets?.[metric.key])}</div>
                      </td>
                    `;
                  }).join("")}
                  <td>${escapeHtml(formatWhr(row.whr))}</td>
                  <td>${row.leftRightGap === "" ? "-" : `${escapeHtml(row.leftRightGap.toFixed(1))}cm`}</td>
                  <td class="measurement-memo-cell">${escapeHtml(row.memo || "-")}</td>
                  <td>
                    <div class="measurement-row-actions">
                      <button class="secondary-button" type="button" data-edit-measurement="${escapeHtml(row.id)}">編集</button>
                      <button class="secondary-button danger-button" type="button" data-delete-measurement="${escapeHtml(row.id)}">削除</button>
                    </div>
                  </td>
                </tr>
              `)
              .join("")}
          </tbody>
        </table>
      </div>
    </article>
  `;
}

function renderMeasurementSection(customerName) {
  const allMeasurements = getCustomerMeasurements(customerName);
  const filteredMeasurements = filterMeasurementsByPeriod(allMeasurements, state.selectedMeasurementPeriod);
  const targets = getMeasurementTargetsForCustomer(customerName);
  return `
    <section class="measurement-section">
      ${renderMeasurementSummaryCards(allMeasurements, targets)}
      ${renderMeasurementTargetsForm(customerName)}
      ${renderMeasurementForm(customerName)}
      <article class="answer-item">
        <div class="stage-head">
          <div>
            <strong>推移グラフ</strong>
            <div class="meta">期間を絞り込み、表示する部位を切り替えられます。</div>
          </div>
          <label>
            表示期間
            <select id="measurementPeriodFilter">
              ${MEASUREMENT_PERIOD_OPTIONS.map((option) => `
                <option value="${escapeHtml(option.value)}" ${state.selectedMeasurementPeriod === option.value ? "selected" : ""}>
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
                ${state.measurementMetricVisibility[metric.key] !== false ? "checked" : ""}
              />
              ${escapeHtml(metric.label)}
            </label>
          `).join("")}
          <label class="measurement-toggle-chip">
            <input
              type="checkbox"
              data-toggle-measurement-metric="whr"
              ${state.measurementMetricVisibility.whr !== false ? "checked" : ""}
            />
            WHR
          </label>
        </div>
      </article>
      ${renderMeasurementLineChart("部位別の推移", filteredMeasurements, MEASUREMENT_METRICS, {
        targets,
        unit: "cm",
      })}
      ${
        state.measurementMetricVisibility.whr !== false
          ? renderMeasurementLineChart(
              "WHR の推移",
              filteredMeasurements,
              [{ key: "whr", label: "WHR", shortLabel: "WHR", color: "#7a5b9f", unit: "" }],
              {
                compact: true,
                pad: 0.03,
                valueFormatter: (value) => Number(value).toFixed(3),
              },
            )
          : ""
      }
      ${renderMeasurementHistoryTable(filteredMeasurements, targets)}
    </section>
  `;
}

async function saveCustomerMeasurementTargets(form) {
  const customerName = String(form.dataset.customerName || "").trim();
  if (!customerName) return;
  const currentDraft = getCurrentTicketDraft(customerName);
  const formData = new FormData(form);
  const measurementTargets = MEASUREMENT_METRICS.reduce((accumulator, metric) => {
    accumulator[metric.key] = String(formData.get(`target_${metric.key}`) || "").trim();
    return accumulator;
  }, {});
  const submitButton = form.querySelector('button[type="submit"]');
  if (submitButton) {
    submitButton.disabled = true;
    submitButton.textContent = "保存中";
  }
  try {
    await api.request(`/api/admin/customers/${encodeURIComponent(customerName)}`, {
      method: "PUT",
      token: state.token,
      body: {
        name: customerName,
        memberNumber: currentDraft.memberNumber,
        ticketPlan: currentDraft.ticketPlan,
        ticketSheet: currentDraft.ticketSheet,
        ticketRound: currentDraft.ticketRound,
        measurementTargets,
      },
    });
    await loadAdminData();
    state.selectedCustomerViewName = customerName;
    setPage("customers");
    showToast("目標値を保存しました。");
  } catch (error) {
    showToast(error.message || "目標値を保存できませんでした。");
  } finally {
    if (submitButton) {
      submitButton.disabled = false;
      submitButton.textContent = "目標値を保存";
    }
  }
}

async function saveMeasurement(form) {
  const customerName = String(form.dataset.customerName || "").trim();
  if (!customerName) return;
  const formData = new FormData(form);
  const measurementId = String(formData.get("measurementId") || "").trim();
  const payload = {
    measuredAt: String(formData.get("measuredAt") || "").trim(),
    waist: String(formData.get("waist") || "").trim(),
    hip: String(formData.get("hip") || "").trim(),
    thighRight: String(formData.get("thighRight") || "").trim(),
    thighLeft: String(formData.get("thighLeft") || "").trim(),
    memo: String(formData.get("memo") || "").trim(),
  };
  if (!payload.measuredAt) {
    showToast("測定日を入力してください。");
    return;
  }
  if (![payload.waist, payload.hip, payload.thighRight, payload.thighLeft].some(Boolean)) {
    showToast("測定値を1つ以上入力してください。");
    return;
  }
  const submitButton = form.querySelector('button[type="submit"]');
  if (submitButton) {
    submitButton.disabled = true;
    submitButton.textContent = measurementId ? "更新中" : "保存中";
  }
  try {
    if (measurementId) {
      await api.request(`/api/admin/measurements/${encodeURIComponent(measurementId)}`, {
        method: "PUT",
        token: state.token,
        body: payload,
      });
    } else {
      await api.request(`/api/admin/customers/${encodeURIComponent(customerName)}/measurements`, {
        method: "POST",
        token: state.token,
        body: payload,
      });
    }
    state.selectedMeasurementEditId = "";
    await loadAdminData();
    state.selectedCustomerViewName = customerName;
    setPage("customers");
    showToast(measurementId ? "測定データを更新しました。" : "測定データを保存しました。");
  } catch (error) {
    showToast(error.message || "測定データを保存できませんでした。");
  } finally {
    if (submitButton) {
      submitButton.disabled = false;
      submitButton.textContent = measurementId ? "測定を更新" : "測定を保存";
    }
  }
}

async function deleteMeasurement(measurementId) {
  if (!measurementId) return;
  if (!confirm("この測定データを削除しますか？")) return;
  try {
    await api.request(`/api/admin/measurements/${encodeURIComponent(measurementId)}`, {
      method: "DELETE",
      token: state.token,
    });
    if (state.selectedMeasurementEditId === measurementId) {
      state.selectedMeasurementEditId = "";
    }
    await loadAdminData();
    showToast("測定データを削除しました。");
  } catch (error) {
    showToast(error.message || "測定データを削除できませんでした。");
  }
}

async function saveCustomerProfile(form) {
  const currentName = String(form.dataset.customerName || "").trim();
  const formData = new FormData(form);
  const name = String(formData.get("customerName") || "").trim();
  const memberNumber = String(formData.get("memberNumber") || "").trim().toUpperCase();
  const ticketPlan = String(formData.get("ticketPlan") || "").trim();
  const ticketSheet = String(formData.get("ticketSheet") || "").trim();
  const ticketRound = String(formData.get("ticketRound") || "").trim();
  if (!name) {
    showToast("お名前を入力してください。");
    return;
  }
  const hasAnyTicketValue = Boolean(ticketPlan || ticketSheet || ticketRound);
  if (hasAnyTicketValue && !(ticketPlan && ticketSheet && ticketRound)) {
    showToast("回数券情報は種類・何枚目・何回目をすべて選択してください。");
    return;
  }
  const submitButton = form.querySelector('button[type="submit"]');
  if (submitButton) {
    submitButton.disabled = true;
    submitButton.textContent = "保存中";
  }
  try {
    const result = await api.request(`/api/admin/customers/${encodeURIComponent(currentName)}`, {
      method: "PUT",
      token: state.token,
      body: {
        name,
        memberNumber,
        ticketPlan,
        ticketSheet,
        ticketRound,
      },
    });
    state.selectedCustomerViewName = result.customerName || name;
    await loadAdminData();
    setPage("customers");
    showToast("顧客情報を保存しました。");
  } catch (error) {
    showToast(error.message || "顧客情報を保存できませんでした。");
  } finally {
    if (submitButton) {
      submitButton.disabled = false;
      submitButton.textContent = "顧客情報を保存";
    }
  }
}

async function deleteCustomerProfile(customerName) {
  if (!customerName) return;
  if (!confirm("この顧客情報と関連する回答履歴をすべて削除しますか？")) return;
  try {
    await api.request(`/api/admin/customers/${encodeURIComponent(customerName)}`, {
      method: "DELETE",
      token: state.token,
    });
    state.selectedCustomerViewName = "";
    state.selectedCustomerViewSurveyId = "";
    state.selectedCustomerViewResponseId = "";
    state.selectedCustomerTimelineOpen = false;
    state.selectedMeasurementEditId = "";
    await loadAdminData();
    setPage("customers");
    showToast("顧客情報を削除しました。");
  } catch (error) {
    showToast(error.message || "顧客情報を削除できませんでした。");
  }
}

function renderReadOnlyResponseCard(response) {
  const status = normalizeStatus(response.status);
  const survey = findSurveyById(response.surveyId);
  const displayAnswers = getDisplayAnswers(response, survey);
  return `
    <article class="response-card response-card-readonly">
      <div class="response-head">
        <div>
          <strong>${escapeHtml(getCustomerNameWithMember(response.customerName))}</strong>
          <div class="meta">${formatDate(response.submittedAt)}</div>
          <div class="meta">${escapeHtml(response.surveyTitle)}</div>
        </div>
        <span class="badge ${status}">${STATUS_LABELS[status]}</span>
      </div>
      ${renderResponsePhotoGallery(response, survey)}
      <div class="answer-list">
        ${displayAnswers
          .filter((answer) => {
            const question = survey?.questions.find((item) => item.id === answer.questionId) || {
              type: answer.type,
            };
            return question.type !== "photo";
          })
          .map(
            (answer) => `
              <div class="answer-item">
                <strong>${escapeHtml(answer.label)}</strong><br />
                ${answer.questionId === SESSION_CONCERN_QUESTION_ID
                  ? `
                    <div class="concern-answer-groups">
                      ${getConcernAnswerGroups(answer)
                        .map(
                          (group) => `
                            <section class="concern-answer-group">
                              <strong>${escapeHtml(group.label)}</strong>
                              <div class="checkbox-row">
                                ${group.options
                                  .map((option) => `<label>${escapeHtml(option)}</label>`)
                                  .join("")}
                              </div>
                            </section>
                          `,
                        )
                        .join("")}
                    </div>
                  `
                  : renderAnswerValue(answer)}
              </div>
            `,
          )
          .join("")}
      </div>
    </article>
  `;
}

function renderCustomerTimeline(responseList) {
  const chronological = responseList
    .slice()
    .sort((a, b) => new Date(a.submittedAt) - new Date(b.submittedAt));
  if (!chronological.length) {
    return `<div class="empty">まだカルテに表示できる回答がありません。</div>`;
  }
  return `
    <div class="timeline-list">
      ${chronological
        .map((response) => {
          const survey = findSurveyById(response.surveyId);
          return `
            <article class="timeline-entry">
              <div class="timeline-marker"></div>
              <div class="timeline-content">
                <div class="timeline-head">
                  <div>
                    <strong>${escapeHtml(response.surveyTitle)}</strong>
                    <div class="meta">${formatDate(response.submittedAt)}</div>
                  </div>
                  <span class="badge ${normalizeStatus(response.status)}">${STATUS_LABELS[normalizeStatus(response.status)]}</span>
                </div>
                ${renderTicketStampList(getResponseTicketInfo(response))}
                ${renderResponsePhotoPreview(response, 4)}
                ${renderResponsePhotoGallery(response, survey)}
                <div class="answer-list">
                  ${
                    response.adminMemo
                      ? `
                        <div class="answer-item">
                          <strong>管理メモ</strong><br />
                          ${escapeHtml(response.adminMemo)}
                        </div>
                      `
                      : ""
                  }
                  ${getDisplayAnswers(response, survey)
                    .filter((answer) => {
                      const question = survey?.questions.find((item) => item.id === answer.questionId) || {
                        type: answer.type,
                      };
                      return question.type !== "photo";
                    })
                    .map(
                      (answer) => `
                        <div class="answer-item">
                          <strong>${escapeHtml(answer.label)}</strong><br />
                          ${answer.questionId === SESSION_CONCERN_QUESTION_ID
                            ? `
                              <div class="concern-answer-groups">
                                ${getConcernAnswerGroups(answer)
                                  .map(
                                    (group) => `
                                      <section class="concern-answer-group">
                                        <strong>${escapeHtml(group.label)}</strong>
                                        <div class="checkbox-row">
                                          ${group.options.map((option) => `<label>${escapeHtml(option)}</label>`).join("")}
                                        </div>
                                      </section>
                                    `,
                                  )
                                  .join("")}
                              </div>
                            `
                            : renderAnswerValue(answer)}
                        </div>
                      `,
                    )
                    .join("")}
                </div>
              </div>
            </article>
          `;
        })
        .join("")}
    </div>
  `;
}

function printCustomerTimeline(customerName) {
  const responses = getActiveCustomerResponses(customerName);
  if (!responses.length) {
    showToast("PDF出力できるカルテがありません。");
    return;
  }
  const profile = getCustomerProfileByName(customerName);
  const printWindow = window.open("", "_blank", "noopener,noreferrer");
  if (!printWindow) {
    showToast("PDF出力画面を開けませんでした。");
    return;
  }
  printWindow.document.write(`
    <!doctype html>
    <html lang="ja">
      <head>
        <meta charset="utf-8" />
        <title>${escapeHtml(getCustomerNameWithMember(customerName))}_カルテ</title>
        <style>
          body { font-family: sans-serif; padding: 24px; color: #222; }
          h1 { font-size: 22px; margin-bottom: 8px; }
          .meta { color: #666; margin-bottom: 12px; }
          .entry { border: 1px solid #ddd; border-radius: 8px; padding: 12px; margin-bottom: 12px; }
          .entry h2 { font-size: 16px; margin: 0 0 6px; }
          .item { margin-top: 8px; }
          .item strong { display: block; margin-bottom: 4px; }
        </style>
      </head>
      <body>
        <h1>${escapeHtml(customerName)}</h1>
        <div class="meta">会員番号: ${escapeHtml(profile?.memberNumber || "-")}</div>
        <div class="meta">フリガナ: ${escapeHtml(profile?.nameKana || "-")}</div>
        ${responses
          .slice()
          .sort((a, b) => new Date(a.submittedAt) - new Date(b.submittedAt))
          .map((response) => {
            const survey = findSurveyById(response.surveyId);
            return `
              <section class="entry">
                <h2>${escapeHtml(response.surveyTitle)}</h2>
                <div class="meta">${escapeHtml(formatDate(response.submittedAt))}</div>
                ${response.adminMemo ? `<div class="item"><strong>管理メモ</strong><div>${escapeHtml(response.adminMemo)}</div></div>` : ""}
                ${getDisplayAnswers(response, survey)
                  .map(
                    (answer) => `
                      <div class="item">
                        <strong>${escapeHtml(answer.label)}</strong>
                        <div>${Array.isArray(answer.files) && answer.files.length ? escapeHtml(answer.files.map((file) => file.name || "写真").join(", ")) : escapeHtml(answer.value || "未回答")}</div>
                      </div>
                    `,
                  )
                  .join("")}
              </section>
            `;
          })
          .join("")}
      </body>
    </html>
  `);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
}

function renderCustomerManagement() {
  const stage = document.querySelector("#customerManagementStage");
  if (!stage) return;

  const customers = getFilteredCustomerDirectoryEntries();
  if (state.selectedCustomerViewName && !customers.some((customer) => customer.name === state.selectedCustomerViewName)) {
    state.selectedCustomerViewName = "";
    state.selectedCustomerViewSurveyId = "";
    state.selectedCustomerViewResponseId = "";
    state.selectedCustomerTimelineOpen = false;
  }

  const selectedCustomer =
    customers.find((customer) => customer.name === state.selectedCustomerViewName) || null;
  const customerResponses = selectedCustomer ? getActiveCustomerResponses(selectedCustomer.name) : [];
  const customerMeasurements = selectedCustomer ? getCustomerMeasurements(selectedCustomer.name) : [];
  const surveyGroups = selectedCustomer ? groupResponsesBySurvey(customerResponses) : [];

  if (
    state.selectedMeasurementEditId &&
    !customerMeasurements.some((measurement) => measurement.id === state.selectedMeasurementEditId)
  ) {
    state.selectedMeasurementEditId = "";
  }

  if (
    state.selectedCustomerViewSurveyId &&
    !surveyGroups.some((group) => String(group.surveyId) === state.selectedCustomerViewSurveyId)
  ) {
    state.selectedCustomerViewSurveyId = "";
    state.selectedCustomerViewResponseId = "";
  }

  const selectedSurveyGroup =
    surveyGroups.find((group) => String(group.surveyId) === state.selectedCustomerViewSurveyId) || null;
  const surveyResponses = selectedSurveyGroup
    ? selectedSurveyGroup.responses.slice().sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))
    : [];

  if (
    state.selectedCustomerViewResponseId &&
    !surveyResponses.some((response) => response.id === state.selectedCustomerViewResponseId)
  ) {
    state.selectedCustomerViewResponseId = "";
  }

  const selectedResponse =
    surveyResponses.find((response) => response.id === state.selectedCustomerViewResponseId) || null;

  if (!selectedCustomer) {
    stage.innerHTML = `
      <div class="stage-head">
        <div>
          <div class="card-title">顧客一覧</div>
          <div class="meta">顧客を選ぶと、顧客情報とアンケート回答履歴を表示します。</div>
        </div>
      </div>
      <div class="customer-list">
        ${
          customers.length
            ? customers
                .map(
                  (customer) => `
                    <article class="customer-card selectable-card">
                      <button
                        class="history-open-button"
                        type="button"
                        data-open-customer="${escapeHtml(customer.name)}"
                      >
                        <strong>${escapeHtml(customer.memberNumber ? `${customer.memberNumber} / ${customer.name}` : customer.name)}</strong>
                        <div>回答数: ${customer.count}件</div>
                        ${renderCustomerPushStatus(getCustomerProfileByName(customer.name)?.pushStatus)}
                        <div class="meta">${customer.hasResponses ? `最新回答: ${formatDate(customer.latestAt)}` : "回答履歴はまだありません。"}</div>
                        ${renderTicketStampList(getCurrentTicketInfoForCustomer(customer.name))}
                      </button>
                      <div class="action-row">
                        <button
                          class="secondary-button danger-button"
                          type="button"
                          data-delete-customer-card="${escapeHtml(customer.name)}"
                        >
                          顧客削除
                        </button>
                      </div>
                    </article>
                  `,
                )
                .join("")
            : `<div class="empty">まだ顧客データはありません。</div>`
        }
      </div>
    `;
  } else if (state.selectedCustomerTimelineOpen) {
    stage.innerHTML = `
      <div class="stage-head">
        <div>
          <div class="card-title">時系列カルテ</div>
          <div class="meta">${escapeHtml(getCustomerNameWithMember(selectedCustomer.name))}</div>
        </div>
        <div class="action-row">
          <button class="secondary-button" type="button" data-export-customer-timeline="${escapeHtml(selectedCustomer.name)}">PDF出力</button>
          <button class="secondary-button" type="button" data-back-customer-stage="profile">戻る</button>
        </div>
      </div>
      <div class="stack">
        ${renderCustomerSummaryCard(selectedCustomer.name, customerResponses)}
        <article class="answer-item">
          <strong>時系列カルテ</strong>
          <div class="meta">古い回答から順に表示しています。</div>
          ${renderCustomerTimeline(customerResponses)}
        </article>
      </div>
    `;
  } else if (!selectedSurveyGroup) {
    stage.innerHTML = `
      <div class="stage-head">
        <div>
          <div class="card-title">顧客情報</div>
          <div class="meta">名前・回数券スタンプ情報と、アンケート回答履歴を表示しています。</div>
        </div>
        <div class="action-row">
          <button class="secondary-button danger-button" type="button" data-delete-customer="${escapeHtml(selectedCustomer.name)}">顧客削除</button>
          <button class="secondary-button" type="button" data-back-customer-stage="customers">戻る</button>
        </div>
      </div>
      <div class="stack">
        ${renderCustomerSummaryCard(selectedCustomer.name, customerResponses)}
        ${renderCustomerEditorCard(selectedCustomer.name)}
        ${renderMeasurementSection(selectedCustomer.name)}
        <article class="answer-item">
          <strong>カルテ</strong>
          <div class="meta">顧客ごとの時系列カルテを表示して、PDF出力できます。</div>
          <div class="action-row">
            <button class="secondary-button" type="button" data-open-customer-timeline="${escapeHtml(selectedCustomer.name)}">時系列カルテを見る</button>
            <button class="secondary-button" type="button" data-export-customer-timeline="${escapeHtml(selectedCustomer.name)}">PDF出力</button>
          </div>
        </article>
        <article class="answer-item">
          <strong>アンケート回答履歴</strong>
          <div class="survey-history-list">
            ${
              surveyGroups.length
                ? surveyGroups
                    .map(
                      (group) => `
                        <button
                          class="survey-history-row"
                          type="button"
                          data-open-customer-survey="${escapeHtml(group.surveyId)}"
                        >
                          <strong>${escapeHtml(group.surveyTitle)}</strong>
                          <div class="meta">回答数: ${group.count}件 / 最新: ${formatDate(group.latestAt)}</div>
                        </button>
                      `,
                    )
                    .join("")
                : `<div class="empty">回答履歴はありません。</div>`
            }
          </div>
        </article>
      </div>
    `;
  } else if (!selectedResponse) {
    stage.innerHTML = `
        <div class="stage-head">
        <div>
          <div class="card-title">アンケート回答履歴</div>
          <div class="meta">${escapeHtml(getCustomerNameWithMember(selectedCustomer.name))} / ${escapeHtml(selectedSurveyGroup.surveyTitle)}</div>
        </div>
        <button class="secondary-button" type="button" data-back-customer-stage="surveys">戻る</button>
      </div>
      <div class="stack">
        ${renderCustomerSummaryCard(selectedCustomer.name, customerResponses)}
        <div class="response-list">
          ${
            surveyResponses.length
                ? surveyResponses
                    .map(
                      (response) => `
                      <article class="response-history-card selectable-card">
                        <button
                          class="history-open-button"
                          type="button"
                          data-open-customer-response="${escapeHtml(response.id)}"
                        >
                          <strong>${escapeHtml(formatResponseEntryTitle(response))}</strong>
                          <div class="meta">${formatDate(response.submittedAt)}</div>
                          ${renderTicketStampList(getResponseTicketInfo(response))}
                          ${renderResponsePhotoPreview(response, 3)}
                          <span class="badge ${normalizeStatus(response.status)}">
                            ${STATUS_LABELS[normalizeStatus(response.status)]}
                          </span>
                        </button>
                        <div class="action-row">
                          <button
                            class="secondary-button danger-button"
                            type="button"
                            data-delete-response="${escapeHtml(response.id)}"
                          >
                            ${normalizeStatus(response.status) === "trash" ? "完全削除" : "回答削除"}
                          </button>
                        </div>
                      </article>
                    `,
                  )
                  .join("")
              : `<div class="empty">このアンケートの回答履歴はありません。</div>`
          }
        </div>
      </div>
    `;
  } else {
    stage.innerHTML = `
      <div class="stage-head">
        <div>
          <div class="card-title">アンケート回答詳細</div>
          <div class="meta">${escapeHtml(getCustomerNameWithMember(selectedCustomer.name))} / ${escapeHtml(selectedSurveyGroup.surveyTitle)} / ${formatDate(selectedResponse.submittedAt)}</div>
        </div>
        <div class="action-row">
          <button class="secondary-button" type="button" data-print-customer-response="${selectedResponse.id}">印刷</button>
          <button class="secondary-button" type="button" data-download-customer-photos="${selectedResponse.id}">写真保存</button>
          <button class="secondary-button danger-button" type="button" data-delete-response="${selectedResponse.id}">${normalizeStatus(selectedResponse.status) === "trash" ? "完全削除" : "回答削除"}</button>
          <button class="secondary-button" type="button" data-back-customer-stage="responses">戻る</button>
          <button class="secondary-button" type="button" data-close-customer-detail>閉じる</button>
        </div>
      </div>
      <div class="stack">
        ${renderCustomerSummaryCard(selectedCustomer.name, customerResponses)}
        ${renderComparisonSection(selectedResponse)}
        ${renderResponseCard(selectedResponse)}
      </div>
    `;
  }

  stage.querySelectorAll("[data-open-customer]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedCustomerViewName = button.dataset.openCustomer || "";
      state.selectedCustomerViewSurveyId = "";
      state.selectedCustomerViewResponseId = "";
      state.selectedCustomerTimelineOpen = false;
      state.selectedMeasurementEditId = "";
      renderCustomerManagement();
    });
  });

  stage.querySelectorAll("[data-open-customer-timeline]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedCustomerTimelineOpen = true;
      state.selectedCustomerViewSurveyId = "";
      state.selectedCustomerViewResponseId = "";
      renderCustomerManagement();
    });
  });

  stage.querySelectorAll("[data-open-customer-survey]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedCustomerViewSurveyId = button.dataset.openCustomerSurvey || "";
      state.selectedCustomerViewResponseId = "";
      state.selectedCustomerTimelineOpen = false;
      renderCustomerManagement();
    });
  });

  stage.querySelectorAll("[data-open-customer-response]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedCustomerViewResponseId = button.dataset.openCustomerResponse || "";
      state.selectedCustomerTimelineOpen = false;
      renderCustomerManagement();
    });
  });

  stage.querySelectorAll("[data-back-customer-stage]").forEach((button) => {
    button.addEventListener("click", () => {
      const step = button.dataset.backCustomerStage;
      if (step === "customers") {
        state.selectedCustomerViewName = "";
        state.selectedCustomerViewSurveyId = "";
        state.selectedCustomerViewResponseId = "";
        state.selectedCustomerTimelineOpen = false;
        state.selectedMeasurementEditId = "";
      } else if (step === "profile") {
        state.selectedCustomerTimelineOpen = false;
        state.selectedCustomerViewSurveyId = "";
        state.selectedCustomerViewResponseId = "";
      } else if (step === "surveys") {
        state.selectedCustomerViewSurveyId = "";
        state.selectedCustomerViewResponseId = "";
      } else {
        state.selectedCustomerViewResponseId = "";
      }
      renderCustomerManagement();
    });
  });

  stage.querySelectorAll("[data-print-customer-response]").forEach((button) => {
    button.addEventListener("click", () => {
      const response = state.responses.find((item) => item.id === button.dataset.printCustomerResponse);
      if (response) printResponse(response);
    });
  });

  stage.querySelectorAll("[data-delete-customer], [data-delete-customer-card]").forEach((button) => {
    button.addEventListener("click", () => {
      void deleteCustomerProfile(button.dataset.deleteCustomer || button.dataset.deleteCustomerCard || "");
    });
  });

  stage.querySelector("#measurementTargetForm")?.addEventListener("submit", (event) => {
    event.preventDefault();
    void saveCustomerMeasurementTargets(event.currentTarget);
  });

  stage.querySelector("#measurementForm")?.addEventListener("submit", (event) => {
    event.preventDefault();
    void saveMeasurement(event.currentTarget);
  });

  stage.querySelector("[data-cancel-measurement-edit]")?.addEventListener("click", () => {
    state.selectedMeasurementEditId = "";
    renderCustomerManagement();
  });

  stage.querySelectorAll("[data-edit-measurement]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedMeasurementEditId = button.dataset.editMeasurement || "";
      renderCustomerManagement();
    });
  });

  stage.querySelectorAll("[data-delete-measurement]").forEach((button) => {
    button.addEventListener("click", () => {
      void deleteMeasurement(button.dataset.deleteMeasurement || "");
    });
  });

  stage.querySelector("#measurementPeriodFilter")?.addEventListener("change", (event) => {
    state.selectedMeasurementPeriod = String(event.currentTarget.value || "all");
    renderCustomerManagement();
  });

  stage.querySelectorAll("[data-toggle-measurement-metric]").forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      const metricKey = checkbox.dataset.toggleMeasurementMetric;
      if (!metricKey) return;
      state.measurementMetricVisibility = {
        ...state.measurementMetricVisibility,
        [metricKey]: checkbox.checked,
      };
      renderCustomerManagement();
    });
  });

  stage.querySelectorAll("[data-download-customer-photos]").forEach((button) => {
    button.addEventListener("click", () => {
      const response = state.responses.find((item) => item.id === button.dataset.downloadCustomerPhotos);
      if (response) {
        void downloadFiles(collectPhotosFromResponses([response]), response.customerName || "response");
      }
    });
  });

  stage.querySelector("[data-close-customer-detail]")?.addEventListener("click", () => {
    state.selectedCustomerViewSurveyId = "";
    state.selectedCustomerViewResponseId = "";
    state.selectedCustomerTimelineOpen = false;
    renderCustomerManagement();
  });

  stage.querySelectorAll("[data-export-customer-timeline]").forEach((button) => {
    button.addEventListener("click", () => {
      printCustomerTimeline(button.dataset.exportCustomerTimeline || "");
    });
  });

  stage.querySelector("#customerProfileForm")?.addEventListener("submit", (event) => {
    event.preventDefault();
    void saveCustomerProfile(event.currentTarget);
  });

  stage.querySelector("#deleteCustomerButton")?.addEventListener("click", () => {
    void deleteCustomerProfile(selectedCustomer?.name || "");
  });

  stage.querySelectorAll("[data-save-response]").forEach((button) => {
    button.addEventListener("click", () => {
      void saveResponseManagement(
        button.dataset.saveResponse,
        button.closest("[data-response-card]"),
      );
    });
  });

  stage.querySelectorAll("[data-delete-response]").forEach((button) => {
    button.addEventListener("click", () => {
      void deleteResponse(button.dataset.deleteResponse);
    });
  });

  attachLightboxHandlers(stage);
}

function renderCurrentCustomerInfo() {
  const container = document.querySelector("#respondentInfoSection");
  if (!container) return;
  const response = state.responses.find((item) => item.id === state.selectedResponseId);
  if (!response) {
    container.innerHTML = `<div class="empty">回答詳細を開くと、回答者の現在の回数券情報を表示します。</div>`;
    return;
  }
  const customerResponses = state.responses.filter(
    (item) => item.customerName === response.customerName && normalizeStatus(item.status) !== "trash",
  );
  const ticketInfo = getCurrentTicketInfoForCustomer(response.customerName);
  const memoRecord = getCustomerMemoRecord(response.customerName);
  container.innerHTML = `
    <article class="answer-item">
      <strong>${escapeHtml(getCustomerNameWithMember(response.customerName))}</strong>
      <div class="meta">フリガナ: ${escapeHtml(getCustomerProfileByName(response.customerName)?.nameKana || "-")}</div>
      <div class="meta">最新表示中の回答: ${escapeHtml(response.surveyTitle)} / ${formatDate(response.submittedAt)}</div>
      <div class="meta">回答数: ${customerResponses.length}件</div>
      ${
        ticketInfo.length
          ? `
            ${renderTicketStampPanel(ticketInfo)}
          `
          : `<div class="meta">現在の回数券情報はありません。</div>`
      }
    </article>
    <article class="answer-item">
      <strong>お客様メモ</strong>
      <textarea id="customerMemoInput" placeholder="自由にメモを残せます。">${escapeHtml(memoRecord.latestMemo)}</textarea>
      <div class="action-row">
        <button class="secondary-button" type="button" id="saveCustomerMemoButton">メモ保存</button>
        <button class="secondary-button" type="button" id="downloadCustomerPhotosButton">写真一括保存</button>
      </div>
    </article>
    <article class="answer-item">
      <strong>メモ履歴</strong>
      ${renderMemoTimeline(response.customerName)}
    </article>
  `;
  container.querySelector("#saveCustomerMemoButton")?.addEventListener("click", () => {
    void saveCustomerMemo(response.customerName);
  });
  container.querySelector("#downloadCustomerPhotosButton")?.addEventListener("click", () => {
    void downloadFiles(
      collectPhotosFromResponses(customerResponses),
      response.customerName || "customer",
    );
  });
}

function renderAllResponsesBySurveySection() {
  const container = document.querySelector("#allResponsesBySurveySection");
  if (!container) return;
  const groups = groupResponsesBySurvey(
    state.responses.filter((response) => normalizeStatus(response.status) !== "trash"),
  );
  if (!groups.length) {
    container.innerHTML = `<div class="empty">まだ回答履歴はありません。</div>`;
    return;
  }
  container.innerHTML = groups
    .map(
      (group) => `
        <article class="answer-item">
          <strong>${escapeHtml(group.surveyTitle)}</strong>
          <div class="meta">回答数: ${group.count}件 / 最新: ${formatDate(group.latestAt)}</div>
          <div class="survey-history-list">
            ${group.responses
              .slice()
              .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))
              .map(
                (response) => `
                  <button
                    class="survey-history-row"
                    type="button"
                    data-survey-history-open="${escapeHtml(response.id)}"
                    data-survey-history-survey="${escapeHtml(group.surveyId)}"
                  >
                    <strong>${escapeHtml(getCustomerNameWithMember(response.customerName))}</strong>
                    <div class="meta">${formatDate(response.submittedAt)}</div>
                    ${renderTicketStampList(getResponseTicketInfo(response))}
                    ${renderResponsePhotoPreview(response, 2)}
                  </button>
                `,
              )
              .join("")}
          </div>
        </article>
      `,
    )
    .join("");

  container.querySelectorAll("[data-survey-history-open]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedResponseSurveyId = button.dataset.surveyHistorySurvey || "";
      state.selectedResponseId = button.dataset.surveyHistoryOpen || "";
      setPage("responses");
      renderResponses();
    });
  });
}

function findSurveyById(surveyId) {
  return state.surveys.find((survey) => survey.id === surveyId);
}

function getDisplayAnswers(response, survey) {
  const answers = Array.isArray(response?.answers) ? response.answers : [];
  if (!survey) return answers;

  const answerMap = new Map(answers.map((answer) => [answer.questionId, answer]));
  const orderedAnswers = survey.questions.map((question) => {
    const existing = answerMap.get(question.id);
    if (existing) {
      return {
        ...existing,
        label: existing.label || question.label,
        type: existing.type || question.type,
      };
    }
    return {
      questionId: question.id,
      label: question.label,
      type: question.type,
      value: "",
      files: [],
    };
  });

  const extraAnswers = answers.filter(
    (answer) => !survey.questions.some((question) => question.id === answer.questionId),
  );

  return orderedAnswers.concat(extraAnswers);
}

function getAnswerValues(answer) {
  return String(answer?.value || "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function getConcernAnswerGroups(answer) {
  const selected = new Set(getAnswerValues(answer));
  const groups = SESSION_CONCERN_CATEGORIES.map((category) => ({
    id: category.id,
    label: category.label,
    options: category.options.filter((option) => selected.has(option)),
  })).filter((category) => category.options.length);
  const knownOptions = new Set(SESSION_CONCERN_CATEGORIES.flatMap((category) => category.options));
  const extras = Array.from(selected).filter((option) => !knownOptions.has(option));
  if (extras.length) {
    groups.push({
      id: "extras",
      label: "【その他】",
      options: extras,
    });
  }
  return groups;
}

function renderConcernAnswerEditor(answer, questionId) {
  const selected = new Set(getAnswerValues(answer));
  return `
    <div class="concern-answer-groups">
      ${SESSION_CONCERN_CATEGORIES.map(
        (category) => `
          <section class="concern-answer-group">
            <strong>${escapeHtml(category.label)}</strong>
            <div class="checkbox-row">
              ${category.options
                .map(
                  (option) => `
                    <label>
                      <input
                        type="checkbox"
                        data-answer-checkbox="${questionId}"
                        value="${escapeHtml(option)}"
                        ${selected.has(option) ? "checked" : ""}
                      />
                      <span class="option-text">${escapeHtml(option)}</span>
                    </label>
                  `,
                )
                .join("")}
            </div>
          </section>
        `,
      ).join("")}
      <section class="concern-answer-group">
        <strong>【その他】</strong>
        <div class="checkbox-row">
          <label>
            <input
              type="checkbox"
              data-answer-checkbox="${questionId}"
              value="その他（長文）"
              ${selected.has("その他（長文）") ? "checked" : ""}
            />
            <span class="option-text">その他（長文）</span>
          </label>
        </div>
      </section>
    </div>
  `;
}

function collectPhotosFromResponse(response) {
  const answerFiles = (response?.answers || []).flatMap((answer) => getPhotoFilesFromAnswer(answer));
  const responseFiles = (Array.isArray(response?.files) ? response.files : [])
    .map((file, index) => {
      if (file?.previewUrl || file?.thumbnailUrl || file?.dataUrl || file?.url) {
        return ensureDownloadablePhotoFile(file, index);
      }
      return derivePhotoFileFromUrl(file?.url || file?.value || "", index);
    })
    .filter(Boolean);
  const files = answerFiles.concat(responseFiles);
  const seen = new Set();
  return files.filter((file) => {
    const key = String(file?.fileId || file?.previewUrl || file?.url || file?.name || "");
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function renderResponsePhotoGallery(response, survey) {
  const photoAnswers = getDisplayAnswers(response, survey).filter(
    (answer) => Array.isArray(answer.files) && answer.files.length,
  );
  const photos = collectPhotosFromResponse(response);
  if (!photos.length) return "";

  return `
    <article class="answer-item response-photo-panel">
      <strong>アップロード写真</strong>
      <div class="meta">回答で添付された写真を表示しています。</div>
      ${
        photoAnswers.length
          ? photoAnswers
              .map(
                (answer) => `
                  <div class="photo-answer-group">
                    <strong>${escapeHtml(answer.label)}</strong>
                    ${renderAnswerValue(answer)}
                  </div>
                `,
              )
              .join("")
          : `
              <div class="photo-list">
                ${photos
                  .map((file) => {
                    const preview = getPhotoPreviewSrc(file);
                    const href = getPhotoLightboxSrc(file) || "#";
                    return `
                      <button
                        class="photo-thumb lightbox-trigger"
                        type="button"
                        data-lightbox-src="${escapeHtml(href)}"
                        data-lightbox-title="${escapeHtml(file.name || "写真")}"
                      >
                        ${preview ? `<img src="${escapeHtml(preview)}" alt="${escapeHtml(file.name || "写真")}" />` : ""}
                        <span>${escapeHtml(file.name || "写真")}</span>
                      </button>
                    `;
                  })
                  .join("")}
              </div>
            `
      }
    </article>
  `;
}

function renderEditableAnswerField(question, answer) {
  if (question.type === "photo") {
    return `
      ${renderAnswerValue(answer)}
      <div class="meta">写真は管理アプリでは表示のみです。</div>
    `;
  }

  if (question.type === "textarea") {
    return `<textarea data-answer-input="${question.id}">${escapeHtml(answer.value || "")}</textarea>`;
  }

  if (question.type === "rating") {
    return `
      <select data-answer-input="${question.id}">
        <option value="">選択してください</option>
        ${["1", "2", "3", "4", "5"]
          .map(
            (value) => `
              <option value="${value}" ${String(answer.value || "") === value ? "selected" : ""}>
                ${value}
              </option>
            `,
          )
          .join("")}
      </select>
    `;
  }

  if (question.type === "choice") {
    return `
      <select data-answer-input="${question.id}">
        <option value="">選択してください</option>
        ${question.options
          .map(
            (option) => `
              <option value="${escapeHtml(option)}" ${String(answer.value || "") === option ? "selected" : ""}>
                ${escapeHtml(option)}
              </option>
            `,
          )
          .join("")}
      </select>
    `;
  }

  if (question.type === "checkbox") {
    if (question.id === SESSION_CONCERN_QUESTION_ID) {
      return renderConcernAnswerEditor(answer, question.id);
    }
    const selected = new Set(getAnswerValues(answer));
    return `
      <div class="checkbox-row">
        ${question.options
          .map(
            (option) => `
              <label>
                <input
                  type="checkbox"
                  data-answer-checkbox="${question.id}"
                  value="${escapeHtml(option)}"
                  ${selected.has(option) ? "checked" : ""}
                />
                ${escapeHtml(option)}
              </label>
            `,
          )
          .join("")}
      </div>
    `;
  }

  return `<input type="text" data-answer-input="${question.id}" value="${escapeHtml(answer.value || "")}" />`;
}

function formatAnswerForCsv(answer) {
  const photoFiles = getPhotoFilesFromAnswer(answer);
  if (photoFiles.length) {
    return `${answer.label}: ${photoFiles
      .map((file) => file.url || file.previewUrl || file.name)
      .join(", ")}`;
  }
  if (answer.questionId === SESSION_CONCERN_QUESTION_ID) {
    return `${answer.label}: ${getConcernAnswerGroups(answer)
      .map((group) => `${group.label} ${group.options.join(" / ")}`)
      .join(" | ")}`;
  }
  return `${answer.label}: ${answer.value || ""}`;
}

function renderFilters() {
  const surveyFilter = document.querySelector("#surveyFilter");
  const concernCategoryFilter = document.querySelector("#concernCategoryFilter");
  const ticketPlanFilter = document.querySelector("#ticketPlanFilter");
  const ticketSheetFilter = document.querySelector("#ticketSheetFilter");
  const ticketRoundFilter = document.querySelector("#ticketRoundFilter");
  const current = surveyFilter.value;
  surveyFilter.innerHTML = `
    <option value="">すべて</option>
    ${state.surveys.map((survey) => `<option value="${survey.id}">${escapeHtml(survey.title)}</option>`).join("")}
  `;
  surveyFilter.value = Array.from(surveyFilter.options).some((option) => option.value === current) ? current : "";
  if (concernCategoryFilter) {
    const currentCategory = concernCategoryFilter.value;
    concernCategoryFilter.innerHTML = `
      <option value="">すべてのカテゴリ</option>
      ${SESSION_CONCERN_CATEGORIES.map(
        (category) => `<option value="${escapeHtml(category.id)}">${escapeHtml(category.label)}</option>`,
      ).join("")}
      <option value="extras">【その他】</option>
    `;
    concernCategoryFilter.value = Array.from(concernCategoryFilter.options).some(
      (option) => option.value === currentCategory,
    )
      ? currentCategory
      : "";
  }
  if (ticketPlanFilter) {
    const currentPlan = ticketPlanFilter.value;
    ticketPlanFilter.innerHTML = `
      <option value="">すべての回数券</option>
      ${CUSTOMER_TICKET_PLAN_OPTIONS.filter(Boolean)
        .map((option) => `<option value="${escapeHtml(option)}">${escapeHtml(option)}</option>`)
        .join("")}
    `;
    ticketPlanFilter.value = Array.from(ticketPlanFilter.options).some((option) => option.value === currentPlan)
      ? currentPlan
      : "";
  }
  if (ticketSheetFilter) {
    const currentSheet = ticketSheetFilter.value;
    ticketSheetFilter.innerHTML = `
      <option value="">すべての枚数</option>
      ${CUSTOMER_TICKET_SHEET_OPTIONS.map((option) => `<option value="${escapeHtml(option)}">${escapeHtml(option)}</option>`).join("")}
    `;
    ticketSheetFilter.value = Array.from(ticketSheetFilter.options).some((option) => option.value === currentSheet)
      ? currentSheet
      : "";
  }
  if (ticketRoundFilter) {
    const currentRound = ticketRoundFilter.value;
    ticketRoundFilter.innerHTML = `
      <option value="">すべての回数</option>
      ${CUSTOMER_TICKET_ROUND_OPTIONS.map((option) => `<option value="${escapeHtml(option)}">${escapeHtml(option)}</option>`).join("")}
    `;
    ticketRoundFilter.value = Array.from(ticketRoundFilter.options).some((option) => option.value === currentRound)
      ? currentRound
      : "";
  }
}

function getCustomerMemo(customerName) {
  return getCustomerMemoRecord(customerName).latestMemo;
}

async function saveCustomerMemo(customerName) {
  const textarea = document.querySelector("#customerMemoInput");
  if (!textarea || !customerName) return;
  const memo = textarea.value.trim();
  try {
    const result = await api.request(`/api/admin/customer-memos/${encodeURIComponent(customerName)}`, {
      method: "PUT",
      token: state.token,
      body: { memo },
    });
    state.customerMemos = result.memos || state.customerMemos;
    showToast("お客様メモを保存しました。");
    renderResponses();
  } catch (error) {
    showToast(error.message || "お客様メモを保存できませんでした。");
  }
}

async function applyBulkStatus() {
  const nextStatus = document.querySelector("#bulkStatusSelect")?.value;
  if (!nextStatus || !state.selectedResponseIds.length) {
    showToast("一括変更する回答を選択してください。");
    return;
  }
  try {
    const results = await Promise.all(
      state.selectedResponseIds.map((responseId) =>
        api.request(`/api/admin/responses/${encodeURIComponent(responseId)}`, {
          method: "PUT",
          token: state.token,
          body: {
            status: nextStatus,
            adminMemo: state.responses.find((response) => response.id === responseId)?.adminMemo || "",
            answers: state.responses.find((response) => response.id === responseId)?.answers || [],
          },
        }),
      ),
    );
    results.forEach((result) => {
      state.responses = state.responses.map((response) =>
        response.id === result.response.id ? result.response : response,
      );
    });
    state.selectedResponseIds = [];
    showToast("対応状況を一括更新しました。");
    renderAll();
  } catch (error) {
    showToast(error.message || "一括更新できませんでした。");
  }
}

function getResponseSearchHaystack(response) {
  const profile = getCustomerProfileByName(response.customerName);
  return [
    response.customerName,
    response.customerMemberNumber,
    profile?.memberNumber,
    profile?.nameKana,
    response.surveyTitle,
    response.adminMemo,
    getResponseTicketInfo(response).map((item) => `${item.label} ${item.value}`).join(" "),
    getResponseConcernCategoryLabels(response).join(" "),
    response.answers.map((answer) => formatAnswerForCsv(answer)).join(" "),
  ]
    .filter(Boolean)
    .join(" ")
    .toLowerCase();
}

function normalizeLooseSearchText(value) {
  return String(value || "").replace(/\s+/g, "").trim().toLowerCase();
}

function getResponseConcernCategoryIds(response) {
  const answer = (response?.answers || []).find((item) => item.questionId === SESSION_CONCERN_QUESTION_ID);
  return getConcernAnswerGroups(answer).map((group) => group.id);
}

function getResponseConcernCategoryLabels(response) {
  return getResponseConcernCategoryIds(response)
    .map((categoryId) => {
      if (categoryId === "extras") return "【その他】";
      return SESSION_CONCERN_CATEGORIES.find((category) => category.id === categoryId)?.label || "";
    })
    .filter(Boolean);
}

function getPhotoCountBucket(count) {
  if (count >= 5) return "5plus";
  if (count >= 3) return "3-4";
  if (count >= 1) return "1-2";
  return "0";
}

function getFilteredResponses() {
  const surveyId = document.querySelector("#surveyFilter").value;
  const status = document.querySelector("#statusFilter").value;
  const keyword = document.querySelector("#keywordFilter").value.trim().toLowerCase();
  const memberNumber = String(document.querySelector("#memberNumberFilter")?.value || "").trim().toUpperCase();
  const nameKana = normalizeLooseSearchText(document.querySelector("#nameKanaFilter")?.value || "");
  const memoFilter = document.querySelector("#memoPresenceFilter")?.value || "";
  const concernCategory = document.querySelector("#concernCategoryFilter")?.value || "";
  const ticketPlan = document.querySelector("#ticketPlanFilter")?.value || "";
  const ticketSheet = document.querySelector("#ticketSheetFilter")?.value || "";
  const ticketRound = document.querySelector("#ticketRoundFilter")?.value || "";
  const photoFilter = document.querySelector("#photoFilter")?.value || "";
  const photoCountFilter = document.querySelector("#photoCountFilter")?.value || "";
  const dateFrom = document.querySelector("#dateFromFilter")?.value || "";
  const dateTo = document.querySelector("#dateToFilter")?.value || "";

  return state.responses
    .filter((response) => !surveyId || response.surveyId === surveyId)
    .filter((response) =>
      status
        ? normalizeStatus(response.status) === status
        : normalizeStatus(response.status) !== "trash",
    )
    .filter((response) => {
      const submittedAt = new Date(response.submittedAt);
      if (dateFrom && submittedAt < new Date(`${dateFrom}T00:00:00`)) return false;
      if (dateTo && submittedAt > new Date(`${dateTo}T23:59:59`)) return false;
      return true;
    })
    .filter((response) => {
      if (!keyword) return true;
      return getResponseSearchHaystack(response).includes(keyword);
    })
    .filter((response) => {
      const profile = getCustomerProfileByName(response.customerName);
      const normalizedMemberNumber = String(
        response.customerMemberNumber || profile?.memberNumber || "",
      ).trim().toUpperCase();
      const normalizedNameKana = normalizeLooseSearchText(profile?.nameKana || "");
      const responseMemo = String(response.adminMemo || "").trim();
      const concernCategoryIds = getResponseConcernCategoryIds(response);
      const photoCount = collectPhotosFromResponse(response).length;
      const ticketMap = new Map(getResponseTicketInfo(response).map((item) => [item.label, item.value]));
      if (memberNumber && !normalizedMemberNumber.includes(memberNumber)) return false;
      if (nameKana && !normalizedNameKana.includes(nameKana)) return false;
      if (memoFilter === "with" && !responseMemo) return false;
      if (memoFilter === "without" && responseMemo) return false;
      if (concernCategory && !concernCategoryIds.includes(concernCategory)) return false;
      if (ticketPlan && ticketMap.get("回数券") !== ticketPlan) return false;
      if (ticketSheet && ticketMap.get("何枚目") !== ticketSheet) return false;
      if (ticketRound && ticketMap.get("何回目") !== ticketRound) return false;
      if (photoFilter === "with" && !photoCount) return false;
      if (photoFilter === "without" && photoCount) return false;
      if (photoCountFilter && getPhotoCountBucket(photoCount) !== photoCountFilter) return false;
      return true;
    })
    .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt));
}

function getCustomerResponses(customerName) {
  return getFilteredResponses().filter((response) => response.customerName === customerName);
}

function getPreviousResponse(response) {
  return state.responses
    .filter(
      (item) =>
        item.customerName === response.customerName &&
        item.surveyId === response.surveyId &&
        item.id !== response.id &&
        normalizeStatus(item.status) !== "trash" &&
        new Date(item.submittedAt) < new Date(response.submittedAt),
    )
    .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))[0] || null;
}

function hasResponsePhotos(response) {
  return collectPhotosFromResponse(response).length > 0;
}

function getPhotoComparisonSlots(response) {
  const photoResponses = state.responses
    .filter(
      (item) =>
        item.customerName === response.customerName &&
        item.surveyId === response.surveyId &&
        normalizeStatus(item.status) !== "trash" &&
        hasResponsePhotos(item),
    )
    .sort((a, b) => new Date(a.submittedAt) - new Date(b.submittedAt));

  if (!photoResponses.length) return [];

  const first = photoResponses[0] || null;
  const previous =
    photoResponses.filter((item) => new Date(item.submittedAt) < new Date(response.submittedAt)).slice(-1)[0] || null;
  const currentHasPhotos = hasResponsePhotos(response);
  const slots = [];

  if (first) {
    slots.push({ label: "初回", response: first });
  }
  if (previous && previous.id !== first?.id) {
    slots.push({ label: "前回", response: previous });
  }
  if (currentHasPhotos) {
    const existingSlot = slots.find((slot) => slot.response?.id === response.id);
    if (existingSlot) {
      existingSlot.label = `${existingSlot.label} / 今回`;
    } else {
      slots.push({ label: "今回", response });
    }
  } else {
    slots.push({ label: "今回", response: null });
  }

  return slots;
}

function renderPhotoComparisonSection(response) {
  const survey = findSurveyById(response.surveyId);
  const slots = getPhotoComparisonSlots(response);
  if (!survey || !slots.length) {
    return `<article class="answer-item"><strong>写真比較</strong><div class="meta">比較できる写真はありません。</div></article>`;
  }

  return `
    <article class="answer-item photo-compare-section">
      <strong>写真比較</strong>
      <div class="meta">初回・前回・今回の写真を並べて確認できます。</div>
      <div class="photo-compare-grid">
        ${slots
          .map((slot) => {
            if (!slot.response) {
              return `
                <section class="photo-compare-card">
                  <div class="photo-compare-head">
                    <strong>${escapeHtml(slot.label)}</strong>
                    <span class="badge checked">写真なし</span>
                  </div>
                  <div class="empty">この回答には写真が添付されていません。</div>
                </section>
              `;
            }

            const photoAnswers = getDisplayAnswers(slot.response, survey).filter(
              (answer) => Array.isArray(answer.files) && answer.files.length,
            );
            return `
              <section class="photo-compare-card">
                <div class="photo-compare-head">
                  <strong>${escapeHtml(slot.label)}</strong>
                  <span class="meta">${escapeHtml(formatDate(slot.response.submittedAt))}</span>
                </div>
                <div class="meta">${escapeHtml(getCustomerNameWithMember(slot.response.customerName))}</div>
                ${renderTicketStampList(getResponseTicketInfo(slot.response))}
                <div class="photo-compare-gallery">
                  ${
                    photoAnswers.length
                      ? photoAnswers
                          .map(
                            (answer) => `
                              <div class="photo-answer-group">
                                <strong>${escapeHtml(answer.label)}</strong>
                                ${renderAnswerValue(answer)}
                              </div>
                            `,
                          )
                          .join("")
                      : `<div class="empty">比較できる写真はありません。</div>`
                  }
                </div>
              </section>
            `;
          })
          .join("")}
      </div>
    </article>
  `;
}

function renderComparisonSection(response) {
  const survey = findSurveyById(response.surveyId);
  const previous = getPreviousResponse(response);
  if (!survey || !previous) {
    return `<article class="answer-item"><strong>前回比較</strong><div class="meta">比較できる前回回答はありません。</div></article>`;
  }
  const currentMap = new Map(getDisplayAnswers(response, survey).map((answer) => [answer.questionId, answer]));
  const previousMap = new Map(getDisplayAnswers(previous, survey).map((answer) => [answer.questionId, answer]));
  const rows = survey.questions
    .map((question) => {
      const current = currentMap.get(question.id) || { value: "", files: [] };
      const last = previousMap.get(question.id) || { value: "", files: [] };
      const currentValue = Array.isArray(current.files) && current.files.length
        ? `${current.files.length}枚`
        : String(current.value || "未回答");
      const lastValue = Array.isArray(last.files) && last.files.length
        ? `${last.files.length}枚`
        : String(last.value || "未回答");
      if (currentValue === lastValue) return "";
      return `
        <div class="comparison-row">
          <strong>${escapeHtml(question.label)}</strong>
          <div class="meta">前回: ${escapeHtml(lastValue)}</div>
          <div>今回: ${escapeHtml(currentValue)}</div>
        </div>
      `;
    })
    .filter(Boolean);
  return `
    <article class="answer-item">
      <strong>前回比較</strong>
      <div class="meta">前回回答: ${escapeHtml(formatDate(previous.submittedAt))}</div>
      ${rows.length ? rows.join("") : `<div class="meta">前回と差分はありません。</div>`}
    </article>
  `;
}

function renderMemoTimeline(customerName) {
  const record = getCustomerMemoRecord(customerName);
  if (!record.entries.length) {
    return `<div class="empty">メモ履歴はまだありません。</div>`;
  }
  return `
    <div class="memo-timeline">
      ${record.entries
        .map(
          (entry) => `
            <article class="timeline-item">
              <div class="meta">${escapeHtml(entry.at ? formatDate(entry.at) : "-")}</div>
              <div>${escapeHtml(entry.memo || "")}</div>
            </article>
          `,
        )
        .join("")}
    </div>
  `;
}

function collectPhotosFromResponses(responses) {
  return responses.flatMap((response) => collectPhotosFromResponse(response));
}

function sanitizeDownloadNamePart(value, fallback = "") {
  const normalized = String(value || "")
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim();
  return normalized || fallback;
}

function getPhotoDownloadExtension(file) {
  const name = String(file?.name || "").trim();
  const matched = name.match(/\.[a-zA-Z0-9]{2,5}$/);
  if (matched) return matched[0];
  const type = String(file?.type || "").toLowerCase();
  if (type.includes("png")) return ".png";
  if (type.includes("webp")) return ".webp";
  if (type.includes("gif")) return ".gif";
  return ".jpg";
}

function buildPhotoDownloadName(file, prefix, index) {
  const safePrefix = sanitizeDownloadNamePart(prefix, "photo");
  const safeName = sanitizeDownloadNamePart(file?.name || "", "");
  if (safeName) {
    return `${safePrefix}-${String(index + 1).padStart(2, "0")}-${safeName}`;
  }
  return `${safePrefix}-${String(index + 1).padStart(2, "0")}${getPhotoDownloadExtension(file)}`;
}

function triggerDownloadLink(href, fileName) {
  const link = document.createElement("a");
  link.href = href;
  link.rel = "noopener";
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function downloadSingleFile(file, prefix, index) {
  const normalizedFile = ensureDownloadablePhotoFile(file, index);
  if (!normalizedFile) return false;
  const fileName = buildPhotoDownloadName(normalizedFile, prefix, index);
  const sourceCandidates = [
    normalizedFile.downloadUrl,
    normalizedFile.url,
    normalizedFile.previewUrl,
    normalizedFile.thumbnailUrl,
    normalizedFile.dataUrl,
  ].filter(Boolean);

  for (const href of sourceCandidates) {
    if (href.startsWith("data:")) {
      triggerDownloadLink(href, fileName);
      return true;
    }
    try {
      const response = await fetch(href, {
        method: "GET",
        mode: "cors",
        credentials: "omit",
        cache: "no-store",
      });
      if (!response.ok) continue;
      const blob = await response.blob();
      if (!blob.size) continue;
      const objectUrl = URL.createObjectURL(blob);
      try {
        triggerDownloadLink(objectUrl, fileName);
      } finally {
        setTimeout(() => URL.revokeObjectURL(objectUrl), 30000);
      }
      return true;
    } catch {
      // Try the next candidate and fall back to a direct browser download.
    }
  }

  const fallbackHref = sourceCandidates[0];
  if (!fallbackHref) return false;
  triggerDownloadLink(fallbackHref, fileName);
  return true;
}

async function downloadFiles(files, prefix) {
  const normalizedFiles = (Array.isArray(files) ? files : [])
    .map((file, index) => ensureDownloadablePhotoFile(file, index))
    .filter(Boolean);
  if (!normalizedFiles.length) {
    showToast("ダウンロードできる写真がありません。");
    return;
  }
  let successCount = 0;
  for (let index = 0; index < normalizedFiles.length; index += 1) {
    if (await downloadSingleFile(normalizedFiles[index], prefix, index)) {
      successCount += 1;
    }
    await wait(140);
  }
  if (!successCount) {
    showToast("写真を保存できませんでした。");
    return;
  }
  showToast(`${successCount}件の写真保存を開始しました。`);
}

function printResponse(response) {
  const survey = findSurveyById(response.surveyId);
  const answers = getDisplayAnswers(response, survey);
  const customerLabel = getCustomerNameWithMember(response.customerName);
  const printWindow = window.open("", "_blank", "noopener,noreferrer");
  if (!printWindow) {
    showToast("印刷画面を開けませんでした。");
    return;
  }
  printWindow.document.write(`
    <!doctype html>
    <html lang="ja">
      <head>
        <meta charset="utf-8" />
        <title>${escapeHtml(customerLabel)}_${escapeHtml(response.surveyTitle)}</title>
        <style>
          body { font-family: sans-serif; padding: 24px; color: #222; }
          h1 { font-size: 20px; margin-bottom: 8px; }
          .meta { color: #666; margin-bottom: 16px; }
          .item { border: 1px solid #ddd; border-radius: 6px; padding: 12px; margin-bottom: 10px; }
          .item strong { display:block; margin-bottom:6px; }
        </style>
      </head>
      <body>
        <h1>${escapeHtml(response.surveyTitle)}</h1>
        <div class="meta">${escapeHtml(customerLabel)} / ${escapeHtml(formatDate(response.submittedAt))}</div>
        ${answers
          .map(
            (answer) => `
              <div class="item">
                <strong>${escapeHtml(answer.label)}</strong>
                <div>${Array.isArray(answer.files) && answer.files.length ? escapeHtml(answer.files.map((file) => file.name || "写真").join(", ")) : escapeHtml(answer.value || "未回答")}</div>
              </div>
            `,
          )
          .join("")}
      </body>
    </html>
  `);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
}

function renderResponses() {
  const responses = getFilteredResponses();
  const stage = document.querySelector("#responseStage");
  const surveyGroups = groupResponsesBySurvey(responses);

  if (
    state.selectedResponseSurveyId &&
    !surveyGroups.some((group) => String(group.surveyId) === state.selectedResponseSurveyId)
  ) {
    state.selectedResponseSurveyId = "";
    state.selectedResponseId = "";
  }

  const selectedGroup =
    surveyGroups.find((group) => String(group.surveyId) === state.selectedResponseSurveyId) || null;
  const surveyResponses = selectedGroup
    ? selectedGroup.responses.slice().sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))
    : [];

  state.selectedResponseIds = state.selectedResponseIds.filter((responseId) =>
    surveyResponses.some((response) => response.id === responseId),
  );

  if (state.selectedResponseId && !surveyResponses.some((response) => response.id === state.selectedResponseId)) {
    state.selectedResponseId = "";
  }

  const selectedResponse =
    surveyResponses.find((response) => response.id === state.selectedResponseId) || null;
  state.selectedCustomerName = selectedResponse?.customerName || "";

  if (!selectedGroup) {
    stage.innerHTML = `
      <div class="stage-head">
        <div>
          <div class="card-title">アンケートタイトル一覧</div>
          <div class="meta">アンケートを選ぶと、そのアンケートの過去回答を最新順で表示します。</div>
        </div>
      </div>
      <div class="customer-list">
        ${
          surveyGroups.length
            ? surveyGroups
                .map((group) => {
                  const respondentCount = new Set(group.responses.map((response) => response.customerName)).size;
                  return `
                    <button
                      class="customer-card selectable-card"
                      type="button"
                      data-response-survey-id="${escapeHtml(group.surveyId)}"
                    >
                      <strong>${escapeHtml(group.surveyTitle)}</strong>
                      <div>回答数: ${group.count}件 / 回答者: ${respondentCount}名</div>
                      <div class="meta">最新回答: ${formatDate(group.latestAt)}</div>
                      ${
                        group.unreadCount
                          ? `<span class="badge new">未対応 ${group.unreadCount}件</span>`
                          : `<span class="badge checked">未対応なし</span>`
                      }
                    </button>
                  `;
                })
                .join("")
            : `<div class="empty">条件に一致するアンケートはありません。</div>`
        }
      </div>
    `;
  } else if (!selectedResponse) {
    stage.innerHTML = `
      <div class="stage-head">
        <div>
          <div class="card-title">${escapeHtml(selectedGroup.surveyTitle)}</div>
          <div class="meta">過去回答を最新順で表示しています。回答を押すと詳細を表示します。</div>
        </div>
        <button class="secondary-button" type="button" data-back-stage="survey-groups">戻る</button>
      </div>
      <div class="bulk-toolbar">
        <label>
          一括変更
          <select id="bulkStatusSelect">
            <option value="">選択してください</option>
            <option value="new">未対応</option>
            <option value="checked">確認済み</option>
            <option value="done">対応済み</option>
            <option value="trash">ゴミ箱</option>
          </select>
        </label>
        <button class="secondary-button" type="button" id="bulkStatusApplyButton">選択した回答に適用</button>
      </div>
      <div class="response-list">
        ${
          surveyResponses.length
            ? surveyResponses
                .map(
                  (response) => `
                    <article class="response-history-card selectable-card">
                      <label class="inline-toggle history-check">
                        <input type="checkbox" data-select-response-id="${response.id}" ${state.selectedResponseIds.includes(response.id) ? "checked" : ""} />
                        選択
                      </label>
                      <button
                        class="history-open-button"
                        type="button"
                        data-history-response-id="${response.id}"
                      >
                        <strong>${escapeHtml(getCustomerNameWithMember(response.customerName))}</strong>
                        <div class="meta">${formatDate(response.submittedAt)}</div>
                        ${renderTicketStampList(getResponseTicketInfo(response))}
                        ${renderResponsePhotoPreview(response, 3)}
                        <span class="badge ${normalizeStatus(response.status)}">
                          ${STATUS_LABELS[normalizeStatus(response.status)]}
                        </span>
                      </button>
                      <div class="action-row">
                        <button
                          class="secondary-button danger-button"
                          type="button"
                          data-delete-response="${escapeHtml(response.id)}"
                        >
                          ${normalizeStatus(response.status) === "trash" ? "完全削除" : "回答削除"}
                        </button>
                      </div>
                    </article>
                  `,
                )
                .join("")
            : `<div class="empty">このアンケートの回答履歴はありません。</div>`
        }
      </div>
    `;
  } else {
    stage.innerHTML = `
      <div class="stage-head">
        <div>
          <div class="card-title">${escapeHtml(selectedResponse.surveyTitle)}</div>
          <div class="meta">${escapeHtml(getCustomerNameWithMember(selectedResponse.customerName))} / ${formatDate(selectedResponse.submittedAt)}</div>
        </div>
        <div class="action-row">
          <button class="secondary-button" type="button" data-print-response="${selectedResponse.id}">印刷</button>
          <button class="secondary-button" type="button" data-download-response-photos="${selectedResponse.id}">写真保存</button>
          <button class="secondary-button danger-button" type="button" data-delete-response="${selectedResponse.id}">${normalizeStatus(selectedResponse.status) === "trash" ? "完全削除" : "回答削除"}</button>
          <button class="secondary-button" type="button" data-back-stage="survey-history">戻る</button>
          <button class="secondary-button" type="button" data-close-response-detail>閉じる</button>
        </div>
      </div>
      ${renderComparisonSection(selectedResponse)}
      ${renderPhotoComparisonSection(selectedResponse)}
      ${renderResponseCard(selectedResponse)}
    `;
  }

  stage.querySelectorAll("[data-response-survey-id]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedResponseSurveyId = button.dataset.responseSurveyId || "";
      state.selectedResponseId = "";
      state.selectedResponseIds = [];
      renderResponses();
    });
  });

  stage.querySelectorAll("[data-history-response-id]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedResponseId = button.dataset.historyResponseId;
      renderResponses();
    });
  });

  stage.querySelectorAll("[data-select-response-id]").forEach((input) => {
    input.addEventListener("change", () => {
      if (input.checked) {
        state.selectedResponseIds = Array.from(new Set(state.selectedResponseIds.concat(input.dataset.selectResponseId)));
      } else {
        state.selectedResponseIds = state.selectedResponseIds.filter((responseId) => responseId !== input.dataset.selectResponseId);
      }
    });
  });

  stage.querySelectorAll("[data-back-stage]").forEach((button) => {
    button.addEventListener("click", () => {
      if (button.dataset.backStage === "survey-groups") {
        state.selectedResponseSurveyId = "";
        state.selectedResponseId = "";
        state.selectedResponseIds = [];
      } else {
        state.selectedResponseId = "";
      }
      renderResponses();
    });
  });

  stage.querySelector("[data-close-response-detail]")?.addEventListener("click", () => {
    state.selectedResponseSurveyId = "";
    state.selectedResponseId = "";
    state.selectedResponseIds = [];
    renderResponses();
  });

  stage.querySelectorAll("[data-save-response]").forEach((button) => {
    button.addEventListener("click", () => {
      void saveResponseManagement(
        button.dataset.saveResponse,
        button.closest("[data-response-card]"),
      );
    });
  });

  stage.querySelectorAll("[data-delete-response]").forEach((button) => {
    button.addEventListener("click", () => {
      void deleteResponse(button.dataset.deleteResponse);
    });
  });

  stage.querySelector("#bulkStatusApplyButton")?.addEventListener("click", () => {
    void applyBulkStatus();
  });
  stage.querySelectorAll("[data-print-response]").forEach((button) => {
    button.addEventListener("click", () => {
      const response = state.responses.find((item) => item.id === button.dataset.printResponse);
      if (response) printResponse(response);
    });
  });
  stage.querySelectorAll("[data-download-response-photos]").forEach((button) => {
    button.addEventListener("click", () => {
      const response = state.responses.find((item) => item.id === button.dataset.downloadResponsePhotos);
      if (response) {
        void downloadFiles(collectPhotosFromResponses([response]), response.customerName || "response");
      }
    });
  });

  renderCurrentCustomerInfo();
  renderAllResponsesBySurveySection();
  attachLightboxHandlers(stage);
}

function renderResponseCard(response) {
  const status = normalizeStatus(response.status);
  const survey = findSurveyById(response.surveyId);
  const displayAnswers = getDisplayAnswers(response, survey);
  return `
    <article class="response-card" data-response-card="${response.id}">
      <div class="response-head">
        <div>
          <strong>${escapeHtml(getCustomerNameWithMember(response.customerName))}</strong>
          <div class="meta">${formatDate(response.submittedAt)}</div>
          <div class="meta">${escapeHtml(response.surveyTitle)}</div>
        </div>
        <span class="badge ${status}">${STATUS_LABELS[status]}</span>
      </div>
      ${renderResponseTicketInfo(response)}
      ${renderResponsePhotoGallery(response, survey)}
      <div class="answer-list">
        ${displayAnswers
          .filter((answer) => {
            const question = survey?.questions.find((item) => item.id === answer.questionId) || {
              type: answer.type,
            };
            return question.type !== "photo";
          })
          .map((answer) => {
            const question = survey?.questions.find((item) => item.id === answer.questionId) || {
              id: answer.questionId,
              label: answer.label,
              type: answer.type,
              options: [],
            };
            return `
              <div class="answer-item">
                <strong>${escapeHtml(answer.label)}</strong><br />
                ${renderEditableAnswerField(question, answer)}
              </div>
            `;
          })
          .join("")}
      </div>
      <div class="management-grid">
        <label>
          対応状況
          <select data-response-status>
            <option value="new" ${status === "new" ? "selected" : ""}>未対応</option>
            <option value="checked" ${status === "checked" ? "selected" : ""}>確認済み</option>
            <option value="done" ${status === "done" ? "selected" : ""}>対応済み</option>
            <option value="trash" ${status === "trash" ? "selected" : ""}>ゴミ箱</option>
          </select>
        </label>
        <label>
          管理メモ
          <textarea data-response-memo>${escapeHtml(response.adminMemo || "")}</textarea>
        </label>
        <button class="secondary-button" type="button" data-save-response="${response.id}">保存</button>
        <button class="secondary-button danger-button" type="button" data-delete-response="${response.id}">${status === "trash" ? "完全削除" : "ゴミ箱へ移動"}</button>
      </div>
    </article>
  `;
}

async function saveResponseManagement(responseId, sourceCard = null) {
  const card = sourceCard || document.querySelector(`[data-response-card="${responseId}"]`);
  if (!card) return;
  const response = state.responses.find((item) => item.id === responseId);
  const survey = findSurveyById(response?.surveyId);
  const status = card.querySelector("[data-response-status]").value;
  const adminMemo = card.querySelector("[data-response-memo]").value;
  const answers = survey
    ? survey.questions.map((question) => {
        const existingAnswer = response.answers.find((answer) => answer.questionId === question.id) || {
          questionId: question.id,
          label: question.label,
          type: question.type,
          value: "",
        };

        if (question.type === "photo") {
          return {
            questionId: question.id,
            value: existingAnswer.value || "",
          };
        }

        if (question.type === "checkbox") {
          return {
            questionId: question.id,
            value: Array.from(
              card.querySelectorAll(`[data-answer-checkbox="${question.id}"]:checked`),
            ).map((input) => input.value),
          };
        }

        return {
          questionId: question.id,
          value: String(
            card.querySelector(`[data-answer-input="${question.id}"]`)?.value || "",
          ).trim(),
        };
      })
    : [];

  try {
    const result = await api.request(`/api/admin/responses/${encodeURIComponent(responseId)}`, {
      method: "PUT",
      token: state.token,
      body: { status, adminMemo, answers },
    });
    state.responses = state.responses.map((response) =>
      response.id === responseId ? result.response : response,
    );
    showToast("回答管理を保存しました。");
    renderAll();
  } catch (error) {
    showToast(error.message || "保存できませんでした。");
  }
}

async function deleteResponse(responseId) {
  const target = state.responses.find((response) => response.id === responseId);
  if (!target) return;
  if (!confirm(normalizeStatus(target.status) === "trash" ? "この回答を完全削除しますか？" : "この回答をゴミ箱へ移動しますか？")) return;
  try {
    if (normalizeStatus(target.status) === "trash") {
      await api.request(`/api/admin/responses/${encodeURIComponent(responseId)}`, {
        method: "DELETE",
        token: state.token,
      });
      state.responses = state.responses.filter((response) => response.id !== responseId);
      showToast("回答を完全削除しました。");
    } else {
      const result = await api.request(`/api/admin/responses/${encodeURIComponent(responseId)}`, {
        method: "PUT",
        token: state.token,
        body: {
          status: "trash",
          adminMemo: target.adminMemo || "",
          answers: target.answers || [],
        },
      });
      state.responses = state.responses.map((response) =>
        response.id === responseId ? result.response : response,
      );
      showToast("回答をゴミ箱へ移動しました。");
    }
    renderAll();
  } catch (error) {
    showToast(error.message || "削除できませんでした。");
  }
}

function groupByCustomer() {
  return groupByCustomerFrom(
    state.responses.filter((response) => normalizeStatus(response.status) !== "trash"),
  );
}

function groupByCustomerFrom(responses) {
  const map = new Map();
  responses.forEach((response) => {
    const key = response.customerName;
    const current = map.get(key) || {
      name: response.customerName,
      count: 0,
      latestAt: response.submittedAt,
      surveys: new Map(),
    };
    current.count += 1;
    current.latestAt =
      new Date(response.submittedAt) > new Date(current.latestAt)
        ? response.submittedAt
        : current.latestAt;
    current.surveys.set(response.surveyTitle, (current.surveys.get(response.surveyTitle) || 0) + 1);
    map.set(key, current);
  });
  return Array.from(map.values()).sort((a, b) => new Date(b.latestAt) - new Date(a.latestAt));
}

function filterLogs(entries, keyword) {
  const q = String(keyword || "").trim().toLowerCase();
  if (!q) return entries;
  return entries.filter((entry) =>
    JSON.stringify(entry || {}).toLowerCase().includes(q),
  );
}

function renderSettings() {
  const credentialInfo = document.querySelector("#credentialInfo");
  const storageInfo = document.querySelector("#storageInfo");
  const pushStatusInfo = document.querySelector("#pushStatusInfo");
  const preferencesCard = document.querySelector("#preferencesCard");
  const adminUsersCard = document.querySelector("#adminUsersCard");
  const auditLogList = document.querySelector("#auditLogList");
  const errorLogList = document.querySelector("#errorLogList");
  const versionInfo = document.querySelector("#versionInfo");
  const backupMetaInfo = document.querySelector("#backupMetaInfo");
  if (!state.adminInfo) {
    credentialInfo.textContent = "認証情報を読み込み中です。";
    storageInfo.innerHTML = `<div class="empty">保存先情報を読み込み中です。</div>`;
    if (pushStatusInfo) pushStatusInfo.innerHTML = `<div class="empty">通知設定を読み込み中です。</div>`;
    return;
  }

  if (!credentialForm.dataset.dirty) {
    credentialForm.elements.loginId.value = state.adminInfo.adminUsername || "";
  }

  credentialInfo.innerHTML = `
    現在のログインID: ${escapeHtml(state.adminInfo.adminUsername || "")}<br />
    パスワードは入力した場合のみ変更します。
  `;

  const spreadsheetLink = state.adminInfo.spreadsheetUrl
    ? `<a href="${escapeHtml(state.adminInfo.spreadsheetUrl)}" target="_blank" rel="noopener">
        ${escapeHtml(state.adminInfo.spreadsheetUrl)}
      </a>`
    : escapeHtml(state.adminInfo.masterSheetName || "未設定");
  const photoFolderLink = state.adminInfo.photoRootFolderUrl
    ? `<a href="${escapeHtml(state.adminInfo.photoRootFolderUrl)}" target="_blank" rel="noopener">
        ${escapeHtml(state.adminInfo.photoRootFolderName || "未設定")}
      </a>`
    : escapeHtml(state.adminInfo.photoRootFolderName || "未設定");
  const bijirisFolderLink = state.adminInfo.bijirisPostsFolderUrl
    ? `<a href="${escapeHtml(state.adminInfo.bijirisPostsFolderUrl)}" target="_blank" rel="noopener">
        豆知識
      </a>`
    : "未設定";

  storageInfo.innerHTML = `
    <article class="answer-item">
      <strong>回答保存スプレッドシート</strong><br />
      ${spreadsheetLink}
    </article>
    <article class="answer-item">
      <strong>画像保存フォルダ</strong><br />
      ${photoFolderLink}
    </article>
    <article class="answer-item">
      <strong>豆知識フォルダ</strong><br />
      ${bijirisFolderLink}
    </article>
    <article class="answer-item">
      <strong>実行アカウント</strong><br />
      ${escapeHtml(state.adminInfo.ownerEmail || "未取得")}
    </article>
  `;

  if (pushStatusInfo) {
    const customerAppUrl = escapeHtml(state.adminInfo.customerAppUrl || "未設定");
    const pushAppId = escapeHtml(state.adminInfo.pushAppId || "未設定");
    const configured = state.adminInfo.pushConfigured === true;
    pushStatusInfo.innerHTML = `
      <article class="answer-item">
        <strong>送信状態</strong><br />
        <span class="badge ${configured ? "published" : "draft"}">${configured ? "送信可能" : "未設定"}</span>
      </article>
      <article class="answer-item">
        <strong>OneSignal App ID</strong><br />
        ${pushAppId}
      </article>
      <article class="answer-item">
        <strong>顧客アプリURL</strong><br />
        ${state.adminInfo.customerAppUrl
          ? `<a href="${escapeHtml(state.adminInfo.customerAppUrl)}" target="_blank" rel="noopener">${customerAppUrl}</a>`
          : customerAppUrl}
      </article>
      <article class="answer-item">
        <strong>不足情報</strong><br />
        ${
          configured
            ? "OneSignal の通知送信設定は有効です。"
            : "Apps Script のスクリプトプロパティに ONESIGNAL_REST_API_KEY を設定してください。必要に応じて CUSTOMER_APP_URL も設定できます。"
        }
      </article>
      <form id="pushConfigForm" class="stack">
        <div class="survey-question-grid">
          <label>
            OneSignal App ID
            <input name="pushAppId" type="text" value="${pushAppId === "未設定" ? "" : pushAppId}" required />
          </label>
          <label>
            顧客アプリURL
            <input
              name="customerAppUrl"
              type="url"
              value="${state.adminInfo.customerAppUrl ? escapeHtml(state.adminInfo.customerAppUrl) : ""}"
              placeholder="https://..."
              required
            />
          </label>
        </div>
        <label>
          REST APIキー
          <input name="restApiKey" type="password" autocomplete="new-password" placeholder="変更時のみ入力" />
        </label>
        <div class="meta">REST APIキーを空欄で保存すると、現在のキーをそのまま維持します。</div>
        <button class="primary-button" type="submit">通知設定を保存</button>
      </form>
    `;
    pushStatusInfo.querySelector("#pushConfigForm")?.addEventListener("submit", (event) => {
      event.preventDefault();
      void savePushConfig(event.currentTarget);
    });
  }

  if (preferencesCard) {
    const preferences = state.preferences || {
      notificationEnabled: true,
      notificationEmail: state.adminInfo.ownerEmail || "",
      notificationSubject: "",
      notificationBody: "",
      dataPolicyText: "",
      requireConsent: true,
      consentText: "",
      autoBackupEnabled: true,
      backupHour: 3,
      retentionDays: 365,
      recoveryMemo: "",
    };
    preferencesCard.innerHTML = `
      <form id="preferencesForm" class="stack">
        <label class="inline-toggle">
          <input name="notificationEnabled" type="checkbox" ${preferences.notificationEnabled === false ? "" : "checked"} />
          新着回答をメール通知する
        </label>
        <label>
          通知先メールアドレス
          <input name="notificationEmail" type="email" value="${escapeHtml(preferences.notificationEmail || "")}" />
        </label>
        <label>
          通知件名
          <input name="notificationSubject" type="text" value="${escapeHtml(preferences.notificationSubject || "")}" />
        </label>
        <label>
          通知本文
          <textarea name="notificationBody">${escapeHtml(preferences.notificationBody || "")}</textarea>
        </label>
        <label>
          お客様向けデータ方針
          <textarea name="dataPolicyText">${escapeHtml(preferences.dataPolicyText || "")}</textarea>
        </label>
        <label class="inline-toggle">
          <input name="requireConsent" type="checkbox" ${preferences.requireConsent === false ? "" : "checked"} />
          お客様アプリで同意チェックを必須にする
        </label>
        <label>
          同意文言
          <textarea name="consentText">${escapeHtml(preferences.consentText || "")}</textarea>
        </label>
        <label class="inline-toggle">
          <input name="autoBackupEnabled" type="checkbox" ${preferences.autoBackupEnabled === false ? "" : "checked"} />
          毎日自動バックアップを実行する
        </label>
        <div class="survey-question-grid">
          <label>
            バックアップ実行時刻
            <input name="backupHour" type="number" min="0" max="23" value="${escapeHtml(preferences.backupHour || 3)}" />
          </label>
          <label>
            ごみ箱保管日数
            <input name="retentionDays" type="number" min="0" value="${escapeHtml(preferences.retentionDays || 365)}" />
          </label>
        </div>
        <label>
          復旧手順メモ
          <textarea name="recoveryMemo">${escapeHtml(preferences.recoveryMemo || "")}</textarea>
        </label>
        <button class="primary-button" type="submit">設定を保存</button>
      </form>
    `;
    preferencesCard.querySelector("form")?.addEventListener("submit", (event) => {
      event.preventDefault();
      void savePreferences();
    });
  }

  if (adminUsersCard) {
    const users = Array.isArray(state.adminUsers) && state.adminUsers.length
      ? state.adminUsers
      : [{ id: "", username: "", email: "", active: true }];
    adminUsersCard.innerHTML = `
      <form id="adminUsersForm" class="stack">
        <div id="adminUsersList" class="stack">
          ${users
            .map(
              (user, index) => `
                <article class="answer-item admin-user-item" data-admin-user-id="${escapeHtml(user.id || "")}">
                  <strong>管理者 ${index + 1}</strong>
                  <div class="survey-question-grid">
                    <label>
                      ログインID
                      <input name="username" type="text" value="${escapeHtml(user.username || "")}" required />
                    </label>
                    <label>
                      通知メール
                      <input name="email" type="email" value="${escapeHtml(user.email || "")}" />
                    </label>
                  </div>
                  <div class="survey-question-grid">
                    <label>
                      パスワード
                      <input name="password" type="password" placeholder="${user.id ? "変更時のみ入力" : "4文字以上"}" />
                    </label>
                    <label class="inline-toggle">
                      <input name="active" type="checkbox" ${user.active === false ? "" : "checked"} />
                      有効
                    </label>
                  </div>
                  <button class="secondary-button" type="button" data-remove-admin-user>削除</button>
                </article>
              `,
            )
            .join("")}
        </div>
        <div class="action-row">
          <button id="addAdminUserButton" class="secondary-button" type="button">管理者を追加</button>
          <button class="primary-button" type="submit">管理者を保存</button>
        </div>
      </form>
    `;
    adminUsersCard.onclick = (event) => {
      const addButton = event.target.closest("#addAdminUserButton");
      if (addButton) {
        document.querySelector("#adminUsersList")?.insertAdjacentHTML(
          "beforeend",
          `
            <article class="answer-item admin-user-item" data-admin-user-id="">
              <strong>管理者 追加</strong>
              <div class="survey-question-grid">
                <label>
                  ログインID
                  <input name="username" type="text" value="" required />
                </label>
                <label>
                  通知メール
                  <input name="email" type="email" value="" />
                </label>
              </div>
              <div class="survey-question-grid">
                <label>
                  パスワード
                  <input name="password" type="password" placeholder="4文字以上" />
                </label>
                <label class="inline-toggle">
                  <input name="active" type="checkbox" checked />
                  有効
                </label>
              </div>
              <button class="secondary-button" type="button" data-remove-admin-user>削除</button>
            </article>
          `,
        );
        return;
      }
      const removeButton = event.target.closest("[data-remove-admin-user]");
      if (!removeButton) return;
      const items = adminUsersCard.querySelectorAll(".admin-user-item");
      if (items.length <= 1) {
        showToast("管理者アカウントは1件以上必要です。");
        return;
      }
      removeButton.closest(".admin-user-item")?.remove();
    };
    adminUsersCard.querySelector("#adminUsersForm")?.addEventListener("submit", (event) => {
      event.preventDefault();
      void saveAdminUsers(event.currentTarget);
    });
  }

  if (auditLogList) {
    const auditLogs = filterLogs(
      Array.isArray(state.logs?.auditLogs) ? state.logs.auditLogs : [],
      document.querySelector("#auditLogFilter")?.value,
    );
    auditLogList.innerHTML = auditLogs.length
      ? auditLogs
          .map(
            (entry) => `
              <article class="answer-item">
                <strong>${escapeHtml(entry.type || "audit")}</strong>
                <div class="meta">${formatDate(entry.at)}</div>
                <div class="meta">${escapeHtml(JSON.stringify(entry.detail || {}))}</div>
              </article>
            `,
          )
          .join("")
      : `<div class="empty">監査ログはまだありません。</div>`;
  }

  if (errorLogList) {
    const errorLogs = filterLogs(
      Array.isArray(state.logs?.errorLogs) ? state.logs.errorLogs : [],
      document.querySelector("#errorLogFilter")?.value,
    );
    errorLogList.innerHTML = errorLogs.length
      ? errorLogs
          .map(
            (entry) => `
              <article class="answer-item">
                <strong>${escapeHtml(entry.source || "error")}</strong>
                <div class="meta">${formatDate(entry.at)}</div>
                <div>${escapeHtml(entry.message || "")}</div>
              </article>
            `,
          )
          .join("")
      : `<div class="empty">エラーログはまだありません。</div>`;
  }

  if (versionInfo) {
    versionInfo.innerHTML = `
      <div class="meta">管理アプリ Version: ${escapeHtml(state.adminInfo.version || "-")}</div>
      <div class="meta">API モード: ${escapeHtml(api.mode || "-")}</div>
      <div class="meta">バックアップ保存先: ${state.adminInfo.backupFolderUrl ? `<a href="${escapeHtml(state.adminInfo.backupFolderUrl)}" target="_blank" rel="noopener">Google Drive</a>` : "-"}</div>
    `;
  }

  if (backupMetaInfo) {
    const latestBackup = state.adminInfo?.latestBackup || {};
    const lastMaintenance = state.adminInfo?.lastMaintenance || {};
    backupMetaInfo.innerHTML = `
      <article class="answer-item">
        <strong>最終バックアップ</strong>
        <div class="meta">${latestBackup.at ? formatDate(latestBackup.at) : "まだ実行されていません。"}</div>
        ${
          latestBackup.fileUrl
            ? `<div class="meta"><a href="${escapeHtml(latestBackup.fileUrl)}" target="_blank" rel="noopener">${escapeHtml(latestBackup.fileName || "バックアップファイル")}</a></div>`
            : ""
        }
      </article>
      <article class="answer-item">
        <strong>最終メンテナンス</strong>
        <div class="meta">${lastMaintenance.at ? formatDate(lastMaintenance.at) : "まだ実行されていません。"}</div>
        <div class="meta">ごみ箱削除件数: ${Number(lastMaintenance.purged || 0)}件</div>
      </article>
    `;
  }
}

function formatSurveyStatusLabel(status) {
  if (normalizeSurveyStatus(status) === "draft") return "下書き";
  if (normalizeSurveyStatus(status) === "archived") return "アーカイブ";
  return "公開中";
}

function questionSupportsOptions(type) {
  return type === "choice" || type === "checkbox";
}

function makeEditorQuestionId() {
  return `question_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function normalizeVisibilityCondition(condition) {
  const questionId = String(condition?.questionId || "").trim();
  const value = String(condition?.value || "").trim();
  if (!questionId || !value) return null;
  return { questionId, value };
}

function readVisibilityConditionDraft(row) {
  return {
    questionId: String(row?.querySelector("[data-condition-question-id]")?.value || "").trim(),
    value: String(row?.querySelector("[data-condition-value]")?.value || "").trim(),
  };
}

function getQuestionVisibilityConditions(question) {
  const conditions = Array.isArray(question?.visibilityConditions)
    ? question.visibilityConditions.map(normalizeVisibilityCondition).filter(Boolean)
    : [];
  if (conditions.length) return conditions;
  const fallback = normalizeVisibilityCondition(question?.visibleWhen);
  return fallback ? [fallback] : [];
}

function normalizeQuestionForEditor(question) {
  return {
    id: question?.id || makeEditorQuestionId(),
    label: question?.label || "",
    type: question?.type || "text",
    required: question?.required === false ? false : true,
    options: Array.isArray(question?.options) ? question.options : [],
    visibilityConditions: getQuestionVisibilityConditions(question),
  };
}

function getQuestionConditionOptions(question) {
  const type = question?.type || "text";
  if (type === "rating") return ["1", "2", "3", "4", "5"];
  if (!questionSupportsOptions(type)) return [];
  return Array.isArray(question?.options)
    ? question.options.map((option) => String(option || "").trim()).filter(Boolean)
    : [];
}

function isSessionConcernQuestion(question) {
  return String(question?.id || "").trim() === SESSION_CONCERN_QUESTION_ID;
}

function getSessionConcernCategoryGroups(options) {
  const normalizedOptions = Array.isArray(options)
    ? options.map((option) => String(option || "").trim()).filter(Boolean)
    : [];
  const definedValues = new Set(SESSION_CONCERN_CATEGORIES.flatMap((category) => category.options));
  const groups = SESSION_CONCERN_CATEGORIES.map((category) => ({
    id: category.id,
    label: category.label,
    options: normalizedOptions.filter((option) => category.options.includes(option)),
  }));
  const extras = normalizedOptions.filter((option) => !definedValues.has(option));
  if (extras.length) {
    groups.push({
      id: "custom",
      label: "【その他の項目】",
      options: extras,
    });
  }
  return groups;
}

function renderQuestionOptionsEditorContent(question) {
  const normalizedQuestion = normalizeQuestionForEditor(question);
  const type = normalizedQuestion.type;
  if (!questionSupportsOptions(type)) return "";

  if (type === "checkbox" && isSessionConcernQuestion(normalizedQuestion)) {
    const groups = getSessionConcernCategoryGroups(normalizedQuestion.options);
    return `
      <div class="category-editor-list">
        <div class="meta">カテゴリは固定表示です。各欄を1行1項目で編集できます。</div>
        ${groups
          .map(
            (group) => `
              <label class="category-editor-card">
                <span>${escapeHtml(group.label)}</span>
                <textarea data-category-options-input data-category-id="${escapeHtml(group.id)}" placeholder="1行に1項目">${escapeHtml(group.options.join("\n"))}</textarea>
              </label>
            `,
          )
          .join("")}
      </div>
    `;
  }

  const options = Array.isArray(normalizedQuestion.options) && normalizedQuestion.options.length
    ? normalizedQuestion.options
    : ["", ""];

  return `
    <div class="option-list" data-option-list>
      ${options
        .map(
          (option) => `
            <div class="option-row">
              <input type="text" data-option-input value="${escapeHtml(option)}" placeholder="選択肢" />
              <button class="secondary-button" type="button" data-remove-option>削除</button>
            </div>
          `,
        )
        .join("")}
    </div>
    <button class="secondary-button" type="button" data-add-option>選択肢を追加</button>
  `;
}

function makeBlankQuestion() {
  return {
    id: makeEditorQuestionId(),
    label: "",
    type: "text",
    required: true,
    options: [],
    visibilityConditions: [],
  };
}

function readQuestionOptionsFromBlock(block) {
  if (!block) return [];
  if (block.querySelectorAll("[data-category-options-input]").length) {
    return Array.from(block.querySelectorAll("[data-category-options-input]"))
      .flatMap((input) =>
        String(input.value || "")
          .split(/\r?\n/)
          .map((value) => value.trim())
          .filter(Boolean),
      );
  }
  return Array.from(block.querySelectorAll("[data-option-input]"))
    .map((input) => String(input.value || "").trim())
    .filter(Boolean);
}

function makeBlankSurveyDraft() {
  return {
    id: "",
    title: "",
    description: "",
    introMessage: "",
    completionMessage: "",
    status: "published",
    sortOrder: 0,
    acceptingResponses: true,
    startAt: "",
    endAt: "",
    questions: [makeBlankQuestion()],
    createdAt: "",
    updatedAt: "",
  };
}

function getSurveyEditorDraft() {
  const selected = findSurveyById(state.selectedSurveyEditorId);
  if (!selected) return makeBlankSurveyDraft();
  return {
    ...selected,
    questions: Array.isArray(selected.questions) && selected.questions.length
      ? selected.questions.map(normalizeQuestionForEditor)
      : [makeBlankQuestion()],
  };
}

function toDateTimeLocalValue(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  const offset = date.getTimezoneOffset();
  const local = new Date(date.getTime() - offset * 60000);
  return local.toISOString().slice(0, 16);
}

function renderVisibilityConditionEditor(condition, previousQuestions) {
  const selectedQuestionId = String(condition?.questionId || "").trim();
  const selectedQuestion =
    previousQuestions.find((item) => String(item.id || "").trim() === selectedQuestionId) || null;
  const valueOptions = getQuestionConditionOptions(selectedQuestion);
  const currentValue = String(condition?.value || "").trim();
  const hasCustomValue = currentValue && valueOptions.length && !valueOptions.includes(currentValue);
  return `
    <div class="condition-row" data-visibility-condition-row>
      <select data-condition-question-id>
        <option value="">質問を選択</option>
        ${previousQuestions
          .map(
            (item) => `
              <option value="${escapeHtml(item.id || "")}" ${condition?.questionId === item.id ? "selected" : ""}>
                ${escapeHtml(item.label || item.id || "質問")}
              </option>
            `,
          )
          .join("")}
      </select>
      ${
        valueOptions.length
          ? `
              <select data-condition-value>
                <option value="">値を選択</option>
                ${valueOptions
                  .map(
                    (option) => `
                      <option value="${escapeHtml(option)}" ${currentValue === option ? "selected" : ""}>
                        ${escapeHtml(option)}
                      </option>
                    `,
                  )
                  .join("")}
                ${
                  hasCustomValue
                    ? `<option value="${escapeHtml(currentValue)}" selected>${escapeHtml(currentValue)}</option>`
                    : ""
                }
              </select>
            `
          : `<input type="text" data-condition-value value="${escapeHtml(currentValue)}" placeholder="一致する値" />`
      }
      <button class="secondary-button" type="button" data-remove-visibility-condition>削除</button>
    </div>
  `;
}

function renderSurveyQuestionEditor(question, index, allQuestions) {
  const normalizedQuestion = normalizeQuestionForEditor(question);
  const type = normalizedQuestion.type;
  const supportsOptions = questionSupportsOptions(type);
  const previousQuestions = Array.isArray(allQuestions)
    ? allQuestions.slice(0, index).map(normalizeQuestionForEditor).filter((item) => item?.id || item?.label)
    : [];
  const visibilityConditions = getQuestionVisibilityConditions(normalizedQuestion);

  return `
    <article class="survey-question-card" data-question-block data-question-id="${escapeHtml(normalizedQuestion.id || "")}">
      <div class="survey-question-head">
        <strong data-question-number>質問 ${index + 1}</strong>
        <button class="secondary-button" type="button" data-remove-question>削除</button>
      </div>
      <div class="question-conditions">
        <strong class="question-section-title">表示条件</strong>
        <div class="meta">表示条件を追加すると、すべて一致したときだけこの質問を表示します。</div>
        <div class="condition-list" data-visibility-condition-list>
          ${visibilityConditions
            .map((condition) => renderVisibilityConditionEditor(condition, previousQuestions))
            .join("")}
        </div>
        <button class="secondary-button" type="button" data-add-visibility-condition>表示条件を追加</button>
      </div>
      <div class="question-conditions">
        <strong class="question-section-title">質問内容</strong>
        <div class="survey-question-grid">
          <label>
            質問文
            <input type="text" data-question-label value="${escapeHtml(normalizedQuestion.label || "")}" required />
          </label>
          <label>
            質問形式
            <select data-question-type>
              <option value="text" ${type === "text" ? "selected" : ""}>テキスト</option>
              <option value="textarea" ${type === "textarea" ? "selected" : ""}>長文</option>
              <option value="choice" ${type === "choice" ? "selected" : ""}>単一選択</option>
              <option value="checkbox" ${type === "checkbox" ? "selected" : ""}>複数選択</option>
              <option value="rating" ${type === "rating" ? "selected" : ""}>5段階評価</option>
              <option value="photo" ${type === "photo" ? "selected" : ""}>写真</option>
            </select>
          </label>
        </div>
        <label class="inline-toggle">
          <input type="checkbox" data-question-required ${normalizedQuestion.required === false ? "" : "checked"} />
          表示されたときに必須回答
        </label>
      </div>
      <div data-options-wrapper ${supportsOptions ? "" : "hidden"}>
        ${renderQuestionOptionsEditorContent(normalizedQuestion)}
      </div>
    </article>
  `;
}

function syncSurveyQuestionNumbers(container) {
  container.querySelectorAll("[data-question-block]").forEach((block, index) => {
    const label = block.querySelector("[data-question-number]");
    if (label) label.textContent = `質問 ${index + 1}`;
  });
}

function addSurveyQuestionEditor(container, question) {
  const questions = Array.from(container.querySelectorAll("[data-question-block]")).map((block) => ({
    id: block.dataset.questionId || "",
    label: block.querySelector("[data-question-label]")?.value || "",
  }));
  const nextQuestion = normalizeQuestionForEditor(question || makeBlankQuestion());
  container.insertAdjacentHTML(
    "beforeend",
    renderSurveyQuestionEditor(
      nextQuestion,
      container.children.length,
      questions.concat(nextQuestion),
    ),
  );
  syncSurveyQuestionNumbers(container);
}

function getEditorQuestionReferences(container) {
  return Array.from(container?.querySelectorAll("[data-question-block]") || []).map((block) => {
    const type = String(block.querySelector("[data-question-type]")?.value || "text");
    return {
      id: block.dataset.questionId || "",
      label: String(block.querySelector("[data-question-label]")?.value || "").trim(),
      type,
      options: questionSupportsOptions(type)
        ? readQuestionOptionsFromBlock(block)
        : [],
    };
  });
}

function getPreviousQuestionReferences(block) {
  const container = block?.parentElement;
  if (!container) return [];
  const questions = getEditorQuestionReferences(container);
  const index = Array.from(container.querySelectorAll("[data-question-block]")).indexOf(block);
  if (index <= 0) return [];
  return questions.slice(0, index);
}

function addVisibilityConditionEditor(block, condition = { questionId: "", value: "" }) {
  const list = block?.querySelector("[data-visibility-condition-list]");
  if (!list) return;
  list.insertAdjacentHTML(
    "beforeend",
    renderVisibilityConditionEditor(condition, getPreviousQuestionReferences(block)),
  );
}

function refreshVisibilityConditionEditors(container) {
  Array.from(container?.querySelectorAll("[data-question-block]") || []).forEach((block) => {
    const list = block.querySelector("[data-visibility-condition-list]");
    if (!list) return;
    const drafts = Array.from(list.querySelectorAll("[data-visibility-condition-row]")).map(
      readVisibilityConditionDraft,
    );
    const previousQuestions = getPreviousQuestionReferences(block);
    list.innerHTML = drafts
      .map((condition) => renderVisibilityConditionEditor(condition, previousQuestions))
      .join("");
  });
}

function ensureQuestionOptions(block) {
  const list = block.querySelector("[data-option-list]");
  if (!list || list.children.length) return;
  list.insertAdjacentHTML(
    "beforeend",
    `
      <div class="option-row">
        <input type="text" data-option-input value="" placeholder="選択肢" />
        <button class="secondary-button" type="button" data-remove-option>削除</button>
      </div>
      <div class="option-row">
        <input type="text" data-option-input value="" placeholder="選択肢" />
        <button class="secondary-button" type="button" data-remove-option>削除</button>
      </div>
    `,
  );
}

function refreshQuestionOptionsEditor(block) {
  const wrapper = block?.querySelector("[data-options-wrapper]");
  if (!block || !wrapper) return;
  const type = String(block.querySelector("[data-question-type]")?.value || "text");
  if (!questionSupportsOptions(type)) {
    wrapper.hidden = true;
    wrapper.innerHTML = "";
    return;
  }

  const currentQuestion = {
    id: block.dataset.questionId || "",
    type,
    options: readQuestionOptionsFromBlock(block),
  };
  wrapper.hidden = false;
  wrapper.innerHTML = renderQuestionOptionsEditorContent(currentQuestion);
  ensureQuestionOptions(block);
}

function buildSurveyPayload(form) {
  const title = String(form.elements.title.value || "").trim();
  const description = String(form.elements.description.value || "").trim();
  const introMessage = String(form.elements.introMessage?.value || "").trim();
  const completionMessage = String(form.elements.completionMessage?.value || "").trim();
  const status = normalizeSurveyStatus(form.elements.status.value);
  const acceptingResponses = Boolean(form.elements.acceptingResponses?.checked);
  const startAt = String(form.elements.startAt?.value || "").trim();
  const endAt = String(form.elements.endAt?.value || "").trim();
  const questionBlocks = Array.from(form.querySelectorAll("[data-question-block]"));

  const questions = questionBlocks.map((block) => {
    const type = String(block.querySelector("[data-question-type]")?.value || "text");
    const visibilityConditions = Array.from(
      block.querySelectorAll("[data-visibility-condition-row]"),
    )
      .map((row) =>
        normalizeVisibilityCondition({
          questionId: row.querySelector("[data-condition-question-id]")?.value,
          value: row.querySelector("[data-condition-value]")?.value,
        }),
      )
      .filter(Boolean);
    return {
      id: block.dataset.questionId || "",
      label: String(block.querySelector("[data-question-label]")?.value || "").trim(),
      type,
      required: Boolean(block.querySelector("[data-question-required]")?.checked),
      options: questionSupportsOptions(type)
        ? readQuestionOptionsFromBlock(block)
        : [],
      visibilityConditions,
      visibleWhen: visibilityConditions[0] || null,
    };
  });

  return {
    title,
    description,
    introMessage,
    completionMessage,
    status,
    acceptingResponses,
    startAt,
    endAt,
    questions,
  };
}

async function persistSurveyOrder(nextSurveys) {
  const surveys = nextSurveys.map((survey, index) => ({
    ...survey,
    sortOrder: index,
  }));
  const result = await api.request("/api/admin/surveys/replace", {
    method: "POST",
    token: state.token,
    body: { surveys },
  });
  state.surveys = result.surveys || surveys;
}

async function duplicateSurvey(surveyId) {
  const survey = findSurveyById(surveyId);
  if (!survey) return;
  const payload = {
    ...cloneSurveyForEditor(survey),
    title: `${survey.title}（複製）`,
    status: "draft",
  };
  delete payload.id;
  try {
    const result = await api.request("/api/admin/surveys", {
      method: "POST",
      token: state.token,
      body: payload,
    });
    state.surveys = [result.survey, ...state.surveys];
    state.selectedSurveyEditorId = result.survey.id;
    renderAll();
    setPage("surveys");
    showToast("アンケートを複製しました。");
  } catch (error) {
    showToast(error.message || "アンケートを複製できませんでした。");
  }
}

function cloneSurveyForEditor(survey) {
  return JSON.parse(JSON.stringify(survey));
}

async function moveSurvey(surveyId, direction) {
  const index = state.surveys.findIndex((survey) => survey.id === surveyId);
  const nextIndex = index + direction;
  if (index < 0 || nextIndex < 0 || nextIndex >= state.surveys.length) return;
  const nextSurveys = state.surveys.slice();
  const current = nextSurveys[index];
  nextSurveys[index] = nextSurveys[nextIndex];
  nextSurveys[nextIndex] = current;
  try {
    await persistSurveyOrder(nextSurveys);
    renderAll();
    setPage("surveys");
  } catch (error) {
    showToast(error.message || "並び順を更新できませんでした。");
  }
}

async function archiveSurvey(surveyId) {
  const survey = findSurveyById(surveyId);
  if (!survey) return;
  try {
    const result = await api.request(`/api/admin/surveys/${encodeURIComponent(surveyId)}`, {
      method: "PUT",
      token: state.token,
      body: {
        ...cloneSurveyForEditor(survey),
        status: "archived",
      },
    });
    state.surveys = state.surveys.map((item) => (item.id === surveyId ? result.survey : item));
    renderAll();
    setPage("surveys");
    showToast("アンケートをアーカイブしました。");
  } catch (error) {
    showToast(error.message || "アーカイブできませんでした。");
  }
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

function createEmptyBijirisPostDraft() {
  return normalizeBijirisPost({
    id: "",
    title: "",
    category: BIJIRIS_POST_CATEGORY_OPTIONS[0],
    summary: "",
    body: "",
    status: "published",
    pinned: false,
    notifyCustomers: false,
    notificationTitle: "",
    notificationBody: "",
    photos: [],
    documents: [],
  });
}

function cloneBijirisPostForEditor(post) {
  return normalizeBijirisPost(JSON.parse(JSON.stringify(post || createEmptyBijirisPostDraft())));
}

function getBijirisPostById(postId) {
  return state.bijirisPosts.find((post) => post.id === postId) || null;
}

function getBijirisEditorDraft() {
  if (state.bijirisEditorDraft) return state.bijirisEditorDraft;
  const selected = getBijirisPostById(state.selectedBijirisPostId);
  state.bijirisEditorDraft = selected ? cloneBijirisPostForEditor(selected) : createEmptyBijirisPostDraft();
  return state.bijirisEditorDraft;
}

function getBijirisTemplateById(templateId) {
  return BIJIRIS_POST_TEMPLATES.find((template) => template.id === templateId) || null;
}

function applyBijirisTemplate(templateId) {
  const template = getBijirisTemplateById(templateId);
  if (!template) throw new Error("テンプレートが見つかりません。");
  const draft = getBijirisEditorDraft();
  draft.title = template.title;
  draft.category = template.category;
  draft.summary = template.summary;
  draft.body = template.body;
  state.bijirisEditorDraft = draft;
}

function renderAdminBijirisDocumentPreview(file, index, compact = false) {
  const href = String(file?.downloadUrl || file?.previewUrl || file?.url || "#").trim();
  return `
    <a class="history-card bijiris-preview-document ${compact ? "compact" : ""}" href="${escapeHtml(href)}" target="_blank" rel="noreferrer">
      <img class="bijiris-preview-document-thumb" src="${escapeHtml(getBijirisDocumentThumbnailSrc(file))}" alt="${escapeHtml(file?.name || `資料${index + 1}`)}" />
      <strong>${escapeHtml(file?.name || `資料${index + 1}`)}</strong>
      <div class="meta">PDFを開く</div>
    </a>
  `;
}

function renderAdminBijirisPreviewCard(post) {
  const publishedAt = post.publishedAt || post.updatedAt || post.createdAt || new Date().toISOString();
  const preview = post.summary || post.body.slice(0, 90);
  return `
    <article class="history-card bijiris-preview-card">
      <div class="section-head">
        <div>
          <strong>${escapeHtml(post.title || "タイトル未入力")}</strong>
          <div class="meta">${escapeHtml(post.category || "豆知識")} / ${escapeHtml(formatDate(publishedAt))}</div>
        </div>
        <div class="action-row">
          ${post.pinned ? `<span class="badge open">重要</span>` : ""}
          ${post.status === "draft" ? `<span class="badge draft">下書き</span>` : ""}
          ${post.documents.length ? `<span class="badge">PDF ${post.documents.length}</span>` : ""}
          ${post.photos.length ? `<span class="badge draft">写真 ${post.photos.length}</span>` : ""}
        </div>
      </div>
      ${preview ? `<div class="meta">${escapeHtml(preview)}</div>` : `<div class="meta">概要または本文が入ると一覧に表示されます。</div>`}
      ${
        post.documents.length
          ? `<div class="bijiris-preview-document-grid compact">${post.documents.slice(0, 2).map((file, index) => renderAdminBijirisDocumentPreview(file, index, true)).join("")}</div>`
          : ""
      }
    </article>
  `;
}

function renderAdminBijirisPreviewDetail(post) {
  return `
    <article class="history-card bijiris-preview-detail">
      <div class="section-head">
        <div>
          <strong>${escapeHtml(post.title || "タイトル未入力")}</strong>
          <div class="meta">${escapeHtml(post.category || "豆知識")}</div>
        </div>
        <div class="action-row">
          ${post.pinned ? `<span class="badge open">重要</span>` : ""}
          <span class="badge ${escapeHtml(post.status)}">${escapeHtml(post.status === "published" ? "公開" : post.status === "archived" ? "アーカイブ" : "下書き")}</span>
        </div>
      </div>
      ${post.summary ? `<div class="history-card"><strong>概要</strong><div class="survey-editor-preview-text">${escapeHtml(post.summary).replaceAll("\n", "<br />")}</div></div>` : ""}
      <div class="history-card">
        <strong>内容</strong>
        <div class="survey-editor-preview-text">${post.body ? escapeHtml(post.body).replaceAll("\n", "<br />") : "本文はありません。"}</div>
      </div>
      <div class="history-card">
        <strong>写真</strong>
        ${
          post.photos.length
            ? `
              <div class="history-photo-grid">
                ${post.photos.map((file, index) => `
                  <article class="history-photo-card">
                    ${getPhotoPreviewSrc(file) ? `<img src="${escapeHtml(getPhotoPreviewSrc(file))}" alt="${escapeHtml(file.name || `写真${index + 1}`)}" />` : ""}
                    <div class="history-photo-meta"><strong>${escapeHtml(file.name || `写真${index + 1}`)}</strong></div>
                  </article>
                `).join("")}
              </div>
            `
            : `<div class="empty">写真はありません。</div>`
        }
      </div>
      <div class="history-card">
        <strong>PDF資料</strong>
        ${
          post.documents.length
            ? `<div class="bijiris-preview-document-grid">${post.documents.map((file, index) => renderAdminBijirisDocumentPreview(file, index)).join("")}</div>`
            : `<div class="empty">PDFはありません。</div>`
        }
      </div>
    </article>
  `;
}

function refreshBijirisPreview() {
  const draft = normalizeBijirisPost(getBijirisEditorDraft());
  const listPreview = document.querySelector("#bijirisPreviewList");
  const detailPreview = document.querySelector("#bijirisPreviewDetail");
  if (listPreview) {
    listPreview.innerHTML = renderAdminBijirisPreviewCard(draft);
  }
  if (detailPreview) {
    detailPreview.innerHTML = renderAdminBijirisPreviewDetail(draft);
  }
}

function syncBijirisDraftFromForm(form) {
  if (!form) return getBijirisEditorDraft();
  const draft = getBijirisEditorDraft();
  draft.title = String(form.elements.title?.value || "").trim();
  draft.category = String(form.elements.category?.value || "").trim();
  draft.summary = String(form.elements.summary?.value || "").trim();
  draft.body = String(form.elements.body?.value || "").trim();
  draft.status = ["published", "draft", "archived"].includes(String(form.elements.status?.value || "").trim())
    ? String(form.elements.status.value).trim()
    : "published";
  draft.pinned = Boolean(form.elements.pinned?.checked);
  draft.notifyCustomers = Boolean(form.elements.notifyCustomers?.checked);
  draft.notificationTitle = String(form.elements.notificationTitle?.value || "").trim();
  draft.notificationBody = String(form.elements.notificationBody?.value || "").trim();
  return draft;
}

function readFileAsDataUrl(file, errorMessage = "ファイルを読み込めませんでした。") {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => resolve(String(reader.result || "")));
    reader.addEventListener("error", () => reject(new Error(errorMessage)));
    reader.readAsDataURL(file);
  });
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.addEventListener("load", () => resolve(image));
    image.addEventListener("error", () => reject(new Error("写真を読み込めませんでした。")));
    image.src = src;
  });
}

async function prepareBijirisImageFile(file) {
  if (!String(file?.type || "").startsWith("image/")) {
    throw new Error("写真ファイルを選択してください。");
  }
  const source = await readFileAsDataUrl(file, "写真を読み込めませんでした。");
  const image = await loadImage(source);
  const maxSize = 1600;
  const quality = 0.8;
  const scale = Math.min(1, maxSize / Math.max(image.naturalWidth, image.naturalHeight));
  const width = Math.max(1, Math.round(image.naturalWidth * scale));
  const height = Math.max(1, Math.round(image.naturalHeight * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  context.drawImage(image, 0, 0, width, height);
  return {
    kind: "photo",
    name: file.name || "photo.jpg",
    type: "image/jpeg",
    dataUrl: canvas.toDataURL("image/jpeg", quality),
  };
}

async function prepareBijirisDocumentFile(file) {
  const isPdf = String(file?.type || "").includes("pdf") || /\.pdf$/i.test(String(file?.name || ""));
  if (!isPdf) {
    throw new Error("PDF ファイルを選択してください。");
  }
  const source = await readFileAsDataUrl(file, "PDF を読み込めませんでした。");
  const match = source.match(/^data:application\/pdf;base64,(.+)$/i);
  if (!match) {
    throw new Error("PDF データの形式が正しくありません。");
  }
  return {
    kind: "pdf",
    name: file.name || "document.pdf",
    mimeType: "application/pdf",
    type: "application/pdf",
    base64Data: match[1],
  };
}

function renderBijirisAttachmentPreview(files, kind) {
  if (!files.length) {
    return `<div class="empty">${kind === "pdf" ? "PDF はまだありません。" : "写真はまだありません。"}</div>`;
  }
  if (kind === "pdf") {
    return `
      <div class="bijiris-attachment-list">
        ${files
          .map((file, index) => `
            <article class="history-photo-badge bijiris-attachment-chip">
              <span>${escapeHtml(file.name || `資料${index + 1}`)}</span>
              <button class="ghost-button" type="button" data-remove-bijiris-document="${index}">削除</button>
            </article>
          `)
          .join("")}
      </div>
    `;
  }
  return `
    <div class="history-photo-grid">
      ${files
        .map((file, index) => {
          const preview = getPhotoPreviewSrc(file);
          return `
            <article class="history-photo-card">
              ${preview ? `<img src="${escapeHtml(preview)}" alt="${escapeHtml(file.name || `写真${index + 1}`)}" />` : ""}
              <div class="history-photo-meta">
                <strong>${escapeHtml(file.name || `写真${index + 1}`)}</strong>
                <button class="ghost-button" type="button" data-remove-bijiris-photo="${index}">削除</button>
              </div>
            </article>
          `;
        })
        .join("")}
    </div>
  `;
}

async function handleBijirisImageInput(input) {
  const form = input?.form;
  syncBijirisDraftFromForm(form);
  const files = Array.from(input?.files || []);
  if (!files.length) return;
  const prepared = [];
  for (const file of files) {
    prepared.push(await prepareBijirisImageFile(file));
  }
  state.bijirisEditorDraft.photos = state.bijirisEditorDraft.photos.concat(prepared).slice(0, 8);
  renderBijirisManager();
  setPage("bijiris");
}

async function handleBijirisDocumentInput(input) {
  const form = input?.form;
  syncBijirisDraftFromForm(form);
  const files = Array.from(input?.files || []);
  if (!files.length) return;
  const prepared = [];
  for (const file of files) {
    prepared.push(await prepareBijirisDocumentFile(file));
  }
  state.bijirisEditorDraft.documents = state.bijirisEditorDraft.documents.concat(prepared).slice(0, 5);
  renderBijirisManager();
  setPage("bijiris");
}

async function saveBijirisPost() {
  const draft = normalizeBijirisPost(getBijirisEditorDraft());
  if (!draft.title) throw new Error("タイトルを入力してください。");
  if (!draft.summary && !draft.body && !draft.photos.length && !draft.documents.length) {
    throw new Error("本文、要約、写真、PDF のいずれかを入力してください。");
  }
  const payload = {
    id: draft.id,
    title: draft.title,
    category: draft.category,
    summary: draft.summary,
    body: draft.body,
    status: draft.status,
    pinned: draft.pinned,
    publishedAt: draft.publishedAt,
    notifyCustomers: draft.notifyCustomers === true,
    notificationTitle: draft.notificationTitle,
    notificationBody: draft.notificationBody,
    photos: draft.photos,
    documents: draft.documents,
  };
  const result = draft.id
    ? await api.request(`/api/admin/bijiris-posts/${encodeURIComponent(draft.id)}`, {
        method: "PUT",
        token: state.token,
        body: payload,
      })
    : await api.request("/api/admin/bijiris-posts", {
        method: "POST",
        token: state.token,
        body: payload,
      });
  await loadAdminData();
  state.selectedBijirisPostId = result.post?.id || draft.id;
  state.bijirisEditorDraft = cloneBijirisPostForEditor(result.post || getBijirisPostById(state.selectedBijirisPostId));
  state.selectedBijirisView = "history";
  renderAll();
  setPage("bijiris");
  showToast(draft.id ? "豆知識を更新しました。" : "豆知識を作成しました。");
}

async function deleteBijirisPost(postId) {
  const targetId = String(postId || state.selectedBijirisPostId || "").trim();
  if (!targetId) return;
  await api.request(`/api/admin/bijiris-posts/${encodeURIComponent(targetId)}`, {
    method: "DELETE",
    token: state.token,
  });
  await loadAdminData();
  state.selectedBijirisPostId = "";
  state.bijirisEditorDraft = createEmptyBijirisPostDraft();
  state.selectedBijirisView = "list";
  renderAll();
  setPage("bijiris");
  showToast("豆知識を削除しました。");
}

function renderBijirisHistoryDetail(post) {
  return `
    <div class="survey-history-list">
      <article class="history-card">
        <div class="survey-manager-card-head">
          <strong>${escapeHtml(post.title)}</strong>
          <span class="badge ${escapeHtml(post.status)}">${escapeHtml(post.status === "published" ? "公開" : post.status === "archived" ? "アーカイブ" : "下書き")}</span>
        </div>
        <div>${escapeHtml(post.category || "豆知識")}</div>
        <div class="meta">作成: ${post.createdAt ? escapeHtml(formatDate(post.createdAt)) : "-"}</div>
        <div class="meta">更新: ${post.updatedAt ? escapeHtml(formatDate(post.updatedAt)) : "-"}</div>
        <div class="meta">写真 ${post.photos.length} / PDF ${post.documents.length}</div>
        <div class="action-row">
          ${post.pinned ? `<span class="badge open">重要固定</span>` : ""}
          ${post.status === "draft" ? `<span class="badge draft">顧客アプリにはまだ表示されません</span>` : ""}
        </div>
      </article>
      ${renderAdminBijirisPreviewDetail(post)}
    </div>
  `;
}

function renderBijirisManager() {
  const list = document.querySelector("#bijirisPostList");
  const editorCard = document.querySelector("#bijirisEditorCard");
  const createButton = document.querySelector("#createBijirisPostButton");
  const listSection = document.querySelector("#bijirisListSection");
  const editorSection = document.querySelector("#bijirisEditorSection");
  const stageTitle = document.querySelector("#bijirisStageTitle");
  if (!list || !editorCard || !createButton || !listSection || !editorSection || !stageTitle) return;

  if (state.selectedBijirisPostId && !getBijirisPostById(state.selectedBijirisPostId)) {
    state.selectedBijirisPostId = "";
    state.bijirisEditorDraft = null;
    state.selectedBijirisView = "list";
  }

  const posts = sortBijirisPosts(state.bijirisPosts);
  const selectedPost = getBijirisPostById(state.selectedBijirisPostId);
  const view = selectedPost && state.selectedBijirisView === "history"
    ? "history"
    : state.selectedBijirisView === "editor"
      ? "editor"
      : "list";
  state.selectedBijirisView = view;

  stageTitle.textContent = view === "history" ? "過去の履歴" : "投稿一覧";
  listSection.hidden = view === "editor";
  editorSection.hidden = view !== "editor";
  createButton.hidden = view !== "list";

  if (view === "history" && selectedPost) {
    list.innerHTML = `
      <div class="survey-editor-head">
        <div>
          <div class="card-title">過去の履歴</div>
          <div class="meta">投稿内容の確認、編集、削除ができます。</div>
        </div>
        <div class="action-row">
          <button class="secondary-button" type="button" data-bijiris-back-list>投稿一覧へ戻る</button>
          <button class="primary-button" type="button" data-bijiris-open-editor="${escapeHtml(selectedPost.id)}">編集する</button>
          <button class="secondary-button danger-button" type="button" data-bijiris-delete-post="${escapeHtml(selectedPost.id)}">削除</button>
        </div>
      </div>
      ${renderBijirisHistoryDetail(selectedPost)}
    `;
    editorCard.innerHTML = "";
    list.querySelector("[data-bijiris-back-list]")?.addEventListener("click", () => {
      state.selectedBijirisView = "list";
      renderBijirisManager();
      setPage("bijiris");
    });
    list.querySelector("[data-bijiris-open-editor]")?.addEventListener("click", () => {
      state.selectedBijirisPostId = selectedPost.id;
      state.bijirisEditorDraft = cloneBijirisPostForEditor(selectedPost);
      state.selectedBijirisView = "editor";
      renderBijirisManager();
      setPage("bijiris");
    });
    list.querySelector("[data-bijiris-delete-post]")?.addEventListener("click", () => {
      deleteBijirisPost(selectedPost.id).catch((error) => {
        showToast(error.message || "豆知識を削除できませんでした。");
      });
    });
    return;
  }

  if (view === "list") {
    list.innerHTML = posts.length
      ? posts
          .map((post) => `
            <article class="survey-manager-card selectable-card ${post.id === state.selectedBijirisPostId ? "active" : ""}">
              <button class="survey-open-button" type="button" data-open-bijiris-history="${escapeHtml(post.id)}">
                <div class="survey-manager-card-head">
                  <strong>${escapeHtml(post.title)}</strong>
                  <span class="badge ${escapeHtml(post.status)}">${escapeHtml(post.status === "published" ? "公開" : post.status === "archived" ? "アーカイブ" : "下書き")}</span>
                </div>
                <div>${escapeHtml(post.category || "豆知識")}</div>
                <div class="meta">更新: ${post.updatedAt ? escapeHtml(formatDate(post.updatedAt)) : "-"}</div>
                <div class="meta">写真 ${post.photos.length} / PDF ${post.documents.length}</div>
                ${post.pinned ? `<span class="badge open">重要固定</span>` : ""}
              </button>
            </article>
          `)
          .join("")
      : `<div class="empty">豆知識の投稿はまだありません。</div>`;
    editorCard.innerHTML = "";
    createButton.onclick = () => {
      state.selectedBijirisPostId = "";
      state.bijirisEditorDraft = createEmptyBijirisPostDraft();
      state.selectedBijirisView = "editor";
      renderBijirisManager();
      setPage("bijiris");
    };
    list.querySelectorAll("[data-open-bijiris-history]").forEach((button) => {
      button.addEventListener("click", () => {
        state.selectedBijirisPostId = button.dataset.openBijirisHistory || "";
        state.selectedBijirisView = "history";
        renderBijirisManager();
        setPage("bijiris");
      });
    });
    return;
  }

  const draft = getBijirisEditorDraft();
  editorCard.innerHTML = `
    <div class="survey-editor-head">
      <div>
        <div class="card-title">${selectedPost ? "豆知識を編集" : "豆知識を新規作成"}</div>
        <div class="meta">
          ${selectedPost
            ? `作成: ${draft.createdAt ? formatDate(draft.createdAt) : "-"} / 更新: ${draft.updatedAt ? formatDate(draft.updatedAt) : "-"}`
            : "顧客アプリに表示する内容を登録します。"}
        </div>
      </div>
      <div class="action-row">
        <button class="secondary-button" type="button" data-bijiris-cancel-editor>${selectedPost ? "過去の履歴へ戻る" : "投稿一覧へ戻る"}</button>
        ${selectedPost ? `<button class="secondary-button danger-button" type="button" data-delete-bijiris-post="${escapeHtml(selectedPost.id)}">削除</button>` : ""}
      </div>
    </div>
    <div class="history-card bijiris-template-card">
      <div class="survey-editor-head">
        <div>
          <div class="card-title">テンプレート投稿</div>
          <div class="meta">よく使う構成を選んで本文を下書きできます。</div>
        </div>
      </div>
      <div class="survey-question-grid bijiris-template-grid">
        <label>
          テンプレート
          <select id="bijirisTemplateSelect">
            <option value="">選択してください</option>
            ${BIJIRIS_POST_TEMPLATES.map((template) => `
              <option value="${escapeHtml(template.id)}">${escapeHtml(template.label)}</option>
            `).join("")}
          </select>
        </label>
        <button id="applyBijirisTemplateButton" class="secondary-button" type="button">テンプレートを反映</button>
      </div>
    </div>
    <form id="bijirisEditorForm" class="survey-editor-form" data-post-id="${escapeHtml(draft.id || "")}">
      <label>
        タイトル
        <input name="title" type="text" value="${escapeHtml(draft.title || "")}" required />
      </label>
      <div class="survey-question-grid">
        <label>
          カテゴリ
          <select name="category">
            ${BIJIRIS_POST_CATEGORY_OPTIONS.map((option) => `<option value="${escapeHtml(option)}" ${draft.category === option ? "selected" : ""}>${escapeHtml(option)}</option>`).join("")}
          </select>
        </label>
        <label>
          公開状態
          <select name="status">
            <option value="published" ${draft.status === "published" ? "selected" : ""}>公開</option>
            <option value="draft" ${draft.status === "draft" ? "selected" : ""}>下書き</option>
            <option value="archived" ${draft.status === "archived" ? "selected" : ""}>アーカイブ</option>
          </select>
          <div class="meta">顧客アプリに表示されるのは「公開」の投稿だけです。</div>
        </label>
      </div>
      <label class="inline-toggle">
        <input name="pinned" type="checkbox" ${draft.pinned ? "checked" : ""} />
        重要として一覧の上部に固定表示
      </label>
      <article class="history-card bijiris-notification-card">
        <div class="survey-editor-head">
          <div>
            <div class="card-title">通知設定</div>
            <div class="meta">公開中の投稿に対して、保存時に顧客アプリへ通知を送るかを指定できます。</div>
          </div>
        </div>
        <label class="inline-toggle">
          <input name="notifyCustomers" type="checkbox" ${draft.notifyCustomers ? "checked" : ""} />
          この保存内容を顧客アプリへ通知する
        </label>
        <div class="survey-question-grid">
          <label>
            通知タイトル
            <input
              name="notificationTitle"
              type="text"
              value="${escapeHtml(draft.notificationTitle || "")}"
              placeholder="${escapeHtml(selectedPost ? "豆知識を更新しました" : "新しい豆知識を追加しました")}"
            />
          </label>
          <label>
            通知本文
            <input
              name="notificationBody"
              type="text"
              value="${escapeHtml(draft.notificationBody || "")}"
              placeholder="例: 最新の豆知識を確認してください。"
            />
          </label>
        </div>
        <div class="meta">下書きやアーカイブでは通知されません。通知送信には OneSignal 設定が必要です。</div>
      </article>
      <label>
        要約
        <textarea name="summary" placeholder="一覧に表示する短い説明">${escapeHtml(draft.summary || "")}</textarea>
      </label>
      <label>
        本文
        <textarea name="body" placeholder="豆知識やアドバイス本文">${escapeHtml(draft.body || "")}</textarea>
      </label>
      <div class="survey-editor-head">
        <div class="card-title">写真</div>
        <label class="secondary-button upload-button">
          写真を追加
          <input id="bijirisPhotoInput" type="file" accept="image/*" multiple hidden />
        </label>
      </div>
      ${renderBijirisAttachmentPreview(draft.photos, "photo")}
      <div class="survey-editor-head">
        <div class="card-title">PDF資料</div>
        <label class="secondary-button upload-button">
          PDFを追加
          <input id="bijirisDocumentInput" type="file" accept="application/pdf,.pdf" multiple hidden />
        </label>
      </div>
      ${renderBijirisAttachmentPreview(draft.documents, "pdf")}
      <div class="survey-editor-actions">
        <button class="primary-button" type="submit">
          ${draft.status === "published"
            ? (selectedPost ? "公開内容を保存" : "作成して公開")
            : draft.status === "draft"
              ? (selectedPost ? "下書きを保存" : "下書きを作成")
              : (selectedPost ? "保存" : "作成")}
        </button>
      </div>
    </form>
    <article class="history-card bijiris-preview-wrapper">
      <div class="survey-editor-head">
        <div>
          <div class="card-title">公開プレビュー</div>
          <div class="meta">顧客アプリでの見え方を確認できます。</div>
        </div>
      </div>
      <div class="bijiris-preview-grid">
        <section class="bijiris-preview-column">
          <strong>一覧プレビュー</strong>
          <div id="bijirisPreviewList">
            ${renderAdminBijirisPreviewCard(draft)}
          </div>
        </section>
        <section class="bijiris-preview-column">
          <strong>詳細プレビュー</strong>
          <div id="bijirisPreviewDetail">
            ${renderAdminBijirisPreviewDetail(draft)}
          </div>
        </section>
      </div>
    </article>
  `;

  const form = document.querySelector("#bijirisEditorForm");
  document.querySelector("[data-bijiris-cancel-editor]")?.addEventListener("click", () => {
    state.selectedBijirisView = selectedPost ? "history" : "list";
    if (!selectedPost) state.bijirisEditorDraft = createEmptyBijirisPostDraft();
    renderBijirisManager();
    setPage("bijiris");
  });
  document.querySelector("#applyBijirisTemplateButton")?.addEventListener("click", () => {
    const templateId = document.querySelector("#bijirisTemplateSelect")?.value || "";
    if (!templateId) {
      showToast("テンプレートを選択してください。");
      return;
    }
    try {
      applyBijirisTemplate(templateId);
      renderBijirisManager();
      setPage("bijiris");
      showToast("テンプレートを反映しました。");
    } catch (error) {
      showToast(error.message || "テンプレートを反映できませんでした。");
    }
  });
  form?.addEventListener("input", () => {
    syncBijirisDraftFromForm(form);
    refreshBijirisPreview();
  });
  form?.addEventListener("change", () => {
    syncBijirisDraftFromForm(form);
    refreshBijirisPreview();
  });
  form?.addEventListener("submit", (event) => {
    event.preventDefault();
    syncBijirisDraftFromForm(form);
    saveBijirisPost().catch((error) => {
      showToast(error.message || "豆知識を保存できませんでした。");
    });
  });
  document.querySelector("#bijirisPhotoInput")?.addEventListener("change", (event) => {
    handleBijirisImageInput(event.currentTarget).catch((error) => {
      showToast(error.message || "写真を追加できませんでした。");
    });
  });
  document.querySelector("#bijirisDocumentInput")?.addEventListener("change", (event) => {
    handleBijirisDocumentInput(event.currentTarget).catch((error) => {
      showToast(error.message || "PDF を追加できませんでした。");
    });
  });
  document.querySelectorAll("[data-remove-bijiris-photo]").forEach((button) => {
    button.addEventListener("click", () => {
      syncBijirisDraftFromForm(form);
      state.bijirisEditorDraft.photos.splice(Number(button.dataset.removeBijirisPhoto), 1);
      renderBijirisManager();
      setPage("bijiris");
    });
  });
  document.querySelectorAll("[data-remove-bijiris-document]").forEach((button) => {
    button.addEventListener("click", () => {
      syncBijirisDraftFromForm(form);
      state.bijirisEditorDraft.documents.splice(Number(button.dataset.removeBijirisDocument), 1);
      renderBijirisManager();
      setPage("bijiris");
    });
  });
  document.querySelector("[data-delete-bijiris-post]")?.addEventListener("click", () => {
    deleteBijirisPost(state.selectedBijirisPostId).catch((error) => {
      showToast(error.message || "豆知識を削除できませんでした。");
    });
  });
}

function renderSurveyManager() {
  const surveyList = document.querySelector("#surveyManagerList");
  const surveyEditorCard = document.querySelector("#surveyEditorCard");
  const createButton = document.querySelector("#createSurveyButton");
  if (!surveyList || !surveyEditorCard || !createButton) return;

  if (state.selectedSurveyEditorId && !state.surveys.some((survey) => survey.id === state.selectedSurveyEditorId)) {
    state.selectedSurveyEditorId = "";
  }

  const selectedSurvey = findSurveyById(state.selectedSurveyEditorId);
  const draft = getSurveyEditorDraft();

  surveyList.innerHTML = state.surveys.length
    ? state.surveys
        .map((survey, index) => `
            <article class="survey-manager-card selectable-card ${survey.id === state.selectedSurveyEditorId ? "active" : ""}">
              <button
                class="survey-open-button"
                type="button"
                data-edit-survey-id="${escapeHtml(survey.id)}"
              >
              <div class="survey-manager-card-head">
                <strong>${index + 1}. ${escapeHtml(survey.title)}</strong>
                <span class="badge ${normalizeSurveyStatus(survey.status)}">${formatSurveyStatusLabel(survey.status)}</span>
              </div>
              <div>${survey.questions.length}問</div>
              <div class="meta">更新: ${survey.updatedAt ? formatDate(survey.updatedAt) : "-"}</div>
              <div class="meta">${survey.acceptingResponses === false ? "受付停止中" : "受付中"}</div>
              </button>
              <div class="survey-card-actions">
                <button class="secondary-button" type="button" data-move-survey="${survey.id}" data-direction="-1">↑</button>
                <button class="secondary-button" type="button" data-move-survey="${survey.id}" data-direction="1">↓</button>
                <button class="secondary-button" type="button" data-duplicate-survey="${survey.id}">複製</button>
                <button class="secondary-button" type="button" data-archive-survey="${survey.id}">アーカイブ</button>
              </div>
            </article>
          `)
        .join("")
    : `<div class="empty">アンケートはまだありません。</div>`;

  surveyEditorCard.innerHTML = `
    <div class="survey-editor-head">
      <div>
        <div class="card-title">${selectedSurvey ? "アンケート編集" : "アンケート新規作成"}</div>
        <div class="meta">
          ${selectedSurvey
            ? `作成: ${draft.createdAt ? formatDate(draft.createdAt) : "-"} / 更新: ${draft.updatedAt ? formatDate(draft.updatedAt) : "-"}`
            : "タイトル、説明、質問項目を入力してください。"}
        </div>
      </div>
      ${
        selectedSurvey
          ? `
              <div class="action-row">
                <button class="secondary-button" type="button" data-duplicate-survey="${escapeHtml(selectedSurvey.id)}">複製</button>
                <button class="secondary-button" type="button" data-archive-survey="${escapeHtml(selectedSurvey.id)}">アーカイブ</button>
                <button class="secondary-button danger-button" type="button" data-delete-survey="${escapeHtml(selectedSurvey.id)}">削除</button>
              </div>
            `
          : ""
      }
    </div>
    <form id="surveyEditorForm" class="survey-editor-form" data-survey-id="${escapeHtml(selectedSurvey?.id || "")}">
      <label>
        タイトル
        <input name="title" type="text" value="${escapeHtml(draft.title || "")}" required />
      </label>
      <label>
        説明
        <textarea name="description" required>${escapeHtml(draft.description || "")}</textarea>
      </label>
      <label>
        回答前メッセージ
        <textarea name="introMessage">${escapeHtml(draft.introMessage || "")}</textarea>
      </label>
      <label>
        回答完了メッセージ
        <textarea name="completionMessage">${escapeHtml(draft.completionMessage || "")}</textarea>
      </label>
      <label>
        公開状態
        <select name="status">
          <option value="published" ${normalizeSurveyStatus(draft.status) === "published" ? "selected" : ""}>公開</option>
          <option value="draft" ${normalizeSurveyStatus(draft.status) === "draft" ? "selected" : ""}>下書き</option>
          <option value="archived" ${normalizeSurveyStatus(draft.status) === "archived" ? "selected" : ""}>アーカイブ</option>
        </select>
      </label>
      <label class="inline-toggle">
        <input name="acceptingResponses" type="checkbox" ${draft.acceptingResponses === false ? "" : "checked"} />
        回答受付中
      </label>
      <div class="survey-question-grid">
        <label>
          受付開始日時
          <input name="startAt" type="datetime-local" value="${escapeHtml(toDateTimeLocalValue(draft.startAt))}" />
        </label>
        <label>
          受付終了日時
          <input name="endAt" type="datetime-local" value="${escapeHtml(toDateTimeLocalValue(draft.endAt))}" />
        </label>
      </div>
      <div class="survey-editor-head">
        <div class="card-title">質問項目</div>
        <button class="secondary-button" id="addSurveyQuestionButton" type="button">質問を追加</button>
      </div>
      <div id="surveyQuestions" class="survey-question-list">
        ${draft.questions.map((question, index) => renderSurveyQuestionEditor(question, index, draft.questions)).join("")}
      </div>
      <div class="survey-editor-actions">
        <button class="primary-button" type="submit">${selectedSurvey ? "保存" : "作成"}</button>
      </div>
    </form>
  `;

  createButton.onclick = () => {
    state.selectedSurveyEditorId = "";
    renderSurveyManager();
    setPage("surveys");
  };

  surveyList.querySelectorAll("[data-edit-survey-id]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedSurveyEditorId = button.dataset.editSurveyId;
      renderSurveyManager();
      setPage("surveys");
    });
  });

  surveyList.querySelectorAll("[data-move-survey]").forEach((button) => {
    button.addEventListener("click", (event) => {
      event.stopPropagation();
      void moveSurvey(button.dataset.moveSurvey, Number(button.dataset.direction));
    });
  });

  document.querySelectorAll("[data-duplicate-survey]").forEach((button) => {
    button.addEventListener("click", (event) => {
      event.stopPropagation();
      void duplicateSurvey(button.dataset.duplicateSurvey);
    });
  });

  document.querySelectorAll("[data-archive-survey]").forEach((button) => {
    button.addEventListener("click", (event) => {
      event.stopPropagation();
      void archiveSurvey(button.dataset.archiveSurvey);
    });
  });

  const form = surveyEditorCard.querySelector("#surveyEditorForm");
  const questionsContainer = surveyEditorCard.querySelector("#surveyQuestions");
  surveyEditorCard.querySelector("#addSurveyQuestionButton")?.addEventListener("click", () => {
    addSurveyQuestionEditor(questionsContainer, makeBlankQuestion());
  });

  questionsContainer?.addEventListener("click", (event) => {
    const removeQuestionButton = event.target.closest("[data-remove-question]");
    if (removeQuestionButton) {
      const blocks = questionsContainer.querySelectorAll("[data-question-block]");
      if (blocks.length <= 1) {
        showToast("質問は1つ以上必要です。");
        return;
      }
      removeQuestionButton.closest("[data-question-block]")?.remove();
      syncSurveyQuestionNumbers(questionsContainer);
      refreshVisibilityConditionEditors(questionsContainer);
      return;
    }

    const addVisibilityConditionButton = event.target.closest("[data-add-visibility-condition]");
    if (addVisibilityConditionButton) {
      addVisibilityConditionEditor(addVisibilityConditionButton.closest("[data-question-block]"));
      return;
    }

    const removeVisibilityConditionButton = event.target.closest("[data-remove-visibility-condition]");
    if (removeVisibilityConditionButton) {
      removeVisibilityConditionButton.closest("[data-visibility-condition-row]")?.remove();
      return;
    }

    const addOptionButton = event.target.closest("[data-add-option]");
    if (addOptionButton) {
      const list = addOptionButton
        .closest("[data-options-wrapper]")
        ?.querySelector("[data-option-list]");
      list?.insertAdjacentHTML(
        "beforeend",
        `
          <div class="option-row">
            <input type="text" data-option-input value="" placeholder="選択肢" />
            <button class="secondary-button" type="button" data-remove-option>削除</button>
          </div>
        `,
      );
      return;
    }

    const removeOptionButton = event.target.closest("[data-remove-option]");
    if (removeOptionButton) {
      const list = removeOptionButton.closest("[data-option-list]");
      if (list && list.querySelectorAll(".option-row").length <= 2) {
        showToast("選択肢は2つ以上必要です。");
        return;
      }
      removeOptionButton.closest(".option-row")?.remove();
      refreshVisibilityConditionEditors(questionsContainer);
    }
  });

  questionsContainer?.addEventListener("change", (event) => {
    const typeSelect = event.target.closest("[data-question-type]");
    if (typeSelect) {
      const block = typeSelect.closest("[data-question-block]");
      if (block) refreshQuestionOptionsEditor(block);
      refreshVisibilityConditionEditors(questionsContainer);
      return;
    }

    if (
      event.target.closest("[data-condition-question-id]") ||
      event.target.closest("[data-option-input]") ||
      event.target.closest("[data-category-options-input]") ||
      event.target.closest("[data-question-label]")
    ) {
      refreshVisibilityConditionEditors(questionsContainer);
    }
  });

  questionsContainer?.addEventListener("input", (event) => {
    if (
      event.target.closest("[data-option-input]") ||
      event.target.closest("[data-category-options-input]") ||
      event.target.closest("[data-question-label]")
    ) {
      refreshVisibilityConditionEditors(questionsContainer);
    }
  });

  form?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const submitButton = form.querySelector('button[type="submit"]');
    const surveyId = form.dataset.surveyId || "";
    submitButton.disabled = true;
    submitButton.textContent = surveyId ? "保存中" : "作成中";

    try {
      const payload = buildSurveyPayload(form);
      const result = surveyId
        ? await api.request(`/api/admin/surveys/${encodeURIComponent(surveyId)}`, {
            method: "PUT",
            token: state.token,
            body: payload,
          })
        : await api.request("/api/admin/surveys", {
            method: "POST",
            token: state.token,
            body: payload,
          });
      const survey = result.survey;
      state.surveys = surveyId
        ? state.surveys.map((item) => (item.id === survey.id ? survey : item))
        : [survey, ...state.surveys];
      state.selectedSurveyEditorId = survey.id;
      renderAll();
      setPage("surveys");
      showToast(surveyId ? "アンケートを保存しました。" : "アンケートを作成しました。");
    } catch (error) {
      showToast(error.message || "アンケートを保存できませんでした。");
    } finally {
      submitButton.disabled = false;
      submitButton.textContent = surveyId ? "保存" : "作成";
    }
  });

  surveyEditorCard.querySelector("[data-delete-survey]")?.addEventListener("click", async () => {
    if (!selectedSurvey) return;
    if (!confirm(`「${selectedSurvey.title}」を削除しますか？`)) return;
    try {
      await api.request(`/api/admin/surveys/${encodeURIComponent(selectedSurvey.id)}`, {
        method: "DELETE",
        token: state.token,
      });
      state.surveys = state.surveys.filter((survey) => survey.id !== selectedSurvey.id);
      state.selectedSurveyEditorId = "";
      renderAll();
      setPage("surveys");
      showToast("アンケートを削除しました。");
    } catch (error) {
      showToast(error.message || "アンケートを削除できませんでした。");
    }
  });
}

function exportCsv() {
  const filteredResponses = getFilteredResponses();
  if (!filteredResponses.length) {
    showToast("出力対象の回答がありません。");
    return;
  }
  const rows = [
    ["日時", "会員番号", "お客様", "アンケート", "回数券", "何枚目", "何回目", "対応状況", "管理メモ", "写真枚数", "回答"],
    ...filteredResponses.map((response) => [
      formatDate(response.submittedAt),
      getCustomerMemberNumber(response.customerName) || response.customerMemberNumber || "",
      response.customerName,
      response.surveyTitle,
      getResponseTicketInfo(response).find((item) => item.label === "回数券")?.value || "",
      getResponseTicketInfo(response).find((item) => item.label === "何枚目")?.value || "",
      getResponseTicketInfo(response).find((item) => item.label === "何回目")?.value || "",
      STATUS_LABELS[normalizeStatus(response.status)],
      response.adminMemo || "",
      String(collectPhotosFromResponse(response).length),
      response.answers.map(formatAnswerForCsv).join(" / "),
    ]),
  ];
  const csv = rows
    .map((row) => row.map((cell) => `"${String(cell).replaceAll('"', '""')}"`).join(","))
    .join("\n");
  const blob = new Blob([`\uFEFF${csv}`], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `survey-responses-${Date.now()}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function exportResponsesPdf() {
  const filteredResponses = getFilteredResponses();
  if (!filteredResponses.length) {
    showToast("出力対象の回答がありません。");
    return;
  }
  const printWindow = window.open("", "_blank", "noopener,noreferrer");
  if (!printWindow) {
    showToast("PDF出力画面を開けませんでした。");
    return;
  }
  const rows = filteredResponses
    .map(
      (response) => `
        <section class="entry">
          <h2>${escapeHtml(response.surveyTitle)}</h2>
          <div class="meta">${escapeHtml(formatDate(response.submittedAt))} / ${escapeHtml(getCustomerNameWithMember(response.customerName))}</div>
          <div class="meta">対応状況: ${escapeHtml(STATUS_LABELS[normalizeStatus(response.status)])}</div>
          <div class="meta">回数券: ${escapeHtml(getResponseTicketInfo(response).map((item) => `${item.label} ${item.value}`).join(" / ") || "-")}</div>
          <div class="meta">写真枚数: ${collectPhotosFromResponse(response).length}</div>
          ${response.adminMemo ? `<div class="item"><strong>管理メモ</strong><div>${escapeHtml(response.adminMemo)}</div></div>` : ""}
          <div class="item"><strong>回答</strong><div>${escapeHtml(response.answers.map(formatAnswerForCsv).join(" / "))}</div></div>
        </section>
      `,
    )
    .join("");
  printWindow.document.write(`
    <!doctype html>
    <html lang="ja">
      <head>
        <meta charset="UTF-8" />
        <title>回答一覧 PDF</title>
        <style>
          body { font-family: sans-serif; padding: 24px; color: #222; }
          h1 { font-size: 22px; margin-bottom: 8px; }
          .meta { color: #666; margin-bottom: 8px; }
          .entry { border: 1px solid #ddd; border-radius: 8px; padding: 12px; margin-bottom: 12px; }
          .entry h2 { font-size: 16px; margin: 0 0 6px; }
          .item { margin-top: 8px; }
          .item strong { display: block; margin-bottom: 4px; }
        </style>
      </head>
      <body>
        <h1>回答一覧</h1>
        <div class="meta">件数: ${filteredResponses.length}件</div>
        ${rows}
      </body>
    </html>
  `);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
}

function exportBackup() {
  const payload = {
    exportedAt: new Date().toISOString(),
    surveys: state.surveys,
    preferences: state.preferences,
    adminUsers: state.adminUsers,
    customerMemos: state.customerMemos,
    customerProfiles: Object.values(state.customerProfiles || {}),
    measurements: state.measurements,
    bijirisPosts: state.bijirisPosts,
    responses: state.responses,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: "application/json;charset=utf-8",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `mayumi-survey-backup-${Date.now()}.json`;
  link.click();
  URL.revokeObjectURL(url);
}

async function restoreBackup(file) {
  if (!file) return;
  const text = await file.text();
  const payload = JSON.parse(text);
  if (!Array.isArray(payload.surveys) || !payload.surveys.length) {
    throw new Error("復元ファイルにアンケート情報がありません。");
  }
  const surveysResult = await api.request("/api/admin/surveys/replace", {
    method: "POST",
    token: state.token,
    body: { surveys: payload.surveys },
  });
  state.surveys = surveysResult.surveys || state.surveys;
  if (payload.preferences) {
    const preferencesResult = await api.request("/api/admin/preferences", {
      method: "PUT",
      token: state.token,
      body: payload.preferences,
    });
    state.preferences = preferencesResult.preferences || state.preferences;
  }
  if (payload.customerMemos && typeof payload.customerMemos === "object") {
    const entries = Object.entries(payload.customerMemos);
    for (const [customerName, record] of entries) {
      const latestMemo = normalizeMemoRecord(record).latestMemo;
      await api.request(`/api/admin/customer-memos/${encodeURIComponent(customerName)}`, {
        method: "PUT",
        token: state.token,
        body: { memo: latestMemo },
      });
    }
    state.customerMemos = payload.customerMemos;
  }
  if (Array.isArray(payload.customerProfiles) && payload.customerProfiles.length) {
    for (const profile of payload.customerProfiles) {
      const ticketCard = profile?.activeTicketCard || null;
      await api.request(`/api/admin/customers/${encodeURIComponent(profile.name)}`, {
        method: "PUT",
        token: state.token,
        body: {
          name: String(profile?.name || "").trim(),
          memberNumber: String(profile?.memberNumber || "").trim(),
          ticketPlan: String(ticketCard?.plan || "").trim(),
          ticketSheet: ticketCard?.sheetNumber ? `${ticketCard.sheetNumber}枚目` : "",
          ticketRound: ticketCard?.round ? `${ticketCard.round}回目` : "",
          measurementTargets: profile?.measurementTargets || {},
        },
      });
    }
  }
  if (Array.isArray(payload.measurements)) {
    await api.request("/api/admin/measurements/replace", {
      method: "POST",
      token: state.token,
      body: { measurements: payload.measurements },
    });
  }
  if (Array.isArray(payload.bijirisPosts)) {
    await api.request("/api/admin/bijiris-posts/replace", {
      method: "POST",
      token: state.token,
      body: { posts: payload.bijirisPosts },
    });
  }
  if (Array.isArray(payload.adminUsers) && payload.adminUsers.length) {
    await api.request("/api/admin/users", {
      method: "PUT",
      token: state.token,
      body: { adminUsers: payload.adminUsers },
    });
    state.adminUsers = payload.adminUsers;
  }
  await loadAdminData();
  showToast("バックアップを復元しました。");
  renderAll();
}

async function savePreferences() {
  const form = document.querySelector("#preferencesForm");
  if (!form) return;
  const formData = new FormData(form);
  const payload = {
    notificationEnabled: formData.get("notificationEnabled") === "on",
    notificationEmail: String(formData.get("notificationEmail") || "").trim(),
    notificationSubject: String(formData.get("notificationSubject") || "").trim(),
    notificationBody: String(formData.get("notificationBody") || "").trim(),
    dataPolicyText: String(formData.get("dataPolicyText") || "").trim(),
    requireConsent: formData.get("requireConsent") === "on",
    consentText: String(formData.get("consentText") || "").trim(),
    autoBackupEnabled: formData.get("autoBackupEnabled") === "on",
    backupHour: Number(formData.get("backupHour") || 0),
    retentionDays: Number(formData.get("retentionDays") || 0),
    recoveryMemo: String(formData.get("recoveryMemo") || "").trim(),
    twoFactorEnabled: false,
  };
  try {
    const result = await api.request("/api/admin/preferences", {
      method: "PUT",
      token: state.token,
      body: payload,
    });
    state.preferences = result.preferences || payload;
    showToast("設定を保存しました。");
    renderSettings();
  } catch (error) {
    showToast(error.message || "設定を保存できませんでした。");
  }
}

async function savePushConfig(form) {
  const formData = new FormData(form);
  const button = form.querySelector('button[type="submit"]');
  if (button) {
    button.disabled = true;
    button.textContent = "保存中";
  }
  try {
    const result = await api.request("/api/admin/push-config", {
      method: "PUT",
      token: state.token,
      body: {
        pushAppId: String(formData.get("pushAppId") || "").trim(),
        customerAppUrl: String(formData.get("customerAppUrl") || "").trim(),
        restApiKey: String(formData.get("restApiKey") || "").trim(),
      },
    });
    state.adminInfo = result.adminInfo || state.adminInfo;
    state.customerProfiles = indexCustomerProfiles(state.adminInfo?.customerProfiles);
    state.adminUsers = Array.isArray(state.adminInfo?.adminUsers) ? state.adminInfo.adminUsers : state.adminUsers;
    renderSettings();
    showToast("通知設定を保存しました。");
  } catch (error) {
    showToast(error.message || "通知設定を保存できませんでした。");
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = "通知設定を保存";
    }
  }
}

async function saveAdminUsers(form) {
  const items = Array.from(form.querySelectorAll(".admin-user-item")).map((item) => ({
    id: item.dataset.adminUserId || "",
    username: String(item.querySelector('[name="username"]')?.value || "").trim(),
    password: String(item.querySelector('[name="password"]')?.value || "").trim(),
    email: String(item.querySelector('[name="email"]')?.value || "").trim(),
    active: Boolean(item.querySelector('[name="active"]')?.checked),
  }));
  try {
    const result = await api.request("/api/admin/users", {
      method: "PUT",
      token: state.token,
      body: { adminUsers: items },
    });
    state.adminUsers = result.adminUsers || state.adminUsers;
    state.adminInfo = state.adminInfo ? { ...state.adminInfo, adminUsers: state.adminUsers } : state.adminInfo;
    showToast("管理者アカウントを保存しました。");
    renderSettings();
  } catch (error) {
    showToast(error.message || "管理者アカウントを保存できませんでした。");
  }
}

async function runMaintenanceNow() {
  try {
    await api.request("/api/admin/maintenance/run", {
      method: "POST",
      token: state.token,
    });
    await loadAdminData();
    showToast("メンテナンスを実行しました。");
  } catch (error) {
    showToast(error.message || "メンテナンスを実行できませんでした。");
  }
}

function openLightbox(src, title) {
  const modal = document.querySelector("#lightboxModal");
  if (!modal || !src) return;
  modal.querySelector("img").src = src;
  modal.querySelector("img").alt = title || "画像";
  modal.querySelector(".lightbox-caption").textContent = title || "";
  modal.hidden = false;
}

function closeLightbox() {
  const modal = document.querySelector("#lightboxModal");
  if (!modal) return;
  modal.hidden = true;
  modal.querySelector("img").src = "";
}

function attachLightboxHandlers(root = document) {
  root.querySelectorAll("[data-lightbox-src]").forEach((button) => {
    button.addEventListener("click", () => {
      openLightbox(button.dataset.lightboxSrc, button.dataset.lightboxTitle);
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
          .filter((key) => key.startsWith("mayumi-admin-survey-"))
          .map((key) => caches.delete(key)),
      );
    }
    showToast("アプリを更新しています。");
    const url = new URL(window.location.href);
    url.searchParams.set("appRefresh", String(Date.now()));
    window.location.replace(url.toString());
  } catch (error) {
    reportClientError("admin.appUpdate", error);
    showToast(error.message || "アプリを更新できませんでした。");
  } finally {
    if (appUpdateButton) {
      appUpdateButton.disabled = false;
      appUpdateButton.textContent = "アプリ更新";
    }
  }
}

function setupInstall() {
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker
        .register("./sw.js?v=20260415-07", { updateViaCache: "none" })
        .then((registration) => registration.update().catch(() => {}))
        .catch(() => {});
    });
  }

  window.addEventListener("beforeinstallprompt", (event) => {
    event.preventDefault();
    state.installPrompt = event;
    installButton.hidden = false;
  });

  installButton.addEventListener("click", async () => {
    if (!state.installPrompt) return;
    state.installPrompt.prompt();
    await state.installPrompt.userChoice;
    state.installPrompt = null;
    installButton.hidden = true;
  });
}

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(event.currentTarget);
  const button = event.currentTarget.querySelector('button[type="submit"]');
  button.disabled = true;
  button.textContent = "確認中";
  try {
    const result = await api.request("/api/admin/login", {
      method: "POST",
      body: {
        loginId: String(formData.get("loginId") || ""),
        password: String(formData.get("password") || ""),
      },
    });
    state.token = result.token;
    localStorage.setItem(TOKEN_KEY, state.token);
    setLoggedIn(true);
    await loadAdminData();
    showToast("ログインしました。");
  } catch (error) {
    showToast(error.message || "ログインできませんでした。");
  } finally {
    button.disabled = false;
    button.textContent = "ログイン";
  }
});

credentialForm.addEventListener("input", () => {
  credentialForm.dataset.dirty = "true";
});

credentialForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(event.currentTarget);
  const button = event.currentTarget.querySelector('button[type="submit"]');
  button.disabled = true;
  button.textContent = "保存中";
  try {
    const result = await api.request("/api/admin/credentials", {
      method: "PUT",
      token: state.token,
      body: {
        loginId: String(formData.get("loginId") || ""),
        password: String(formData.get("password") || ""),
      },
    });
    state.adminInfo = result.adminInfo || state.adminInfo;
    state.customerProfiles = indexCustomerProfiles(state.adminInfo?.customerProfiles);
    state.adminUsers = Array.isArray(state.adminInfo?.adminUsers) ? state.adminInfo.adminUsers : state.adminUsers;
    credentialForm.dataset.dirty = "";
    credentialForm.reset();
    renderSettings();
    showToast("管理者認証を更新しました。");
  } catch (error) {
    showToast(error.message || "認証情報を更新できませんでした。");
  } finally {
    button.disabled = false;
    button.textContent = "保存";
  }
});

document.querySelector("#logoutButton").addEventListener("click", () => {
  state.token = "";
  state.adminInfo = null;
  state.customerProfiles = {};
  localStorage.removeItem(TOKEN_KEY);
  void api.logout?.();
  setLoggedIn(false);
});

document.querySelector("#refreshButton").addEventListener("click", () => {
  void loadAdminData();
});
appUpdateButton?.addEventListener("click", () => {
  void runAppUpdate();
});

document.querySelector("#exportCsvButton").addEventListener("click", exportCsv);
document.querySelector("#exportPdfButton")?.addEventListener("click", exportResponsesPdf);
document.querySelector("#backupExportButton")?.addEventListener("click", exportBackup);
document.querySelector("#runMaintenanceButton")?.addEventListener("click", () => {
  void runMaintenanceNow();
});
document.querySelector("#backupImportInput")?.addEventListener("change", (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  restoreBackup(file).catch((error) => {
    showToast(error.message || "バックアップを復元できませんでした。");
  });
  event.target.value = "";
});

["surveyFilter", "statusFilter", "keywordFilter", "memberNumberFilter", "nameKanaFilter", "memoPresenceFilter", "concernCategoryFilter", "ticketPlanFilter", "ticketSheetFilter", "ticketRoundFilter", "photoFilter", "photoCountFilter", "dateFromFilter", "dateToFilter"].forEach((id) => {
  document.querySelector(`#${id}`).addEventListener("input", renderResponses);
  document.querySelector(`#${id}`).addEventListener("change", renderResponses);
});
["customerKeywordFilter", "customerTicketPlanFilter", "customerResponseStateFilter"].forEach((id) => {
  document.querySelector(`#${id}`)?.addEventListener("input", renderCustomerManagement);
  document.querySelector(`#${id}`)?.addEventListener("change", renderCustomerManagement);
});
document.querySelector("#onlyUnreadButton")?.addEventListener("click", () => {
  document.querySelector("#statusFilter").value = "new";
  setPage("responses");
  renderResponses();
});
["auditLogFilter", "errorLogFilter"].forEach((id) => {
  document.querySelector(`#${id}`)?.addEventListener("input", renderSettings);
});

document.querySelectorAll("[data-page]").forEach((button) => {
  button.addEventListener("click", () => setPage(button.dataset.page));
});

setupInstall();
window.addEventListener("error", (event) => {
  reportClientError("admin.window", event.error || event.message || "window error");
});
window.addEventListener("unhandledrejection", (event) => {
  reportClientError("admin.promise", event.reason || "unhandled rejection");
});
document.querySelector("#lightboxModal")?.addEventListener("click", (event) => {
  if (event.target.id === "lightboxModal" || event.target.closest("[data-close-lightbox]")) {
    closeLightbox();
  }
});
setLoggedIn(Boolean(state.token));
if (state.token) {
  loadAdminData().catch((error) => {
    localStorage.removeItem(TOKEN_KEY);
    state.token = "";
    setLoggedIn(false);
    showToast(error.message || "再ログインしてください。");
  });
}
