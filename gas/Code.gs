var SPREADSHEET_ID = "1oDNTqlvKv1rGOGXIpnzPlegpFDeQ0WHGRLuY3ZAnZYc";
var DEFAULT_SPREADSHEET_TITLE = "まゆみ助産院 ビジリス アンケート回答";
var MASTER_SHEET_NAME = "回答一覧";
var ROOT_DRIVE_FOLDER_NAME = "Bijiris";
var LEGACY_ROOT_DRIVE_FOLDER_NAMES = ["bijiris"];
var BIJIRIS_POSTS_FOLDER_NAME = "ビジリス通信";
var DEFAULT_CUSTOMER_APP_URL = "https://mayumi-josanin.github.io/mayumi_bijiris/customer-app/";
var DEFAULT_ONESIGNAL_APP_ID = "88023099-c99e-44c6-9f7c-2ef08d363768";
var ONE_SIGNAL_APP_SCOPE_KEY = "app_scope";
var ONE_SIGNAL_APP_SCOPE_VALUE = "mayumi_bijiris";
var DEFAULT_ADMIN_USERNAME = "mayumi2026";
var DEFAULT_ADMIN_PASSWORD = "3939";
var LEGACY_ADMIN_USERNAME = "admin";
var LEGACY_ADMIN_PASSWORD = "admin123";
var DEFAULT_TOKEN_SECRET = "change-this-gas-secret";
var SURVEYS_PROPERTY_KEY = "SURVEYS_JSON";
var QUESTION_TYPES = ["text", "textarea", "rating", "choice", "checkbox", "photo"];
var CUSTOMER_TICKET_INFO_QUESTION_IDS = {
  plan: ["q_bijiris_session_ticket_plan", "q_ticket_end_ticket_size"],
  sheet: ["q_bijiris_session_ticket_sheet", "q_ticket_end_ticket_sheet"],
  round: ["q_bijiris_session_ticket_round", "q_ticket_end_ticket_round"],
};
var PREFERENCES_PROPERTY_KEY = "ADMIN_PREFERENCES_JSON";
var CUSTOMER_MEMOS_PROPERTY_KEY = "CUSTOMER_MEMOS_JSON";
var CUSTOMER_PROFILES_PROPERTY_KEY = "CUSTOMER_PROFILES_JSON";
var AUDIT_LOGS_PROPERTY_KEY = "AUDIT_LOGS_JSON";
var ERROR_LOGS_PROPERTY_KEY = "ERROR_LOGS_JSON";
var LOGIN_ATTEMPTS_PROPERTY_KEY = "LOGIN_ATTEMPTS_JSON";
var ADMIN_USERS_PROPERTY_KEY = "ADMIN_USERS_JSON";
var OTP_SESSIONS_PROPERTY_KEY = "OTP_SESSIONS_JSON";
var MAINTENANCE_TRIGGER_IDS_PROPERTY_KEY = "MAINTENANCE_TRIGGER_IDS_JSON";
var NEXT_MEMBER_NUMBER_PROPERTY_KEY = "NEXT_MEMBER_NUMBER";
var BACKUP_META_PROPERTY_KEY = "BACKUP_META_JSON";
var LAST_MAINTENANCE_META_PROPERTY_KEY = "LAST_MAINTENANCE_META_JSON";
var VERSION = "20260415-03";
var RESPONSE_EDIT_WINDOW_MS = 24 * 60 * 60 * 1000;
var RESPONSE_RESUBMIT_COOLDOWN_MS = 24 * 60 * 60 * 1000;
var TICKET_CARD_ACQUIRE_COOLDOWN_MS = 24 * 60 * 60 * 1000;
var DUPLICATE_RESPONSE_WINDOW_MS = 10 * 60 * 1000;
var LOGIN_LOCK_WINDOW_MS = 15 * 60 * 1000;
var LOGIN_MAX_ATTEMPTS = 5;
var MAX_LOG_ENTRIES = 200;
var OTP_TTL_MS = 10 * 60 * 1000;

var BIJIRIS_SESSION_CONCERN_CATEGORIES = [
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

var BIJIRIS_SESSION_TICKET_SHEET_OPTIONS = [
  "1枚目",
  "2枚目",
  "3枚目",
  "4枚目",
  "5枚目",
  "6枚目",
  "7枚目",
  "8枚目",
  "9枚目",
  "10枚目",
];

var BIJIRIS_SESSION_TICKET_ROUND_OPTIONS = [
  "1回目",
  "2回目",
  "3回目",
  "4回目",
  "5回目",
  "6回目",
  "7回目",
  "8回目",
  "9回目",
  "10回目",
];

var BIJIRIS_SESSION_LIFE_CHANGE_OPTIONS = [
  "咳やくしゃみをした時の不安が以前より減った",
  "急な尿意を気にする場面が減った",
  "外出時にトイレの場所を気にしすぎなくなった",
  "夜中にトイレで起きる回数が減った",
  "お腹の奥に力が入りやすくなった",
  "骨盤まわりが安定した感じがある",
  "姿勢を意識しやすくなった",
  "下腹まわりがすっきりした感じがある",
  "歩く・立つ・動くことが以前より楽になった",
  "その他（自由記述）",
];

var MASTER_HEADERS = [
  "送信日時",
  "回答ID",
  "アンケートID",
  "アンケート名",
  "端末ID",
  "お名前",
  "メールアドレス",
  "対応状況",
  "管理メモ",
  "回答JSON",
  "写真JSON",
  "管理更新日時",
];

var MEASUREMENTS_SHEET_NAME = "測定履歴";
var BIJIRIS_POSTS_SHEET_NAME = "ビジリス通信";
var BIJIRIS_POST_ATTACHMENTS_SHEET_NAME = "ビジリス通信添付";
var MEASUREMENT_HEADERS = [
  "作成日時",
  "更新日時",
  "測定ID",
  "顧客名",
  "会員番号",
  "測定日",
  "ウエスト(cm)",
  "ヒップ(cm)",
  "太もも右(cm)",
  "太もも左(cm)",
  "WHR",
  "スタッフメモ",
];
var BIJIRIS_POST_HEADERS = [
  "作成日時",
  "更新日時",
  "公開日時",
  "投稿ID",
  "タイトル",
  "カテゴリ",
  "要約",
  "本文",
  "公開状態",
  "固定表示",
];
var BIJIRIS_POST_ATTACHMENT_HEADERS = [
  "投稿ID",
  "種別",
  "表示順",
  "ファイル名",
  "MIMEタイプ",
  "ファイルID",
  "URL",
  "プレビューURL",
  "ダウンロードURL",
  "サムネイルURL",
  "表示タイトル",
  "サムネイルファイルID",
];

function getBijirisSessionConcernOptions_() {
  var options = [];
  BIJIRIS_SESSION_CONCERN_CATEGORIES.forEach(function (category) {
    (category.options || []).forEach(function (option) {
      options.push(option);
    });
  });
  options.push("その他（長文）");
  return options;
}

var SURVEYS = [
  {
    id: "survey_bijiris_session",
    title: "ビジリス施術アンケート",
    description: "ビジリス施術後の体感やお悩みをお聞かせください。",
    introMessage: "施術内容を選択後、本日の体感や気になることをご回答ください。",
    completionMessage: "ビジリス施術アンケートのご回答ありがとうございました。",
    status: "published",
    questions: [
      { id: "q_bijiris_session_type", label: "施術内容", type: "choice", required: true, options: ["初回お試し", "回数券", "単発", "トライアル"] },
      { id: "q_bijiris_session_ticket_plan", label: "回数券の種類", type: "choice", required: true, options: ["6回券", "10回券"], visibleWhen: { questionId: "q_bijiris_session_type", value: "回数券" } },
      { id: "q_bijiris_session_ticket_sheet", label: "回数券の何枚目ですか？", type: "choice", required: true, options: BIJIRIS_SESSION_TICKET_SHEET_OPTIONS, visibleWhen: { questionId: "q_bijiris_session_type", value: "回数券" } },
      { id: "q_bijiris_session_ticket_round", label: "回数券の何回目ですか？", type: "choice", required: true, options: BIJIRIS_SESSION_TICKET_ROUND_OPTIONS, visibleWhen: { questionId: "q_bijiris_session_type", value: "回数券" } },
      { id: "q_bijiris_session_feeling", label: "本日のビジリスの体感はいかがでしたか？　以前と比べて変化したことなどがあればご記載ください", type: "textarea", required: true, options: [] },
      { id: "q_bijiris_session_concern", label: "普段のお身体のお悩みや、ビジリス（骨盤底筋ケア）について気になること・知りたいことはありますか？（複数選択可）", type: "checkbox", required: false, options: getBijirisSessionConcernOptions_() },
      { id: "q_bijiris_session_concern_other", label: "気になること・知りたいこと（その他・長文）", type: "textarea", required: false, options: [], visibleWhen: { questionId: "q_bijiris_session_concern", value: "その他（長文）" } },
      { id: "q_bijiris_session_life_changes", label: "日常生活にどのような変化がありましたか？（複数選択可）", type: "checkbox", required: false, options: BIJIRIS_SESSION_LIFE_CHANGE_OPTIONS },
      { id: "q_bijiris_session_life_changes_other", label: "日常生活の変化（その他）", type: "textarea", required: false, options: [], visibleWhen: { questionId: "q_bijiris_session_life_changes", value: "その他（自由記述）" } },
      { id: "q_bijiris_session_monitor_photos_6", label: "モニター時の写真2枚", type: "photo", required: true, options: [], visibilityConditions: [{ questionId: "q_bijiris_session_type", value: "回数券" }, { questionId: "q_bijiris_session_ticket_plan", value: "6回券" }, { questionId: "q_bijiris_session_ticket_round", value: "6回目" }], visibleWhen: { questionId: "q_bijiris_session_type", value: "回数券" } },
      { id: "q_bijiris_session_ticket_end_photos_6", label: "回数券終了時の写真2枚", type: "photo", required: true, options: [], visibilityConditions: [{ questionId: "q_bijiris_session_type", value: "回数券" }, { questionId: "q_bijiris_session_ticket_plan", value: "6回券" }, { questionId: "q_bijiris_session_ticket_round", value: "6回目" }], visibleWhen: { questionId: "q_bijiris_session_type", value: "回数券" } },
      { id: "q_bijiris_session_monitor_photos_10", label: "モニター時の写真2枚", type: "photo", required: true, options: [], visibilityConditions: [{ questionId: "q_bijiris_session_type", value: "回数券" }, { questionId: "q_bijiris_session_ticket_plan", value: "10回券" }, { questionId: "q_bijiris_session_ticket_round", value: "10回目" }], visibleWhen: { questionId: "q_bijiris_session_type", value: "回数券" } },
      { id: "q_bijiris_session_ticket_end_photos_10", label: "回数券終了時の写真2枚", type: "photo", required: true, options: [], visibilityConditions: [{ questionId: "q_bijiris_session_type", value: "回数券" }, { questionId: "q_bijiris_session_ticket_plan", value: "10回券" }, { questionId: "q_bijiris_session_ticket_round", value: "10回目" }], visibleWhen: { questionId: "q_bijiris_session_type", value: "回数券" } },
    ],
  },
];

function doGet(e) {
  try {
    return output_(handleGet_(e || {}), e && e.parameter && e.parameter.callback);
  } catch (error) {
    return output_({ error: error.message || "エラーが発生しました。" }, e && e.parameter && e.parameter.callback);
  }
}

function doPost(e) {
  try {
    return output_(handlePost_(parsePost_(e || {})));
  } catch (error) {
    return output_({ error: error.message || "エラーが発生しました。" });
  }
}

function setup() {
  ensureSpreadsheet_();
  syncMaintenanceTrigger_(getPreferences_());
  var credentials = getAdminCredentials_();
  var rootFolder = getRootPhotoFolder_();
  return {
    ok: true,
    spreadsheetUrl: getSpreadsheet_().getUrl(),
    rootFolderUrl: rootFolder.getUrl(),
    loginId: credentials.username,
    message: "初期設定が完了しました。",
  };
}

function resetSurveysToDefaults() {
  var surveys = SURVEYS.map(function (survey, index) {
    var normalized = validateSurveyPayload_(Object.assign({}, survey, { sortOrder: index }), null);
    normalized.sortOrder = index;
    return normalized;
  });
  saveSurveys_(normalizeSurveyOrder_(surveys));
  surveys.forEach(ensureSurveySheet_);
  appendAuditLog_("survey.reset_defaults", {
    count: surveys.length,
    version: VERSION,
  });
  return {
    ok: true,
    version: VERSION,
    surveys: getSurveys_(),
  };
}

function handleGet_(e) {
  var params = e.parameter || {};
  var action = params.action || "health";

  if (action === "health") return { ok: true, backend: "gas" };
  if (action === "surveys") {
    return {
      surveys: getPublicSurveys_(),
      dataPolicyText: getPreferences_().dataPolicyText,
      requireConsent: getPreferences_().requireConsent,
      consentText: getPreferences_().consentText,
      milestoneRewardConfig: getPreferences_().milestoneRewardConfig,
      pushAppId: getPushAppId_(),
      version: VERSION,
    };
  }
  if (action === "bijirisPosts") return { posts: getBijirisPosts_({ publishedOnly: true }) };
  if (action === "history") {
    return getCustomerHistoryPayload_({
      clientId: params.clientId,
      customerName: params.name,
      customerNameKana: params.nameKana,
      matchByNameOnly: params.recoverByName === "1",
      includeTrashed: false,
    });
  }
  if (action === "adminLogin") return adminLogin_(params.loginId, params.password);
  if (action === "adminVerifyOtp") return verifyAdminOtp_(params.sessionId, params.code);

  requireAdmin_(params.token);
  if (action === "adminInfo") return getAdminInfo_();
  if (action === "adminUpdateCredentials") {
    return updateAdminCredentials_(params.loginId, params.password);
  }
  if (action === "adminUsers") return { adminUsers: getAdminUsers_().map(publicAdminUser_) };
  if (action === "adminSurveys") return { surveys: getSurveys_() };
  if (action === "adminPreferences") return { preferences: getPreferences_() };
  if (action === "adminLogs") return getLogs_();
  if (action === "adminCustomerMemos") return { memos: getCustomerMemos_() };
  if (action === "adminResponses") return { responses: getResponses_({ includeTrashed: true }) };
  if (action === "adminMeasurements") return { measurements: getMeasurements_({}) };
  if (action === "adminBijirisPosts") return { posts: getBijirisPosts_({ includeDrafts: true }) };
  if (action === "adminUpdate") {
    return {
      response: updateResponse_(params.responseId, params.status, params.adminMemo),
    };
  }
  if (action === "adminDelete") {
    deleteResponse_(params.responseId);
    return { ok: true };
  }
  if (action === "adminExport") {
    return {
      surveys: getSurveys_(),
      responses: getResponses_({ includeTrashed: true }),
      preferences: getPreferences_(),
      customerMemos: getCustomerMemos_(),
      customerProfiles: getAdminCustomerProfiles_(),
      measurements: getMeasurements_({}),
      bijirisPosts: getBijirisPosts_({ includeDrafts: true }),
      adminUsers: getAdminUsers_().map(publicAdminUser_),
      exportedAt: new Date().toISOString(),
    };
  }

  throw new Error("API が見つかりません。");
}

function handlePost_(body) {
  if (body.action === "submitResponse") {
    return saveResponse_(body);
  }
  if (body.action === "updatePublicTicketCard") {
    return updatePublicTicketCard_(body);
  }
  if (body.action === "updatePublicPushStatus") {
    return updatePublicPushStatus_(body);
  }
  if (body.action === "updatePublicResponse") {
    return updatePublicResponse_(body);
  }
  if (body.action === "logClientError") {
    return logClientError_(body.payload || {});
  }
  if (body.action === "adminUpdateResponse") {
    requireAdmin_(body.token);
    return {
      response: updateResponse_(
        body.responseId,
        body.payload && body.payload.status,
        body.payload && body.payload.adminMemo,
        body.payload && body.payload.answers
      ),
    };
  }
  if (body.action === "adminCreateSurvey") {
    requireAdmin_(body.token);
    return {
      survey: createSurvey_(body.payload || {}),
    };
  }
  if (body.action === "adminUpdateSurvey") {
    requireAdmin_(body.token);
    return {
      survey: updateSurveyDefinition_(body.surveyId, body.payload || {}),
    };
  }
  if (body.action === "adminDeleteSurvey") {
    requireAdmin_(body.token);
    deleteSurveyDefinition_(body.surveyId);
    return { ok: true };
  }
  if (body.action === "adminReplaceSurveys") {
    requireAdmin_(body.token);
    return {
      surveys: replaceSurveys_(body.payload && body.payload.surveys),
    };
  }
  if (body.action === "adminUpdatePreferences") {
    requireAdmin_(body.token);
    return {
      preferences: updatePreferences_(body.payload || {}),
    };
  }
  if (body.action === "adminUpdatePushConfig") {
    requireAdmin_(body.token);
    return {
      adminInfo: updatePushConfig_(body.payload || {}, { source: "admin" }),
    };
  }
  if (body.action === "adminUpdateCustomerMemo") {
    requireAdmin_(body.token);
    return {
      memos: updateCustomerMemo_(body.customerName, body.memo, body.at),
    };
  }
  if (body.action === "adminUpdateCustomer") {
    requireAdmin_(body.token);
    return updateCustomerProfile_(body.customerName, body.payload || {});
  }
  if (body.action === "adminDeleteCustomer") {
    requireAdmin_(body.token);
    return deleteCustomerProfile_(body.customerName);
  }
  if (body.action === "adminUpdateRewardRedemption") {
    requireAdmin_(body.token);
    return updateCustomerRewardRedemption_(body.customerName, body.threshold, body.handed === true);
  }
  if (body.action === "adminCreateMeasurement") {
    requireAdmin_(body.token);
    return {
      measurement: createMeasurement_(body.customerName, body.payload || {}),
    };
  }
  if (body.action === "adminUpdateMeasurement") {
    requireAdmin_(body.token);
    return {
      measurement: updateMeasurement_(body.measurementId, body.payload || {}),
    };
  }
  if (body.action === "adminDeleteMeasurement") {
    requireAdmin_(body.token);
    return deleteMeasurement_(body.measurementId);
  }
  if (body.action === "adminUploadAnalysisSheet") {
    requireAdmin_(body.token);
    return {
      file: uploadAnalysisSheetFile_(body.payload || {}),
    };
  }
  if (body.action === "adminReplaceMeasurements") {
    requireAdmin_(body.token);
    return {
      measurements: replaceMeasurements_(body.payload && body.payload.measurements),
    };
  }
  if (body.action === "adminCreateBijirisPost") {
    requireAdmin_(body.token);
    return {
      post: createBijirisPost_(body.payload || {}),
    };
  }
  if (body.action === "adminUpdateBijirisPost") {
    requireAdmin_(body.token);
    return {
      post: updateBijirisPost_(body.postId, body.payload || {}),
    };
  }
  if (body.action === "adminDeleteBijirisPost") {
    requireAdmin_(body.token);
    return deleteBijirisPost_(body.postId);
  }
  if (body.action === "adminReplaceBijirisPosts") {
    requireAdmin_(body.token);
    return {
      posts: replaceBijirisPosts_(body.payload && body.payload.posts),
    };
  }
  if (body.action === "adminUpdateUsers") {
    requireAdmin_(body.token);
    return {
      adminUsers: updateAdminUsers_(body.payload && body.payload.adminUsers),
    };
  }
  if (body.action === "adminRunMaintenance") {
    requireAdmin_(body.token);
    return runScheduledMaintenance();
  }
  throw new Error("API が見つかりません。");
}

function parsePost_(e) {
  var contents = e.postData && e.postData.contents ? e.postData.contents : "";
  if (!contents && e.parameter && e.parameter.payload) contents = e.parameter.payload;
  if (!contents) return {};
  return JSON.parse(contents);
}

function output_(data, callback) {
  var body = JSON.stringify(data);
  if (callback) {
    body = String(callback).replace(/[^\w.$]/g, "") + "(" + body + ");";
    return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function getSurveys_() {
  return loadSurveys_().map(cloneSurvey_);
}

function getPublicSurveys_() {
  return getSurveys_().filter(function (survey) {
    return normalizeSurveyStatus_(survey.status) === "published";
  });
}

function findSurvey_(surveyId) {
  var surveys = loadSurveys_();
  for (var i = 0; i < surveys.length; i += 1) {
    if (surveys[i].id === surveyId) return surveys[i];
  }
  throw new Error("アンケートが見つかりません。");
}

function loadSurveys_() {
  var properties = PropertiesService.getScriptProperties();
  var stored = parseJson_(properties.getProperty(SURVEYS_PROPERTY_KEY), null);
  if (Array.isArray(stored)) {
    var normalizedStored = stored.map(function (survey) {
      return validateSurveyPayload_(mergeDefaultSurveyFields_(survey), survey);
    });
    normalizedStored.sort(compareSurveyOrder_);
    if (JSON.stringify(stored) !== JSON.stringify(normalizedStored)) {
      saveSurveys_(normalizedStored);
    }
    return normalizedStored;
  }

  var defaults = SURVEYS.map(function (survey) {
    return validateSurveyPayload_(survey, survey);
  });
  defaults.sort(compareSurveyOrder_);
  saveSurveys_(defaults);
  return defaults;
}

function mergeDefaultSurveyFields_(survey) {
  var normalizedSurvey = cloneSurvey_(survey);
  var defaults = null;
  for (var i = 0; i < SURVEYS.length; i += 1) {
    if (SURVEYS[i].id === normalizedSurvey.id) {
      defaults = SURVEYS[i];
      break;
    }
  }
  if (!defaults || !Array.isArray(normalizedSurvey.questions)) return normalizedSurvey;

  var defaultQuestionMap = {};
  defaults.questions.forEach(function (question) {
    defaultQuestionMap[question.id] = question;
  });

  normalizedSurvey.questions = normalizedSurvey.questions.map(function (question) {
    var defaultQuestion = defaultQuestionMap[normalizeText_(question && question.id)];
    if (!defaultQuestion) return question;
    var currentConditions = getQuestionVisibilityConditions_(question);
    var defaultConditions = getQuestionVisibilityConditions_(defaultQuestion);
    var merged = Object.assign({}, question, {
      visibilityConditions: defaultConditions,
      visibleWhen: defaultConditions.length ? defaultConditions[0] : null,
    });
    if (Array.isArray(defaultQuestion.options) && defaultQuestion.options.length) {
      var currentOptions = Array.isArray(question.options) ? question.options.slice() : [];
      defaultQuestion.options.forEach(function (option) {
        if (currentOptions.indexOf(option) === -1) currentOptions.push(option);
      });
      merged.options = currentOptions;
    }
    if (currentConditions.length) {
      merged.visibilityConditions = currentConditions;
      merged.visibleWhen = currentConditions[0] || null;
    }
    return merged;
  });
  var existingQuestionIds = normalizedSurvey.questions.map(function (question) {
    return normalizeText_(question && question.id);
  });
  defaults.questions.forEach(function (question) {
    if (existingQuestionIds.indexOf(normalizeText_(question && question.id)) >= 0) return;
    normalizedSurvey.questions.push(JSON.parse(JSON.stringify(question)));
  });
  if (!normalizedSurvey.introMessage) normalizedSurvey.introMessage = defaults.introMessage || defaults.description || "";
  if (!normalizedSurvey.completionMessage) {
    normalizedSurvey.completionMessage = defaults.completionMessage || "ご回答ありがとうございました。";
  }
  return normalizedSurvey;
}

function saveSurveys_(surveys) {
  PropertiesService.getScriptProperties().setProperty(
    SURVEYS_PROPERTY_KEY,
    JSON.stringify(Array.isArray(surveys) ? surveys : [])
  );
}

function cloneSurvey_(survey) {
  return JSON.parse(JSON.stringify(survey));
}

function compareSurveyOrder_(left, right) {
  return getSurveySortOrder_(left) - getSurveySortOrder_(right);
}

function getSurveySortOrder_(survey) {
  var sortOrder = Number(survey && survey.sortOrder);
  return Number.isFinite(sortOrder) ? sortOrder : 999999;
}

function makeId_(prefix) {
  return String(prefix) + "_" + Utilities.getUuid();
}

function isChoiceType_(type) {
  return type === "choice" || type === "checkbox";
}

function validateSurveyPayload_(payload, existing) {
  var title = normalizeText_(payload && payload.title);
  var description = normalizeText_(payload && payload.description);
  var introMessage = normalizeText_(payload && payload.introMessage) || description;
  var completionMessage = normalizeText_(payload && payload.completionMessage) || "ご回答ありがとうございました。";
  var status = normalizeSurveyStatus_(payload && payload.status);
  var questions = Array.isArray(payload && payload.questions) ? payload.questions : [];
  var surveys = loadSurveysWithoutValidation_();
  var sortOrder = existing && existing.sortOrder !== undefined
    ? Number(existing.sortOrder)
    : payload && payload.sortOrder !== undefined
      ? Number(payload.sortOrder)
      : surveys.length;
  var acceptingResponses = !(payload && payload.acceptingResponses === false);
  var startAt = normalizeDateTime_(payload && payload.startAt);
  var endAt = normalizeDateTime_(payload && payload.endAt);

  if (!title) throw new Error("タイトルを入力してください。");
  if (!description) throw new Error("説明文を入力してください。");
  if (!questions.length) throw new Error("質問は1つ以上必要です。");
  if (startAt && endAt && new Date(startAt).getTime() > new Date(endAt).getTime()) {
    throw new Error("受付終了日時は開始日時以降にしてください。");
  }

  var normalizedQuestions = questions.map(function (question) {
    var type = QUESTION_TYPES.indexOf(question && question.type) >= 0 ? question.type : "text";
    var label = normalizeText_(question && question.label);
    var options = Array.isArray(question && question.options)
      ? question.options.map(normalizeText_).filter(Boolean)
      : [];
    var visibilityConditions = validateVisibilityConditions_(
      question && question.visibilityConditions,
      question && question.visibleWhen
    );

    if (!label) throw new Error("質問文を入力してください。");
    if (isChoiceType_(type) && options.length < 2) {
      throw new Error("選択式の質問は選択肢を2つ以上入力してください。");
    }

    return {
      id: normalizeText_(question && question.id) || makeId_("question"),
      label: label,
      type: type,
      required: question && question.required === false ? false : true,
      options: isChoiceType_(type) ? options : [],
      visibilityConditions: visibilityConditions,
      visibleWhen: visibilityConditions.length ? visibilityConditions[0] : null,
    };
  });

  return {
    id: existing && existing.id ? existing.id : normalizeText_(payload && payload.id) || makeId_("survey"),
    title: title,
    description: description,
    introMessage: introMessage,
    completionMessage: completionMessage,
    status: status,
    sortOrder: Number.isFinite(sortOrder) ? sortOrder : surveys.length,
    acceptingResponses: acceptingResponses,
    startAt: startAt,
    endAt: endAt,
    questions: normalizedQuestions,
    createdAt: existing && existing.createdAt ? String(existing.createdAt) : new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}

function loadSurveysWithoutValidation_() {
  return parseJson_(PropertiesService.getScriptProperties().getProperty(SURVEYS_PROPERTY_KEY), SURVEYS) || SURVEYS;
}

function validateVisibleWhen_(visibleWhen) {
  if (!visibleWhen || typeof visibleWhen !== "object") return null;
  var questionId = normalizeText_(visibleWhen.questionId);
  var value = normalizeText_(visibleWhen.value);
  if (!questionId || !value) return null;
  return {
    questionId: questionId,
    value: value,
  };
}

function validateVisibilityConditions_(conditions, fallbackVisibleWhen) {
  var normalized = Array.isArray(conditions)
    ? conditions.map(validateVisibleWhen_).filter(Boolean)
    : [];
  if (normalized.length) return normalized;
  var fallback = validateVisibleWhen_(fallbackVisibleWhen);
  return fallback ? [fallback] : [];
}

function getQuestionVisibilityConditions_(question) {
  return validateVisibilityConditions_(
    question && question.visibilityConditions,
    question && question.visibleWhen
  );
}

function normalizeDateTime_(value) {
  var normalized = normalizeText_(value);
  if (!normalized) return "";
  var date = new Date(normalized);
  if (isNaN(date.getTime())) throw new Error("日時の形式が正しくありません。");
  return date.toISOString();
}

function createSurvey_(payload) {
  var surveys = loadSurveys_();
  var survey = validateSurveyPayload_(payload);
  surveys.unshift(survey);
  surveys = normalizeSurveyOrder_(surveys);
  saveSurveys_(surveys);
  ensureSurveySheet_(survey);
  appendAuditLog_("survey.create", {
    surveyId: survey.id,
    title: survey.title,
  });
  return cloneSurvey_(survey);
}

function updateSurveyDefinition_(surveyId, payload) {
  var surveys = loadSurveys_();
  var existing = null;
  for (var i = 0; i < surveys.length; i += 1) {
    if (surveys[i].id === surveyId) {
      existing = surveys[i];
      break;
    }
  }
  if (!existing) throw new Error("アンケートが見つかりません。");

  var survey = validateSurveyPayload_(payload, existing);
  surveys = surveys.map(function (item) {
    return item.id === surveyId ? survey : item;
  });
  surveys = normalizeSurveyOrder_(surveys);
  saveSurveys_(surveys);
  ensureSurveySheet_(survey);
  appendAuditLog_("survey.update", {
    surveyId: survey.id,
    title: survey.title,
  });
  return cloneSurvey_(survey);
}

function deleteSurveyDefinition_(surveyId) {
  var surveys = loadSurveys_();
  var filtered = surveys.filter(function (survey) {
    return survey.id !== surveyId;
  });
  if (filtered.length === surveys.length) throw new Error("アンケートが見つかりません。");
  saveSurveys_(normalizeSurveyOrder_(filtered));
  appendAuditLog_("survey.delete", {
    surveyId: surveyId,
  });
}

function replaceSurveys_(surveysPayload) {
  if (!Array.isArray(surveysPayload) || !surveysPayload.length) {
    throw new Error("復元するアンケートがありません。");
  }

  var existingMap = {};
  loadSurveys_().forEach(function (survey) {
    existingMap[survey.id] = survey;
  });

  var surveys = surveysPayload.map(function (survey, index) {
    var existing = existingMap[normalizeText_(survey && survey.id)] || null;
    var normalized = validateSurveyPayload_(Object.assign({}, survey, { sortOrder: index }), existing);
    normalized.sortOrder = index;
    return normalized;
  });

  saveSurveys_(normalizeSurveyOrder_(surveys));
  surveys.forEach(ensureSurveySheet_);
  appendAuditLog_("survey.restore", {
    count: surveys.length,
  });
  return getSurveys_();
}

function normalizeSurveyOrder_(surveys) {
  return surveys
    .slice()
    .sort(compareSurveyOrder_)
    .map(function (survey, index) {
      return Object.assign({}, survey, { sortOrder: index });
    });
}

function getPreferences_() {
  var properties = PropertiesService.getScriptProperties();
  var stored = parseJson_(properties.getProperty(PREFERENCES_PROPERTY_KEY), {});
  var preferences = {
    notificationEnabled: stored && stored.notificationEnabled !== false,
    notificationEmail: normalizeEmail_(stored && stored.notificationEmail) || normalizeEmail_(getOwnerEmail_()),
    notificationSubject: normalizeText_(stored && stored.notificationSubject) || DEFAULT_NOTIFICATION_SUBJECT_(),
    notificationBody: normalizeText_(stored && stored.notificationBody) || DEFAULT_NOTIFICATION_BODY_(),
    dataPolicyText: normalizeDataPolicyText_(stored && stored.dataPolicyText) || DEFAULT_DATA_POLICY_TEXT_(),
    requireConsent: stored && stored.requireConsent === false ? false : true,
    consentText: normalizeText_(stored && stored.consentText) || DEFAULT_CONSENT_TEXT_(),
    autoBackupEnabled: stored && stored.autoBackupEnabled === false ? false : true,
    backupHour: normalizeBackupHour_(stored && stored.backupHour),
    retentionDays: normalizeRetentionDays_(stored && stored.retentionDays),
    recoveryMemo: normalizeText_(stored && stored.recoveryMemo) || DEFAULT_RECOVERY_MEMO_(),
    twoFactorEnabled: false,
    bijirisCategoryConfig: normalizeBijirisCategoryConfig_(stored && stored.bijirisCategoryConfig),
    gachaPrizeConfig: normalizeGachaPrizeConfig_(stored && stored.gachaPrizeConfig),
    milestoneRewardConfig: normalizeMilestoneRewardConfig_(stored && stored.milestoneRewardConfig),
  };
  if (JSON.stringify(stored || {}) !== JSON.stringify(preferences)) {
    properties.setProperty(PREFERENCES_PROPERTY_KEY, JSON.stringify(preferences));
  }
  return preferences;
}

function updatePreferences_(payload) {
  var current = getPreferences_();
  var next = {
    notificationEnabled: payload && payload.notificationEnabled === false ? false : true,
    notificationEmail: normalizeEmail_(payload && payload.notificationEmail) || current.notificationEmail,
    notificationSubject: normalizeText_(payload && payload.notificationSubject) || current.notificationSubject,
    notificationBody: normalizeText_(payload && payload.notificationBody) || current.notificationBody,
    dataPolicyText: normalizeDataPolicyText_(payload && payload.dataPolicyText) || current.dataPolicyText,
    requireConsent: payload && payload.requireConsent === false ? false : true,
    consentText: normalizeText_(payload && payload.consentText) || current.consentText,
    autoBackupEnabled: payload && payload.autoBackupEnabled === false ? false : true,
    backupHour: normalizeBackupHour_(payload && payload.backupHour),
    retentionDays: normalizeRetentionDays_(payload && payload.retentionDays),
    recoveryMemo: normalizeText_(payload && payload.recoveryMemo) || current.recoveryMemo,
    twoFactorEnabled: false,
    bijirisCategoryConfig: normalizeBijirisCategoryConfig_(
      (payload && payload.bijirisCategoryConfig) || current.bijirisCategoryConfig
    ),
    gachaPrizeConfig: normalizeGachaPrizeConfig_(
      payload && payload.gachaPrizeConfig || current.gachaPrizeConfig
    ),
    milestoneRewardConfig: normalizeMilestoneRewardConfig_(
      payload && payload.milestoneRewardConfig || current.milestoneRewardConfig
    ),
  };

  if (next.notificationEnabled && !next.notificationEmail) {
    throw new Error("通知メールアドレスを入力してください。");
  }

  PropertiesService.getScriptProperties().setProperty(PREFERENCES_PROPERTY_KEY, JSON.stringify(next));
  syncMaintenanceTrigger_(next);
  appendAuditLog_("preferences.update", {
    notificationEnabled: next.notificationEnabled,
    notificationEmail: next.notificationEmail,
    autoBackupEnabled: next.autoBackupEnabled,
    retentionDays: next.retentionDays,
    twoFactorEnabled: false,
  });
  return next;
}

var GACHA_PRIZE_KEYS_ = ["A", "B", "C", "D"];

function normalizeMonthKey_(value) {
  var raw = normalizeText_(value);
  var matched = raw.match(/^(\d{4})[-\/年](\d{1,2})/);
  if (!matched) return "";
  var year = Number(matched[1]);
  var month = Number(matched[2]);
  if (!isFinite(year) || !isFinite(month) || month < 1 || month > 12) return "";
  return Utilities.formatString("%04d-%02d", year, month);
}

function getCurrentMonthKey_() {
  return normalizeMonthKey_(new Date().toISOString()) || "2026-04";
}

function normalizeGachaProbability_(value, fallback) {
  var numeric = Number(value);
  if (!isFinite(numeric)) return fallback;
  return Math.max(0, Math.min(100, Math.round(numeric * 10) / 10));
}

function createDefaultGachaMonthlyPrize_(month) {
  return {
    month: month || getCurrentMonthKey_(),
    prizes: {
      A: { content: "", probability: 25 },
      B: { content: "", probability: 25 },
      C: { content: "", probability: 25 },
      D: { content: "", probability: 25 }
    }
  };
}

function normalizeGachaPrizeConfig_(value) {
  var seenMonths = {};
  var monthlyPrizes = [];
  (Array.isArray(value && value.monthlyPrizes) ? value.monthlyPrizes : []).forEach(function (entry) {
    var month = normalizeMonthKey_(entry && entry.month);
    if (!month || seenMonths[month]) return;
    seenMonths[month] = true;
    var prizes = {};
    GACHA_PRIZE_KEYS_.forEach(function (key) {
      var prize = entry && entry.prizes && entry.prizes[key] || {};
      prizes[key] = {
        content: normalizeText_(prize.content),
        probability: normalizeGachaProbability_(prize.probability, 0)
      };
    });
    monthlyPrizes.push({ month: month, prizes: prizes });
  });
  monthlyPrizes.sort(function (a, b) {
    return a.month < b.month ? -1 : a.month > b.month ? 1 : 0;
  });
  return {
    monthlyPrizes: monthlyPrizes.length ? monthlyPrizes : [createDefaultGachaMonthlyPrize_(getCurrentMonthKey_())]
  };
}

function normalizeMilestoneRewardConfig_(value) {
  var enabled = !(value && value.enabled === false);
  var byThreshold = {};
  (Array.isArray(value && value.milestones) ? value.milestones : []).forEach(function (entry) {
    var threshold = Math.floor(Number(entry && entry.threshold));
    var reward = normalizeText_(entry && entry.reward);
    var description = normalizeText_(entry && entry.description);
    if (!isFinite(threshold) || threshold <= 0 || !reward) return;
    byThreshold[threshold] = { threshold: threshold, reward: reward, description: description };
  });
  var milestones = Object.keys(byThreshold)
    .map(function (key) { return byThreshold[key]; })
    .sort(function (a, b) { return a.threshold - b.threshold; })
    .slice(0, 20);
  return {
    enabled: enabled,
    milestones: milestones,
  };
}

function updatePushConfig_(payload, options) {
  var properties = PropertiesService.getScriptProperties();
  var nextPushAppId = normalizeText_(payload && payload.pushAppId) || getPushAppId_();
  var nextCustomerAppUrl = normalizeText_(payload && payload.customerAppUrl) || getCustomerAppUrl_();
  var nextRestApiKey = normalizeText_(payload && payload.restApiKey);
  var source = normalizeText_(options && options.source) || "unknown";

  if (!nextPushAppId) {
    throw new Error("OneSignal App ID を入力してください。");
  }
  if (!nextCustomerAppUrl) {
    throw new Error("顧客アプリURLを入力してください。");
  }
  if (!/^https?:\/\//i.test(nextCustomerAppUrl)) {
    throw new Error("顧客アプリURLは http または https で始まるURLを入力してください。");
  }

  properties.setProperty("ONESIGNAL_APP_ID", nextPushAppId);
  properties.setProperty("CUSTOMER_APP_URL", nextCustomerAppUrl);
  if (nextRestApiKey) {
    properties.setProperty("ONESIGNAL_REST_API_KEY", nextRestApiKey);
  }

  appendAuditLog_("push.config.update", {
    source: source,
    pushAppId: nextPushAppId,
    customerAppUrl: nextCustomerAppUrl,
    restApiKeyUpdated: !!nextRestApiKey,
    pushConfigured: isPushNotificationConfigured_(),
  });
  return getAdminInfo_();
}

function DEFAULT_DATA_POLICY_TEXT_() {
  return "ご回答内容と添付写真は、まゆみ助産院のアンケート管理と施術サポートのために利用します。";
}

function DEFAULT_NOTIFICATION_SUBJECT_() {
  return "【まゆみ助産院】新しいアンケート回答";
}

function DEFAULT_NOTIFICATION_BODY_() {
  return [
    "新しいアンケート回答が届きました。",
    "",
    "お名前: {{customerName}}",
    "アンケート: {{surveyTitle}}",
    "送信日時: {{submittedAt}}",
    "回答ID: {{responseId}}",
  ].join("\n");
}

function DEFAULT_CONSENT_TEXT_() {
  return "回答内容と添付写真の利用に同意します。";
}

function DEFAULT_RECOVERY_MEMO_() {
  return "障害時は Apps Script の実行ログ、スプレッドシート、Google ドライブの保存フォルダを確認してください。";
}

function DEFAULT_BIJIRIS_GENERAL_CATEGORIES_() {
  return ["豆知識", "アドバイス", "セルフケア", "お知らせ", "よくある質問"];
}

function DEFAULT_BIJIRIS_CONCERN_ROOT_LABEL_() {
  return "お悩み";
}

function DEFAULT_BIJIRIS_CONCERN_PATHS_() {
  return [
    "女性 > 産後女性を含む骨盤底筋まわりのお悩み > 産後の骨盤底筋のゆるみ感が気になる",
    "女性 > 産後女性を含む骨盤底筋まわりのお悩み > くしゃみ、咳、大笑いでヒヤッとする",
    "女性 > 産後女性を含む骨盤底筋まわりのお悩み > 骨盤まわりを土台からケアしたい",
    "女性 > トイレまわりのお悩み > 尿漏れが気になる",
    "女性 > トイレまわりのお悩み > 頻尿が気になる",
    "女性 > トイレまわりのお悩み > 急な尿意が気になる",
    "女性 > トイレまわりのお悩み > 便秘がち",
    "女性 > トイレまわりのお悩み > お通じのリズムが気になる",
    "女性 > 体型・見た目のお悩み > 産後のぽっこりお腹が気になる",
    "女性 > 体型・見た目のお悩み > 年齢とともに体型の変化が気になる",
    "女性 > 体型・見た目のお悩み > 下半身太りが気になる",
    "女性 > 体型・見た目のお悩み > ヒップの下垂が気になる",
    "女性 > 姿勢・日常動作のお悩み > 産後に姿勢が崩れやすくなった",
    "女性 > 姿勢・日常動作のお悩み > 抱っこや家事で下腹や骨盤まわりが気になる",
    "女性 > 姿勢・日常動作のお悩み > 姿勢を整えたい",
    "女性 > デリケートゾーンまわりのお悩み > デリケートゾーンのケアを意識したい",
    "女性 > デリケートゾーンまわりのお悩み > 膣トレを始めてみたい",
    "女性 > 冷え・巡りのお悩み > 冷えやすさが気になる",
    "男性 > トイレまわりのお悩み > 頻尿が気になる",
    "男性 > トイレまわりのお悩み > ちょい漏れが気になる",
    "男性 > トイレまわりのお悩み > 急な尿意で不安がある",
    "男性 > トイレまわりのお悩み > トイレ悩みをケアしたい",
    "男性 > デリケートなお悩み > EDケアを意識したい",
    "男性 > デリケートなお悩み > デリケートなお悩みを人知れずケアしたい",
    "男性 > 姿勢・骨盤まわりのお悩み > 長時間の座り仕事で骨盤まわりが気になる",
    "男性 > 姿勢・骨盤まわりのお悩み > 猫背や前かがみ姿勢が気になる",
    "男性 > 姿勢・骨盤まわりのお悩み > 腰まわりの違和感が気になる",
    "男性 > 下半身・体型のお悩み > 下半身の筋力低下が気になる",
    "男性 > 下半身・体型のお悩み > ヒップラインの崩れが気になる",
    "男性 > 下半身・体型のお悩み > むくみや冷えが気になる",
    "男性 > 下半身・体型のお悩み > 運動不足が気になる",
    "男性 > 下半身・体型のお悩み > 筋トレが続かない",
  ];
}

function normalizeTextList_(values, fallback) {
  var normalized = (Array.isArray(values) ? values : [])
    .map(function (value) {
      return normalizeText_(value);
    })
    .filter(function (value) {
      return !!value;
    });
  return normalized.length ? normalized : (Array.isArray(fallback) ? fallback.slice() : []);
}

function normalizeBijirisCategoryConfig_(value) {
  var defaults = {
    generalCategories: DEFAULT_BIJIRIS_GENERAL_CATEGORIES_(),
    concernRootLabel: DEFAULT_BIJIRIS_CONCERN_ROOT_LABEL_(),
    concernPaths: DEFAULT_BIJIRIS_CONCERN_PATHS_(),
  };
  return {
    generalCategories: normalizeTextList_(value && value.generalCategories, defaults.generalCategories),
    concernRootLabel: normalizeText_(value && value.concernRootLabel) || defaults.concernRootLabel,
    concernPaths: normalizeTextList_(value && value.concernPaths, defaults.concernPaths),
  };
}

function normalizeBackupHour_(value) {
  var hour = Number(value);
  return Number.isFinite(hour) && hour >= 0 && hour <= 23 ? hour : 3;
}

function normalizeRetentionDays_(value) {
  var days = Number(value);
  return Number.isFinite(days) && days >= 0 ? Math.floor(days) : 365;
}

function publicAdminUser_(user) {
  return {
    id: normalizeText_(user && user.id),
    username: normalizeText_(user && user.username),
    email: normalizeEmail_(user && user.email),
    active: user && user.active === false ? false : true,
  };
}

function normalizeAdminUsers_(users, currentUsers) {
  var current = Array.isArray(currentUsers) ? currentUsers : [];
  var currentById = {};
  current.forEach(function (user) {
    currentById[user.id] = user;
  });

  var normalizedUsers = (Array.isArray(users) ? users : []).map(function (user, index) {
    var id = normalizeText_(user && user.id) || makeId_("admin");
    var username = normalizeText_(user && user.username);
    var existing = currentById[id] || null;
    var password = normalizeText_(user && user.password) || (existing && existing.password) || "";
    if (!username) throw new Error("管理者ログインIDを入力してください。");
    if (!password || password.length < 4) throw new Error("管理者パスワードは4文字以上で入力してください。");
    return {
      id: id,
      username: username,
      password: password,
      email: normalizeEmail_(user && user.email),
      active: user && user.active === false ? false : true,
      sortOrder: index,
    };
  });

  if (!normalizedUsers.length) {
    throw new Error("管理者アカウントは1件以上必要です。");
  }

  var seen = {};
  normalizedUsers.forEach(function (user) {
    if (seen[user.username]) throw new Error("管理者ログインIDが重複しています。");
    seen[user.username] = true;
  });

  return normalizedUsers;
}

function getAdminUsers_() {
  var properties = PropertiesService.getScriptProperties();
  var stored = parseJson_(properties.getProperty(ADMIN_USERS_PROPERTY_KEY), null);
  if (Array.isArray(stored) && stored.length) {
    var normalized = normalizeAdminUsers_(stored, stored);
    if (JSON.stringify(stored) !== JSON.stringify(normalized)) {
      properties.setProperty(ADMIN_USERS_PROPERTY_KEY, JSON.stringify(normalized));
    }
    return normalized;
  }

  var migrated = [{
    id: "admin_primary",
    username: normalizeText_(properties.getProperty("ADMIN_USERNAME")) || DEFAULT_ADMIN_USERNAME,
    password: String(properties.getProperty("ADMIN_PASSWORD") || "") || DEFAULT_ADMIN_PASSWORD,
    email: normalizeEmail_(getOwnerEmail_()),
    active: true,
    sortOrder: 0,
  }];

  if (migrated[0].username === LEGACY_ADMIN_USERNAME) migrated[0].username = DEFAULT_ADMIN_USERNAME;
  if (!migrated[0].password || migrated[0].password === LEGACY_ADMIN_PASSWORD) migrated[0].password = DEFAULT_ADMIN_PASSWORD;

  properties.setProperty(ADMIN_USERS_PROPERTY_KEY, JSON.stringify(migrated));
  properties.setProperty("ADMIN_USERNAME", migrated[0].username);
  properties.setProperty("ADMIN_PASSWORD", migrated[0].password);
  return migrated;
}

function findAdminUserByUsername_(loginId) {
  var username = normalizeText_(loginId);
  var users = getAdminUsers_();
  for (var i = 0; i < users.length; i += 1) {
    if (users[i].active !== false && users[i].username === username) return users[i];
  }
  return null;
}

function updateAdminUsers_(users) {
  var normalized = normalizeAdminUsers_(users, getAdminUsers_());
  PropertiesService.getScriptProperties().setProperty(ADMIN_USERS_PROPERTY_KEY, JSON.stringify(normalized));
  PropertiesService.getScriptProperties().setProperty("ADMIN_USERNAME", normalized[0].username);
  PropertiesService.getScriptProperties().setProperty("ADMIN_PASSWORD", normalized[0].password);
  normalized.forEach(function (user) {
    clearLoginFailures_(user.username);
  });
  appendAuditLog_("admin.users.update", {
    count: normalized.length,
  });
  return normalized.map(publicAdminUser_);
}

function getOtpSessions_() {
  return parseJson_(PropertiesService.getScriptProperties().getProperty(OTP_SESSIONS_PROPERTY_KEY), []);
}

function saveOtpSessions_(sessions) {
  PropertiesService.getScriptProperties().setProperty(OTP_SESSIONS_PROPERTY_KEY, JSON.stringify(sessions || []));
}

function pruneOtpSessions_(sessions) {
  var nowTime = Date.now();
  return (Array.isArray(sessions) ? sessions : []).filter(function (session) {
    return Number(session.expiresAt || 0) > nowTime;
  });
}

function createOtpSession_(user) {
  var code = String(Math.floor(100000 + Math.random() * 900000));
  var session = {
    id: makeId_("otp"),
    username: user.username,
    code: code,
    expiresAt: Date.now() + OTP_TTL_MS,
  };
  var sessions = pruneOtpSessions_(getOtpSessions_());
  sessions = sessions.filter(function (item) {
    return item.username !== user.username;
  });
  sessions.unshift(session);
  saveOtpSessions_(sessions);
  return session;
}

function sendOtpEmail_(user, session) {
  var preferences = getPreferences_();
  var to = normalizeEmail_(user.email) || preferences.notificationEmail || normalizeEmail_(getOwnerEmail_());
  if (!to) throw new Error("2段階認証メールの送信先が設定されていません。");
  MailApp.sendEmail(
    to,
    "【まゆみ助産院】確認コード",
    [
      "確認コードを入力してください。",
      "",
      "ログインID: " + user.username,
      "確認コード: " + session.code,
      "有効期限: " + new Date(session.expiresAt).toLocaleString("ja-JP"),
    ].join("\n")
  );
}

function verifyAdminOtp_(sessionId, code) {
  var normalizedSessionId = normalizeText_(sessionId);
  var normalizedCode = normalizeText_(code);
  var sessions = pruneOtpSessions_(getOtpSessions_());
  var session = null;
  var remaining = [];
  sessions.forEach(function (item) {
    if (item.id === normalizedSessionId) {
      session = item;
      return;
    }
    remaining.push(item);
  });
  if (!session) throw new Error("確認コードの有効期限が切れました。もう一度ログインしてください。");
  if (session.code !== normalizedCode) throw new Error("確認コードが違います。");
  saveOtpSessions_(remaining);
  clearLoginFailures_(session.username);
  var expiresAt = Date.now() + 8 * 60 * 60 * 1000;
  appendAuditLog_("admin.login.otp", {
    loginId: session.username,
  });
  return {
    token: makeToken_(session.username, expiresAt),
    expiresAt: new Date(expiresAt).toISOString(),
  };
}

function getBackupFolder_() {
  var rootFolder = getRootPhotoFolder_();
  var folders = rootFolder.getFoldersByName("system-backups");
  return folders.hasNext() ? folders.next() : rootFolder.createFolder("system-backups");
}

function getMaintenanceTriggerIds_() {
  return parseJson_(PropertiesService.getScriptProperties().getProperty(MAINTENANCE_TRIGGER_IDS_PROPERTY_KEY), []);
}

function saveMaintenanceTriggerIds_(ids) {
  PropertiesService.getScriptProperties().setProperty(
    MAINTENANCE_TRIGGER_IDS_PROPERTY_KEY,
    JSON.stringify(ids || [])
  );
}

function syncMaintenanceTrigger_(preferences) {
  var triggerIds = getMaintenanceTriggerIds_();
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (triggerIds.indexOf(trigger.getUniqueId()) >= 0) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  saveMaintenanceTriggerIds_([]);
  if (!preferences || (!preferences.autoBackupEnabled && !(preferences.retentionDays > 0))) return;
  var trigger = ScriptApp.newTrigger("runScheduledMaintenance")
    .timeBased()
    .everyDays(1)
    .atHour(preferences.backupHour || 3)
    .create();
  saveMaintenanceTriggerIds_([trigger.getUniqueId()]);
}

function renderTemplate_(template, data) {
  var result = String(template || "");
  Object.keys(data || {}).forEach(function (key) {
    result = result.replace(new RegExp("{{" + key + "}}", "g"), String(data[key] || ""));
  });
  return result;
}

function writeBackupFile_() {
  var backupFolder = getBackupFolder_();
  var payload = {
    exportedAt: new Date().toISOString(),
    surveys: getSurveys_(),
    responses: getResponses_({ includeTrashed: true }),
    preferences: getPreferences_(),
    customerMemos: getCustomerMemos_(),
    customerProfiles: getAdminCustomerProfiles_(),
    measurements: getMeasurements_({}),
    bijirisPosts: getBijirisPosts_({ includeDrafts: true }),
    adminUsers: getAdminUsers_().map(publicAdminUser_),
  };
  var fileName = "backup_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss") + ".json";
  var file = backupFolder.createFile(fileName, JSON.stringify(payload, null, 2), MimeType.PLAIN_TEXT);
  var meta = {
    at: new Date().toISOString(),
    fileId: file.getId(),
    fileName: file.getName(),
    fileUrl: file.getUrl(),
  };
  saveBackupMeta_(meta);
  return meta;
}

function purgeOldTrashResponses_() {
  var preferences = getPreferences_();
  var retentionDays = Number(preferences.retentionDays || 0);
  if (!(retentionDays > 0)) return 0;
  var threshold = Date.now() - retentionDays * 24 * 60 * 60 * 1000;
  var purged = 0;
  getResponses_({ includeTrashed: true }).forEach(function (response) {
    var referenceTime = response.managedAt || response.submittedAt;
    if (response.status !== "trash") return;
    if (new Date(referenceTime).getTime() > threshold) return;
    if (purgeResponse_(response.id)) purged += 1;
  });
  return purged;
}

function runScheduledMaintenance() {
  var preferences = getPreferences_();
  var backupInfo = null;
  if (preferences.autoBackupEnabled) {
    backupInfo = writeBackupFile_();
  }
  var purged = purgeOldTrashResponses_();
  var maintenanceMeta = {
    at: new Date().toISOString(),
    purged: purged,
    autoBackupEnabled: preferences.autoBackupEnabled,
    backupInfo: backupInfo,
  };
  saveLastMaintenanceMeta_(maintenanceMeta);
  appendAuditLog_("maintenance.run", {
    purged: purged,
    autoBackupEnabled: preferences.autoBackupEnabled,
    backupInfo: backupInfo,
  });
  return {
    ok: true,
    purged: purged,
    autoBackupEnabled: preferences.autoBackupEnabled,
    backupInfo: backupInfo,
    ranAt: maintenanceMeta.at,
  };
}

function normalizeDataPolicyText_(value) {
  return normalizeText_(value)
    .replace(/保存先は Google スプレッドシートおよび Google ドライブです。?/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeCustomerMemoRecord_(value) {
  if (!value) return {
    latestMemo: "",
    entries: [],
  };
  if (typeof value === "string") {
    var latestMemo = normalizeText_(value);
    return {
      latestMemo: latestMemo,
      entries: latestMemo ? [{ at: "", memo: latestMemo }] : [],
    };
  }
  var latest = normalizeText_(value.latestMemo || value.memo);
  var entries = Array.isArray(value.entries)
    ? value.entries
        .map(function (entry) {
          return {
            at: normalizeMemoDate_(entry && entry.at),
            memo: normalizeText_(entry && (entry.memo || entry.content)),
          };
        })
        .filter(function (entry) { return entry.memo; })
        .sort(function (a, b) {
          return memoEntrySortValue_(b.at) - memoEntrySortValue_(a.at);
        })
    : [];
  if (entries.length) latest = entries[0].memo;
  return {
    latestMemo: latest,
    entries: entries.slice(0, 100),
  };
}

function getTodayMemoDate_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Tokyo", "yyyy-MM-dd");
}

function normalizeMemoDate_(value) {
  var text = normalizeText_(value);
  if (!text) return "";
  var matched = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (matched) {
    return matched[1] + "-" + matched[2] + "-" + matched[3];
  }
  var parsed = new Date(text);
  if (isNaN(parsed.getTime())) return "";
  return Utilities.formatDate(parsed, Session.getScriptTimeZone() || "Asia/Tokyo", "yyyy-MM-dd");
}

function memoEntrySortValue_(value) {
  var text = normalizeMemoDate_(value);
  if (!text) return 0;
  var parsed = new Date(text + "T00:00:00");
  return isNaN(parsed.getTime()) ? 0 : parsed.getTime();
}

function getCustomerMemos_() {
  var stored = parseJson_(PropertiesService.getScriptProperties().getProperty(CUSTOMER_MEMOS_PROPERTY_KEY), {});
  var normalized = {};
  Object.keys(stored || {}).forEach(function (customerName) {
    normalized[customerName] = normalizeCustomerMemoRecord_(stored[customerName]);
  });
  if (JSON.stringify(stored || {}) !== JSON.stringify(normalized)) {
    PropertiesService.getScriptProperties().setProperty(CUSTOMER_MEMOS_PROPERTY_KEY, JSON.stringify(normalized));
  }
  return normalized;
}

function updateCustomerMemo_(customerName, memo, at) {
  var name = normalizeText_(customerName);
  if (!name) throw new Error("お客様名が必要です。");
  var memos = getCustomerMemos_();
  var normalizedMemo = normalizeText_(memo);
  var normalizedAt = normalizeMemoDate_(at);
  if (normalizedMemo) {
    var current = normalizeCustomerMemoRecord_(memos[name]);
    var entries = Array.isArray(current.entries) ? current.entries.slice() : [];
    if (normalizedAt) {
      var exists = entries.some(function (entry) {
        return normalizeText_(entry.memo) === normalizedMemo && normalizeMemoDate_(entry.at) === normalizedAt;
      });
      if (!exists) {
        entries.unshift({
          at: normalizedAt,
          memo: normalizedMemo,
        });
      }
    } else if (normalizedMemo !== current.latestMemo) {
      entries.unshift({
        at: getTodayMemoDate_(),
        memo: normalizedMemo,
      });
    }
    memos[name] = normalizeCustomerMemoRecord_({
      latestMemo: current.latestMemo,
      entries: entries,
    });
  } else {
    delete memos[name];
  }
  PropertiesService.getScriptProperties().setProperty(CUSTOMER_MEMOS_PROPERTY_KEY, JSON.stringify(memos));
  appendAuditLog_("customer.memo.update", {
    customerName: name,
  });
  return memos;
}

function normalizeMeasurementValue_(value) {
  if (value === null || value === undefined || value === "") return "";
  var normalized = String(value).replace(/[^\d.-]/g, "");
  if (!normalized) return "";
  var parsed = Number(normalized);
  if (!isFinite(parsed)) return "";
  return Math.round(parsed * 10) / 10;
}

function normalizeMeasurementDate_(value) {
  var normalized = normalizeText_(value);
  if (!normalized) return "";
  var date = new Date(normalized);
  if (isNaN(date.getTime())) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function computeMeasurementWhr_(waist, hip) {
  var normalizedWaist = normalizeMeasurementValue_(waist);
  var normalizedHip = normalizeMeasurementValue_(hip);
  if (normalizedWaist === "" || normalizedHip === "" || !(normalizedHip > 0)) return "";
  return Math.round((normalizedWaist / normalizedHip) * 1000) / 1000;
}

function normalizeMeasurementTargets_(value) {
  var targets = {
    waist: normalizeMeasurementValue_(value && value.waist),
    hip: normalizeMeasurementValue_(value && value.hip),
    thighRight: normalizeMeasurementValue_(value && value.thighRight),
    thighLeft: normalizeMeasurementValue_(value && value.thighLeft),
  };
  var hasAny = Object.keys(targets).some(function (key) {
    return targets[key] !== "";
  });
  return hasAny ? targets : null;
}

function normalizeRewardRedemptions_(value) {
  if (!value || typeof value !== "object") return null;
  var result = {};
  Object.keys(value).forEach(function (key) {
    var threshold = Math.floor(Number(key));
    if (!isFinite(threshold) || threshold <= 0) return;
    var entry = value[key] || {};
    if (entry.handed !== true) return;
    result[String(threshold)] = {
      handed: true,
      handedAt: normalizeTicketCardAcquiredAt_(entry.handedAt) || new Date().toISOString(),
    };
  });
  return Object.keys(result).length ? result : null;
}

function publicRewardRedemptions_(value) {
  return normalizeRewardRedemptions_(value);
}

function normalizePushPermission_(value) {
  var normalized = normalizeText_(value).toLowerCase();
  if (["granted", "denied", "default", "unsupported"].indexOf(normalized) >= 0) {
    return normalized;
  }
  return "";
}

function normalizePushStatus_(value) {
  if (!value || typeof value !== "object") return null;
  var hasEnabled = Object.prototype.hasOwnProperty.call(value, "enabled");
  var hasSupported = Object.prototype.hasOwnProperty.call(value, "supported");
  var permission = normalizePushPermission_(value.permission);
  var hasPermission = Boolean(permission);
  if (!hasEnabled && !hasSupported && !hasPermission) return null;
  return {
    enabled: value.enabled === true,
    supported: value.supported === true,
    permission: permission || (value.supported === false ? "unsupported" : ""),
    updatedAt: normalizeText_(value.updatedAt) || new Date().toISOString(),
  };
}

function publicPushStatus_(value) {
  var normalized = normalizePushStatus_(value);
  if (!normalized) return null;
  return {
    enabled: normalized.enabled,
    supported: normalized.supported,
    permission: normalized.permission,
    updatedAt: normalized.updatedAt,
  };
}

function publicMeasurementTargets_(value) {
  var normalized = normalizeMeasurementTargets_(value);
  if (!normalized) return null;
  return {
    waist: normalized.waist === "" ? "" : normalized.waist,
    hip: normalized.hip === "" ? "" : normalized.hip,
    thighRight: normalized.thighRight === "" ? "" : normalized.thighRight,
    thighLeft: normalized.thighLeft === "" ? "" : normalized.thighLeft,
  };
}

function normalizeActiveTicketCardSource_(value) {
  var normalized = normalizeText_(value).toLowerCase();
  return ["admin", "response", "customer"].indexOf(normalized) >= 0 ? normalized : "";
}

function normalizeTicketCardAcquiredAt_(value) {
  var normalized = normalizeText_(value);
  if (!normalized) return "";
  var parsed = new Date(normalized);
  return isNaN(parsed.getTime()) ? "" : parsed.toISOString();
}

function normalizeCustomerProfileRecord_(record, fallbackName) {
  var name = normalizeText_(record && record.name || fallbackName);
  if (!name) return null;
  var aliases = uniqueValues_(
    (Array.isArray(record && record.aliases) ? record.aliases : [])
      .map(normalizeText_)
      .filter(function (alias) {
        return alias && alias !== name;
      })
  );
  var clientIds = uniqueValues_(
    (Array.isArray(record && record.clientIds) ? record.clientIds : [])
      .map(normalizeText_)
      .filter(Boolean)
  );
  return {
    name: name,
    memberNumber: normalizeMemberNumber_(record && record.memberNumber),
    nameKana: normalizeKana_(record && record.nameKana),
    aliases: aliases,
    clientIds: clientIds,
    activeTicketCard: normalizeActiveTicketCard_(record && record.activeTicketCard),
    activeTicketCardSource: normalizeActiveTicketCardSource_(
      record && (record.activeTicketCardSource || record.ticketCardSource)
    ),
    lastTicketCardAcquiredAt: normalizeTicketCardAcquiredAt_(
      record && (record.lastTicketCardAcquiredAt || record.ticketCardLastAcquiredAt)
    ),
    measurementTargets: normalizeMeasurementTargets_(record && record.measurementTargets),
    pushStatus: normalizePushStatus_(record && record.pushStatus),
    rewardRedemptions: normalizeRewardRedemptions_(record && record.rewardRedemptions),
    adminManaged: record && record.adminManaged === true,
    updatedAt: normalizeText_(record && record.updatedAt) || new Date().toISOString(),
  };
}

function getCustomerProfiles_() {
  var stored = parseJson_(PropertiesService.getScriptProperties().getProperty(CUSTOMER_PROFILES_PROPERTY_KEY), {});
  var normalized = {};
  Object.keys(stored || {}).forEach(function (key) {
    var profile = normalizeCustomerProfileRecord_(stored[key], key);
    if (!profile) return;
    normalized[profile.name] = profile;
  });
  var withMemberNumbers = JSON.parse(JSON.stringify(normalized));
  var memberNumbersChanged = ensureCustomerProfileMemberNumbers_(withMemberNumbers);
  if (JSON.stringify(stored || {}) !== JSON.stringify(withMemberNumbers) || memberNumbersChanged) {
    saveCustomerProfiles_(withMemberNumbers);
  }
  return withMemberNumbers;
}

function saveCustomerProfiles_(profiles) {
  PropertiesService.getScriptProperties().setProperty(
    CUSTOMER_PROFILES_PROPERTY_KEY,
    JSON.stringify(profiles || {})
  );
}

function publicCustomerProfile_(record) {
  if (!record) return null;
  return {
    name: normalizeText_(record.name),
    memberNumber: normalizeMemberNumber_(record.memberNumber),
    nameKana: normalizeKana_(record.nameKana),
    activeTicketCard: publicActiveTicketCard_(record.activeTicketCard),
    activeTicketCardSource: normalizeActiveTicketCardSource_(record.activeTicketCardSource),
    lastTicketCardAcquiredAt: normalizeTicketCardAcquiredAt_(record.lastTicketCardAcquiredAt),
    measurementTargets: publicMeasurementTargets_(record.measurementTargets),
    pushStatus: publicPushStatus_(record.pushStatus),
    rewardRedemptions: publicRewardRedemptions_(record.rewardRedemptions),
    updatedAt: normalizeText_(record.updatedAt),
  };
}

function publicMeasurementRecord_(record) {
  if (!record) return null;
  var waist = normalizeMeasurementValue_(record.waist);
  var hip = normalizeMeasurementValue_(record.hip);
  var thighRight = normalizeMeasurementValue_(record.thighRight);
  var thighLeft = normalizeMeasurementValue_(record.thighLeft);
  var whr = "";
  if (waist !== "" && hip !== "" && Number(hip) > 0) {
    whr = Math.round((Number(waist) / Number(hip)) * 1000) / 1000;
  }
  return {
    id: normalizeText_(record.id),
    customerName: normalizeText_(record.customerName),
    memberNumber: normalizeMemberNumber_(record.memberNumber),
    measuredAt: normalizeText_(record.measuredAt),
    waist: waist,
    hip: hip,
    thighRight: thighRight,
    thighLeft: thighLeft,
    whr: whr,
    createdAt: normalizeText_(record.createdAt),
    updatedAt: normalizeText_(record.updatedAt),
    target: publicMeasurementTargets_(record.target),
  };
}

function getPublicMeasurementsForCustomer_(customerName) {
  var normalizedName = normalizeText_(customerName);
  if (!normalizedName) return [];
  return getMeasurements_({ customerName: normalizedName })
    .map(publicMeasurementRecord_)
    .filter(Boolean);
}

function pickCustomerProfileMatch_(matches, customerNameKana) {
  if (!matches || !matches.length) return null;
  var normalizedKana = normalizeKana_(customerNameKana);
  if (normalizedKana) {
    for (var i = 0; i < matches.length; i += 1) {
      if (normalizeKana_(matches[i].profile && matches[i].profile.nameKana) === normalizedKana) {
        return matches[i];
      }
    }
  }
  return matches[0];
}

function findCustomerProfileByClientId_(profiles, clientId) {
  var normalizedClientId = normalizeText_(clientId);
  if (!normalizedClientId) return null;
  var matches = [];
  Object.keys(profiles || {}).forEach(function (key) {
    var profile = normalizeCustomerProfileRecord_(profiles[key], key);
    if (!profile) return;
    if (profile.clientIds.indexOf(normalizedClientId) >= 0) {
      matches.push({ key: key, profile: profile });
    }
  });
  return matches.length ? matches[0] : null;
}

function findCustomerProfileByName_(profiles, customerName, customerNameKana) {
  var normalizedName = normalizeText_(customerName);
  if (!normalizedName) return null;
  var exactMatches = [];
  var aliasMatches = [];
  Object.keys(profiles || {}).forEach(function (key) {
    var profile = normalizeCustomerProfileRecord_(profiles[key], key);
    if (!profile) return;
    if (profile.name === normalizedName) {
      exactMatches.push({ key: key, profile: profile });
      return;
    }
    if (profile.aliases.indexOf(normalizedName) >= 0) {
      aliasMatches.push({ key: key, profile: profile });
    }
  });
  return pickCustomerProfileMatch_(exactMatches.length ? exactMatches : aliasMatches, customerNameKana);
}

function saveCustomerProfileRecord_(profiles, previousKey, record, options) {
  var normalized = normalizeCustomerProfileRecord_(record, record && record.name);
  if (!normalized) return null;
  var replaceActiveTicketCard = options && options.replaceActiveTicketCard === true;
  var replaceActiveTicketCardSource = options && options.replaceActiveTicketCardSource === true;
  var replaceLastTicketCardAcquiredAt = options && options.replaceLastTicketCardAcquiredAt === true;
  var replaceMeasurementTargets = options && options.replaceMeasurementTargets === true;
  var replacePushStatus = options && options.replacePushStatus === true;
  var replaceRewardRedemptions = options && options.replaceRewardRedemptions === true;
  var requestedActiveTicketCard = normalizeActiveTicketCard_(record && record.activeTicketCard);
  var requestedActiveTicketCardSource = normalizeActiveTicketCardSource_(
    record && (record.activeTicketCardSource || record.ticketCardSource)
  );
  var requestedLastTicketCardAcquiredAt = normalizeTicketCardAcquiredAt_(
    record && (record.lastTicketCardAcquiredAt || record.ticketCardLastAcquiredAt)
  );
  var requestedMeasurementTargets = normalizeMeasurementTargets_(record && record.measurementTargets);
  var requestedPushStatus = normalizePushStatus_(record && record.pushStatus);
  var requestedRewardRedemptions = normalizeRewardRedemptions_(record && record.rewardRedemptions);
  if (previousKey && previousKey !== normalized.name) {
    delete profiles[previousKey];
  }
  var existing = normalizeCustomerProfileRecord_(profiles[normalized.name], normalized.name);
  if (existing) {
    normalized = normalizeCustomerProfileRecord_({
      name: normalized.name,
      memberNumber: normalized.memberNumber || existing.memberNumber,
      nameKana: normalized.nameKana || existing.nameKana,
      aliases: existing.aliases.concat(normalized.aliases).concat(
        existing.name && existing.name !== normalized.name ? [existing.name] : []
      ),
      clientIds: existing.clientIds.concat(normalized.clientIds),
      adminManaged: existing.adminManaged || normalized.adminManaged,
      updatedAt: new Date().toISOString(),
    }, normalized.name);
  }
  if (replaceActiveTicketCard) {
    normalized.activeTicketCard = requestedActiveTicketCard;
  } else if (requestedActiveTicketCard) {
    normalized.activeTicketCard = requestedActiveTicketCard;
  } else if (existing && existing.activeTicketCard) {
    normalized.activeTicketCard = normalizeActiveTicketCard_(existing.activeTicketCard);
  }
  if (replaceActiveTicketCardSource) {
    normalized.activeTicketCardSource = requestedActiveTicketCardSource;
  } else if (requestedActiveTicketCardSource) {
    normalized.activeTicketCardSource = requestedActiveTicketCardSource;
  } else if (existing && existing.activeTicketCardSource) {
    normalized.activeTicketCardSource = normalizeActiveTicketCardSource_(existing.activeTicketCardSource);
  }
  if (replaceLastTicketCardAcquiredAt) {
    normalized.lastTicketCardAcquiredAt = requestedLastTicketCardAcquiredAt;
  } else if (requestedLastTicketCardAcquiredAt) {
    normalized.lastTicketCardAcquiredAt = requestedLastTicketCardAcquiredAt;
  } else if (existing && existing.lastTicketCardAcquiredAt) {
    normalized.lastTicketCardAcquiredAt = normalizeTicketCardAcquiredAt_(existing.lastTicketCardAcquiredAt);
  }
  if (replaceMeasurementTargets) {
    normalized.measurementTargets = requestedMeasurementTargets;
  } else if (requestedMeasurementTargets) {
    normalized.measurementTargets = requestedMeasurementTargets;
  } else if (existing && existing.measurementTargets) {
    normalized.measurementTargets = normalizeMeasurementTargets_(existing.measurementTargets);
  }
  if (replacePushStatus) {
    normalized.pushStatus = requestedPushStatus;
  } else if (requestedPushStatus) {
    normalized.pushStatus = requestedPushStatus;
  } else if (existing && existing.pushStatus) {
    normalized.pushStatus = normalizePushStatus_(existing.pushStatus);
  }
  if (replaceRewardRedemptions) {
    normalized.rewardRedemptions = requestedRewardRedemptions;
  } else if (requestedRewardRedemptions) {
    normalized.rewardRedemptions = requestedRewardRedemptions;
  } else if (existing && existing.rewardRedemptions) {
    normalized.rewardRedemptions = normalizeRewardRedemptions_(existing.rewardRedemptions);
  }
  if (!normalized.memberNumber) {
    var nextIndex = Math.max(getStoredNextMemberNumberIndex_(), getHighestMemberNumberIndex_(profiles) + 1, 1);
    normalized.memberNumber = formatMemberNumber_(nextIndex);
    saveNextMemberNumberIndex_(nextIndex + 1);
  } else {
    var memberNumberIndex = parseMemberNumberIndex_(normalized.memberNumber);
    if (memberNumberIndex > 0) {
      saveNextMemberNumberIndex_(Math.max(getStoredNextMemberNumberIndex_(), memberNumberIndex + 1));
    }
  }
  Object.keys(profiles || {}).forEach(function (key) {
    if (key === normalized.name) return;
    var profile = normalizeCustomerProfileRecord_(profiles[key], key);
    if (!profile) return;
    if (profile.memberNumber && profile.memberNumber === normalized.memberNumber) {
      throw new Error("会員番号が重複しています。");
    }
  });
  profiles[normalized.name] = normalized;
  return normalized;
}

function resolveCustomerProfileForSubmission_(customer, clientId) {
  var submittedName = normalizeText_(customer && customer.name);
  var submittedNameKana = normalizeKana_(customer && customer.nameKana);
  var normalizedClientId = normalizeText_(clientId);
  if (!submittedName && !normalizedClientId) return null;

  var profiles = getCustomerProfiles_();
  var clientMatch = findCustomerProfileByClientId_(profiles, normalizedClientId);
  var nameMatch = clientMatch ? null : findCustomerProfileByName_(profiles, submittedName, submittedNameKana);
  var match = clientMatch || nameMatch;
  var record = match
    ? normalizeCustomerProfileRecord_(match.profile, match.profile && match.profile.name)
    : normalizeCustomerProfileRecord_({
        name: submittedName,
        nameKana: submittedNameKana,
        clientIds: normalizedClientId ? [normalizedClientId] : [],
        aliases: [],
        adminManaged: false,
      }, submittedName);
  if (!record) return null;

  if (normalizedClientId && record.clientIds.indexOf(normalizedClientId) === -1) {
    record.clientIds.push(normalizedClientId);
  }
  if (submittedNameKana && (!record.nameKana || !record.adminManaged)) {
    record.nameKana = submittedNameKana;
  }
  if (submittedName && submittedName !== record.name) {
    if (record.adminManaged || !clientMatch) {
      if (record.aliases.indexOf(submittedName) === -1) {
        record.aliases.push(submittedName);
      }
    } else {
      if (record.name && record.aliases.indexOf(record.name) === -1) {
        record.aliases.push(record.name);
      }
      record.aliases = record.aliases.filter(function (alias) {
        return alias !== submittedName;
      });
      record.name = submittedName;
    }
  }
  record.updatedAt = new Date().toISOString();

  var saved = saveCustomerProfileRecord_(profiles, match && match.key, record);
  saveCustomerProfiles_(profiles);
  return saved;
}

function ensureCustomerProfileFromHistory_(canonicalName, customerNameKana, clientId, aliasName) {
  var name = normalizeText_(canonicalName);
  if (!name) return null;
  var normalizedClientId = normalizeText_(clientId);
  var normalizedAlias = normalizeText_(aliasName);
  var profiles = getCustomerProfiles_();
  var match = findCustomerProfileByClientId_(profiles, normalizedClientId) ||
    findCustomerProfileByName_(profiles, name, customerNameKana);
  var record = match
    ? normalizeCustomerProfileRecord_(match.profile, match.profile && match.profile.name)
    : normalizeCustomerProfileRecord_({
        name: name,
        nameKana: normalizeKana_(customerNameKana),
        clientIds: normalizedClientId ? [normalizedClientId] : [],
        aliases: [],
        adminManaged: false,
      }, name);
  if (!record) return null;

  if (normalizedClientId && record.clientIds.indexOf(normalizedClientId) === -1) {
    record.clientIds.push(normalizedClientId);
  }
  if (!record.nameKana && customerNameKana) {
    record.nameKana = normalizeKana_(customerNameKana);
  }
  if (normalizedAlias && normalizedAlias !== record.name && record.aliases.indexOf(normalizedAlias) === -1) {
    record.aliases.push(normalizedAlias);
  }
  record.updatedAt = new Date().toISOString();

  var saved = saveCustomerProfileRecord_(profiles, match && match.key, record);
  saveCustomerProfiles_(profiles);
  return saved;
}

function updateAdminCustomerProfileRecord_(currentName, nextName, responses, options) {
  var fromName = normalizeText_(currentName);
  var toName = normalizeText_(nextName);
  if (!fromName || !toName) return null;
  var normalizedOptions = options && typeof options === "object" ? options : {};
  var shouldReplaceActiveTicketCard = Object.prototype.hasOwnProperty.call(normalizedOptions, "activeTicketCard");
  var shouldReplaceActiveTicketCardSource = Object.prototype.hasOwnProperty.call(normalizedOptions, "activeTicketCardSource");
  var shouldReplaceMeasurementTargets = Object.prototype.hasOwnProperty.call(normalizedOptions, "measurementTargets");

  var profiles = getCustomerProfiles_();
  var match = findCustomerProfileByName_(profiles, fromName, "");
  var responseClientIds = uniqueValues_(
    (Array.isArray(responses) ? responses : [])
      .map(function (response) {
        return normalizeText_(response && response.customerClientId);
      })
      .filter(Boolean)
  );
  var record = match
    ? normalizeCustomerProfileRecord_(match.profile, match.profile && match.profile.name)
    : normalizeCustomerProfileRecord_({
        name: fromName,
        clientIds: responseClientIds,
        aliases: [],
        adminManaged: true,
      }, fromName);
  if (!record) return null;

  responseClientIds.forEach(function (clientId) {
    if (record.clientIds.indexOf(clientId) === -1) {
      record.clientIds.push(clientId);
    }
  });
  if (record.name && record.name !== toName && record.aliases.indexOf(record.name) === -1) {
    record.aliases.push(record.name);
  }
  if (fromName !== toName && record.aliases.indexOf(fromName) === -1) {
    record.aliases.push(fromName);
  }
  record.aliases = record.aliases.filter(function (alias) {
    return alias !== toName;
  });
  record.name = toName;
  record.adminManaged = true;
  record.updatedAt = new Date().toISOString();
  if (normalizeMemberNumber_(normalizedOptions.memberNumber)) {
    record.memberNumber = normalizeMemberNumber_(normalizedOptions.memberNumber);
  }
  if (normalizeKana_(normalizedOptions.nameKana)) {
    record.nameKana = normalizeKana_(normalizedOptions.nameKana);
  }
  if (shouldReplaceActiveTicketCard) {
    record.activeTicketCard = normalizeActiveTicketCard_(normalizedOptions.activeTicketCard);
  }
  if (shouldReplaceActiveTicketCardSource) {
    record.activeTicketCardSource = normalizeActiveTicketCardSource_(normalizedOptions.activeTicketCardSource);
  }
  if (shouldReplaceMeasurementTargets) {
    record.measurementTargets = normalizeMeasurementTargets_(normalizedOptions.measurementTargets);
  }

  var saved = saveCustomerProfileRecord_(profiles, match && match.key, record, {
    replaceActiveTicketCard: shouldReplaceActiveTicketCard,
    replaceActiveTicketCardSource: shouldReplaceActiveTicketCardSource,
    replaceMeasurementTargets: shouldReplaceMeasurementTargets,
  });
  saveCustomerProfiles_(profiles);
  return saved;
}

function getAdminCustomerProfiles_() {
  var profiles = getCustomerProfiles_();
  var completedCounts = getCompletedTicketCardCountsByCustomer_(getResponses_({}));
  return Object.keys(profiles || {})
    .map(function (key) {
      var publicProfile = publicCustomerProfile_(profiles[key]);
      if (!publicProfile) return null;
      publicProfile.completedTicketCardCount = completedCounts[publicProfile.name] || 0;
      return publicProfile;
    })
    .filter(Boolean)
    .sort(function (a, b) {
      var leftIndex = parseMemberNumberIndex_(a && a.memberNumber);
      var rightIndex = parseMemberNumberIndex_(b && b.memberNumber);
      if (leftIndex !== rightIndex) return leftIndex - rightIndex;
      return normalizeText_(a && a.name).localeCompare(normalizeText_(b && b.name));
    });
}

function syncCustomerProfileTicketCard_(customer, clientId, activeTicketCard, source) {
  var ticketCard = normalizeActiveTicketCard_(activeTicketCard);
  if (!ticketCard) return null;
  var profile = resolveCustomerProfileForSubmission_(customer, clientId);
  if (!profile) return null;
  var ticketCardSource = normalizeActiveTicketCardSource_(source) || "customer";

  var profiles = getCustomerProfiles_();
  var match = findCustomerProfileByClientId_(profiles, clientId) ||
    findCustomerProfileByName_(profiles, profile.name, profile.nameKana);
  var record = match
    ? normalizeCustomerProfileRecord_(match.profile, match.profile && match.profile.name)
    : normalizeCustomerProfileRecord_(profile, profile && profile.name);
  if (!record) return null;

  record.name = normalizeText_(profile.name || record.name);
  if (normalizeKana_(profile.nameKana)) {
    record.nameKana = normalizeKana_(profile.nameKana);
  }
  if (normalizeText_(clientId) && record.clientIds.indexOf(normalizeText_(clientId)) === -1) {
    record.clientIds.push(normalizeText_(clientId));
  }
  if (ticketCardSource === "customer" && record.lastTicketCardAcquiredAt) {
    var acquiredAt = new Date(record.lastTicketCardAcquiredAt);
    if (!isNaN(acquiredAt.getTime()) && Date.now() - acquiredAt.getTime() < TICKET_CARD_ACQUIRE_COOLDOWN_MS) {
      throw new Error("スタンプカードの取得は24時間に1回までです。24時間後にもう一度お試しください。");
    }
  }
  if (record.activeTicketCardSource === "admin" && ticketCardSource !== "admin") {
    record.activeTicketCard = normalizeActiveTicketCard_(record.activeTicketCard);
    record.activeTicketCardSource = "admin";
  } else {
    record.activeTicketCard = ticketCard;
    record.activeTicketCardSource = ticketCardSource;
  }
  if (ticketCardSource === "customer") {
    record.lastTicketCardAcquiredAt = new Date().toISOString();
  }
  record.updatedAt = new Date().toISOString();

  var saved = saveCustomerProfileRecord_(profiles, match && match.key, record, {
    replaceActiveTicketCard: true,
    replaceActiveTicketCardSource: true,
    replaceLastTicketCardAcquiredAt: ticketCardSource === "customer",
  });
  saveCustomerProfiles_(profiles);
  return saved;
}

function syncCustomerProfilePushStatus_(customer, clientId, pushStatus) {
  var normalizedPushStatus = normalizePushStatus_(pushStatus);
  if (!normalizedPushStatus) return null;
  var profile = resolveCustomerProfileForSubmission_(customer, clientId);
  if (!profile) return null;

  var profiles = getCustomerProfiles_();
  var match = findCustomerProfileByClientId_(profiles, normalizeText_(clientId)) ||
    findCustomerProfileByName_(profiles, profile.name, profile.nameKana);
  var record = match
    ? normalizeCustomerProfileRecord_(match.profile, match.key)
    : normalizeCustomerProfileRecord_(profile, profile && profile.name);
  if (!record) return null;

  record.pushStatus = normalizedPushStatus;
  record.updatedAt = new Date().toISOString();
  var saved = saveCustomerProfileRecord_(profiles, match && match.key, record, {
    replacePushStatus: true,
  });
  saveCustomerProfiles_(profiles);
  return saved;
}

function syncCustomerProfileTicketCardFromResponse_(response) {
  if (!response) return null;
  var ticketCard = buildActiveTicketCardFromAnswers_(response.answers);
  if (!ticketCard) return null;
  return syncCustomerProfileTicketCard_({
    name: response.customerName,
    nameKana: "",
  }, response.customerClientId, ticketCard, "response");
}

function deleteCustomerProfileRecord_(customerName) {
  var name = normalizeText_(customerName);
  if (!name) return false;
  var profiles = getCustomerProfiles_();
  var match = findCustomerProfileByName_(profiles, name, "");
  if (!match) return false;
  delete profiles[match.key];
  Object.keys(profiles).forEach(function (key) {
    var profile = normalizeCustomerProfileRecord_(profiles[key], key);
    if (!profile) return;
    if (profile.aliases.indexOf(name) === -1) return;
    profile.aliases = profile.aliases.filter(function (alias) {
      return alias !== name;
    });
    profiles[key] = profile;
  });
  saveCustomerProfiles_(profiles);
  return true;
}

function getCustomerHistoryPayload_(filter) {
  var normalizedFilter = {
    clientId: normalizeText_(filter && filter.clientId),
    customerName: normalizeText_(filter && filter.customerName),
    customerNameKana: normalizeKana_(filter && filter.customerNameKana),
    matchByNameOnly: Boolean(filter && filter.matchByNameOnly),
    includeTrashed: Boolean(filter && filter.includeTrashed),
  };
  var profileMatch = normalizedFilter.matchByNameOnly
    ? findCustomerProfileByName_(getCustomerProfiles_(), normalizedFilter.customerName, normalizedFilter.customerNameKana)
    : findCustomerProfileByClientId_(getCustomerProfiles_(), normalizedFilter.clientId) ||
      findCustomerProfileByName_(getCustomerProfiles_(), normalizedFilter.customerName, normalizedFilter.customerNameKana);
  var responses = [];
  var profile = null;

  if (profileMatch) {
    profile = ensureCustomerProfileFromHistory_(
      profileMatch.profile && profileMatch.profile.name,
      normalizedFilter.customerNameKana || profileMatch.profile && profileMatch.profile.nameKana,
      normalizedFilter.clientId,
      normalizedFilter.customerName
    );
    responses = getResponses_({
      customerName: profile && profile.name,
      includeTrashed: normalizedFilter.includeTrashed,
    });
    return {
      responses: responses,
      customerProfile: publicCustomerProfile_(profile),
      measurements: getPublicMeasurementsForCustomer_(profile && profile.name),
    };
  }

  responses = getResponses_({
    clientId: normalizedFilter.matchByNameOnly ? "" : normalizedFilter.clientId,
    customerName: normalizedFilter.customerName,
    matchByNameOnly: normalizedFilter.matchByNameOnly,
    includeTrashed: normalizedFilter.includeTrashed,
  });

  if (!responses.length && !normalizedFilter.matchByNameOnly && normalizedFilter.clientId) {
    responses = getResponses_({
      clientId: normalizedFilter.clientId,
      includeTrashed: normalizedFilter.includeTrashed,
    });
    if (responses.length) {
      profile = ensureCustomerProfileFromHistory_(
        responses[0].customerName,
        normalizedFilter.customerNameKana,
        normalizedFilter.clientId,
        normalizedFilter.customerName
      );
      responses = getResponses_({
        customerName: profile && profile.name || responses[0].customerName,
        includeTrashed: normalizedFilter.includeTrashed,
      });
    }
  } else if (responses.length) {
    profile = ensureCustomerProfileFromHistory_(
      responses[0].customerName,
      normalizedFilter.customerNameKana,
      normalizedFilter.clientId,
      normalizedFilter.customerName && normalizedFilter.customerName !== responses[0].customerName
        ? normalizedFilter.customerName
        : ""
    );
  } else if (normalizedFilter.customerName) {
    profile = ensureCustomerProfileFromHistory_(
      normalizedFilter.customerName,
      normalizedFilter.customerNameKana,
      normalizedFilter.clientId,
      ""
    );
  }

  return {
    responses: responses,
    customerProfile: publicCustomerProfile_(profile),
    measurements: getPublicMeasurementsForCustomer_(
      profile && profile.name || normalizedFilter.customerName
    ),
  };
}

function getAnswerValueByQuestionIds_(answers, questionIds) {
  var list = Array.isArray(answers) ? answers : [];
  for (var i = 0; i < list.length; i += 1) {
    if (questionIds.indexOf(String(list[i].questionId || "")) >= 0) {
      var value = normalizeText_(list[i].value);
      if (value) return value;
    }
  }
  return "";
}

function hasResponseTicketInfo_(response) {
  return Boolean(
    getAnswerValueByQuestionIds_(response && response.answers, CUSTOMER_TICKET_INFO_QUESTION_IDS.plan) ||
    getAnswerValueByQuestionIds_(response && response.answers, CUSTOMER_TICKET_INFO_QUESTION_IDS.sheet) ||
    getAnswerValueByQuestionIds_(response && response.answers, CUSTOMER_TICKET_INFO_QUESTION_IDS.round)
  );
}

function updateAnswerValueByQuestionIds_(answers, questionIds, value) {
  return (Array.isArray(answers) ? answers : []).map(function (answer) {
    if (questionIds.indexOf(String(answer && answer.questionId || "")) === -1) return answer;
    return Object.assign({}, answer, {
      value: normalizeText_(value),
    });
  });
}

function saveCustomerMemos_(memos) {
  PropertiesService.getScriptProperties().setProperty(CUSTOMER_MEMOS_PROPERTY_KEY, JSON.stringify(memos || {}));
}

function renameCustomerMemo_(currentName, nextName) {
  var fromName = normalizeText_(currentName);
  var toName = normalizeText_(nextName);
  if (!fromName || !toName || fromName === toName) return getCustomerMemos_();
  var memos = getCustomerMemos_();
  if (!memos[fromName]) return memos;
  var source = normalizeCustomerMemoRecord_(memos[fromName]);
  var target = normalizeCustomerMemoRecord_(memos[toName]);
  delete memos[fromName];
  memos[toName] = {
    latestMemo: source.latestMemo || target.latestMemo,
    entries: source.entries.concat(target.entries).slice(0, 100),
  };
  saveCustomerMemos_(memos);
  return memos;
}

function deleteCustomerMemo_(customerName) {
  var name = normalizeText_(customerName);
  if (!name) return getCustomerMemos_();
  var memos = getCustomerMemos_();
  delete memos[name];
  saveCustomerMemos_(memos);
  return memos;
}

function renameCustomerPhotoFolder_(currentName, nextName) {
  var fromName = sanitizeFolderName_(currentName) || "お名前未設定";
  var toName = sanitizeFolderName_(nextName) || "お名前未設定";
  if (!fromName || !toName || fromName === toName) return;
  var rootFolder = getRootPhotoFolder_();
  var sourceFolders = rootFolder.getFoldersByName(fromName);
  if (!sourceFolders.hasNext()) return;
  var targetFolders = rootFolder.getFoldersByName(toName);
  if (targetFolders.hasNext()) return;
  sourceFolders.next().setName(toName);
}

function trashCustomerPhotoFolder_(customerName) {
  var folderName = sanitizeFolderName_(customerName) || "お名前未設定";
  var rootFolder = getRootPhotoFolder_();
  var folders = rootFolder.getFoldersByName(folderName);
  while (folders.hasNext()) {
    try {
      folders.next().setTrashed(true);
    } catch (error) {
      // Ignore inaccessible folders.
    }
  }
}

function getLogs_() {
  return {
    auditLogs: getStoredLogs_(AUDIT_LOGS_PROPERTY_KEY),
    errorLogs: getStoredLogs_(ERROR_LOGS_PROPERTY_KEY),
  };
}

function getStoredLogs_(propertyKey) {
  return parseJson_(PropertiesService.getScriptProperties().getProperty(propertyKey), []);
}

function appendStoredLog_(propertyKey, entry) {
  var logs = getStoredLogs_(propertyKey);
  logs.unshift(entry);
  logs = logs.slice(0, MAX_LOG_ENTRIES);
  PropertiesService.getScriptProperties().setProperty(propertyKey, JSON.stringify(logs));
}

function appendAuditLog_(type, detail) {
  appendStoredLog_(AUDIT_LOGS_PROPERTY_KEY, {
    at: new Date().toISOString(),
    type: type,
    detail: detail || {},
  });
}

function appendErrorLog_(source, message, detail) {
  appendStoredLog_(ERROR_LOGS_PROPERTY_KEY, {
    at: new Date().toISOString(),
    source: source,
    message: message,
    detail: detail || {},
  });
}

function logClientError_(payload) {
  appendErrorLog_(
    normalizeText_(payload && payload.source) || "client",
    normalizeText_(payload && payload.message) || "不明なエラー",
    payload && payload.detail ? payload.detail : {}
  );
  return { ok: true };
}

function canAcceptSurveyResponses_(survey) {
  if (normalizeSurveyStatus_(survey && survey.status) !== "published") return false;
  if (survey && survey.acceptingResponses === false) return false;
  var nowTime = Date.now();
  if (survey && survey.startAt && new Date(survey.startAt).getTime() > nowTime) return false;
  if (survey && survey.endAt && new Date(survey.endAt).getTime() < nowTime) return false;
  return true;
}

function assertSurveyCanAcceptResponses_(survey) {
  if (!canAcceptSurveyResponses_(survey)) {
    throw new Error("このアンケートは現在受付していません。");
  }
}

function saveResponse_(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSpreadsheet_();
    var payload = body.payload || {};
    var responseId = body.responseId || payload.responseId || Utilities.getUuid();
    var existing = getResponseById_(responseId);
    if (existing) return { response: existing };

    var survey = findSurvey_(payload.surveyId);
    assertSurveyCanAcceptResponses_(survey);
    var customer = payload.customer || {};
    var customerClientId = body.clientId || payload.clientId || "";
    var customerProfile = resolveCustomerProfileForSubmission_(customer, customerClientId);
    var customerName = normalizeText_(customerProfile && customerProfile.name || customer.name);
    if (!customerName) throw new Error("お名前を入力してください。");

    var answers = buildAnswers_(survey, payload.answers || [], responseId, customerName);
    var duplicate = findDuplicateResponse_(survey, customerName, customerClientId, answers);
    if (duplicate) return { response: duplicate };
    var recentResponse = findLatestSurveyResponseWithinCooldown_(survey.id, customerName);
    if (recentResponse) {
      if (makeResponseSignature_(survey.id, customerName, recentResponse.answers) === makeResponseSignature_(survey.id, customerName, answers)) {
        return { response: recentResponse };
      }
      throw new Error(buildResponseResubmissionCooldownMessage_(recentResponse));
    }

    var response = {
      id: responseId,
      surveyId: survey.id,
      surveyTitle: survey.title,
      customerClientId: customerClientId,
      customerName: customerName,
      customerEmail: "",
      answers: answers,
      status: "new",
      adminMemo: "",
      submittedAt: new Date().toISOString(),
    };
    syncCustomerProfileTicketCardFromResponse_(response);
    appendMasterRow_(response);
    appendSurveyRow_(survey, response);
    notifyNewResponse_(response);
    appendAuditLog_("response.create", {
      responseId: response.id,
      surveyId: response.surveyId,
      customerName: response.customerName,
    });
    return { response: response };
  } catch (error) {
    appendErrorLog_("saveResponse", error.message || "保存エラー", {
      responseId: body && body.responseId,
      surveyId: body && body.payload && body.payload.surveyId,
    });
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function buildAnswers_(survey, rawAnswers, responseId, customerName) {
  var rawAnswerMap = buildRawAnswerMap_(rawAnswers);
  return survey.questions.map(function (question) {
    var visible = isQuestionVisible_(question, rawAnswerMap, survey);
    var required = isQuestionRequired_(question, visible, survey);
    var raw = findRawAnswer_(rawAnswers, question.id);
    if (question.type === "photo") {
      var requiredPhotoCount = getPhotoQuestionRequiredCount_(question, visible, survey);
      var files = visible
        ? syncPhotoFiles_(
            [],
            raw.files || [],
            responseId,
            question.id,
            customerName,
            getPhotoQuestionMaxFiles_(question, survey),
            survey,
            rawAnswerMap
          )
        : [];
      if (visible && requiredPhotoCount && files.length < requiredPhotoCount) {
        throw new Error(requiredPhotoCount === 1 ? "未回答の質問があります。" : "写真を" + requiredPhotoCount + "枚添付してください。");
      }
      return {
        questionId: question.id,
        label: question.label,
        type: question.type,
        value: files.length ? files.map(function (file) { return file.url; }).join(", ") : "",
        files: files,
      };
    }

    var values = Array.isArray(raw.value) ? raw.value : [raw.value];
    values = values.map(normalizeText_).filter(Boolean);
    if (required && !values.length) {
      throw new Error("未回答の質問があります。");
    }
    if (visible && question.type === "rating" && values[0] && ["1", "2", "3", "4", "5"].indexOf(values[0]) === -1) {
      throw new Error("評価は1から5で回答してください。");
    }
    if (visible && question.type === "choice" && values[0] && question.options.indexOf(values[0]) === -1) {
      throw new Error("選択肢から回答してください。");
    }
    if (visible && question.type === "checkbox" && values.some(function (item) { return question.options.indexOf(item) === -1; })) {
      throw new Error("選択肢から回答してください。");
    }
    return {
      questionId: question.id,
      label: question.label,
      type: question.type,
      value: visible ? values.join(", ") : "",
    };
  });
}

function findRawAnswer_(answers, questionId) {
  for (var i = 0; i < answers.length; i += 1) {
    if (answers[i].questionId === questionId) return answers[i];
  }
  return {};
}

function updatePublicResponse_(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var payload = body.payload || {};
    var responseId = body.responseId || payload.responseId || "";
    var existing = getResponseById_(responseId);
    if (!existing) throw new Error("回答が見つかりません。");
    if (!canCustomerEditResponse_(existing, body.clientId || payload.clientId || "", payload.customer && payload.customer.name)) {
      throw new Error("この回答は修正できません。");
    }

    var survey = findSurvey_(existing.surveyId);
    assertSurveyCanAcceptResponses_(survey);
    var updated = buildUpdatedPublicResponse_(existing, survey, payload.answers || []);
    var sheet = getSpreadsheet_().getSheetByName(MASTER_SHEET_NAME);
    var rowIndex = findMasterRowIndex_(responseId);
    if (!rowIndex) throw new Error("回答が見つかりません。");
    sheet.getRange(rowIndex, 10).setValue(JSON.stringify(updated.answers));
    sheet.getRange(rowIndex, 11).setValue(JSON.stringify(collectFilesFromAnswers_(updated.answers)));
    sheet.getRange(rowIndex, 12).setValue(new Date().toISOString());
    updateSurveySheetResponse_(survey, updated);
    syncCustomerProfileTicketCardFromResponse_(updated);
    appendAuditLog_("response.public_edit", {
      responseId: existing.id,
      surveyId: existing.surveyId,
      customerName: existing.customerName,
    });
    return { response: getResponseById_(responseId) };
  } catch (error) {
    appendErrorLog_("updatePublicResponse", error.message || "更新エラー", {
      responseId: body && body.responseId,
    });
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function updatePublicTicketCard_(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var customer = body.customer || body.payload && body.payload.customer || {};
    var clientId = body.clientId || body.payload && body.payload.clientId || "";
    var ticketCard = body.ticketCard || body.payload && body.payload.ticketCard || {};
    var saved = syncCustomerProfileTicketCard_(customer, clientId, ticketCard, "customer");
    if (!saved) throw new Error("スタンプカード情報を保存できませんでした。");
    appendAuditLog_("customer.ticket_card.public_update", {
      customerName: saved.name,
      ticketPlan: saved.activeTicketCard && saved.activeTicketCard.plan || "",
      ticketSheet: saved.activeTicketCard ? saved.activeTicketCard.sheetNumber + "枚目" : "",
      ticketRound: saved.activeTicketCard ? saved.activeTicketCard.round + "回目" : "",
    });
    return {
      ok: true,
      customerProfile: publicCustomerProfile_(saved),
    };
  } finally {
    lock.releaseLock();
  }
}

function updatePublicPushStatus_(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var customer = body.customer || body.payload && body.payload.customer || {};
    var clientId = body.clientId || body.payload && body.payload.clientId || "";
    var pushStatus = body.pushStatus || body.payload && body.payload.pushStatus || {};
    var saved = syncCustomerProfilePushStatus_(customer, clientId, pushStatus);
    if (!saved) throw new Error("通知設定を保存できませんでした。");
    appendAuditLog_("customer.push.public_update", {
      customerName: saved.name,
      enabled: saved.pushStatus && saved.pushStatus.enabled === true,
      supported: saved.pushStatus && saved.pushStatus.supported === true,
      permission: saved.pushStatus && saved.pushStatus.permission || "",
    });
    return {
      ok: true,
      customerProfile: publicCustomerProfile_(saved),
    };
  } finally {
    lock.releaseLock();
  }
}

function canCustomerEditResponse_(response, clientId, customerName) {
  if (!response) return false;
  if (response.customerClientId && String(response.customerClientId) !== String(clientId || "")) return false;
  var normalizedResponseName = normalizeText_(response.customerName);
  var normalizedCustomerName = normalizeText_(customerName);
  if (normalizedResponseName !== normalizedCustomerName) {
    var profileMatch = findCustomerProfileByClientId_(getCustomerProfiles_(), clientId) ||
      findCustomerProfileByName_(getCustomerProfiles_(), customerName, "");
    if (!profileMatch || normalizeText_(profileMatch.profile && profileMatch.profile.name) !== normalizedResponseName) {
      return false;
    }
  }
  return Date.now() - new Date(response.submittedAt).getTime() <= RESPONSE_EDIT_WINDOW_MS;
}

function buildUpdatedPublicResponse_(existing, survey, rawAnswers) {
  var answerMap = {};
  (Array.isArray(existing.answers) ? existing.answers : []).forEach(function (answer) {
    answerMap[answer.questionId] = answer;
  });
  var rawAnswerMap = buildRawAnswerMap_(rawAnswers);

  var answers = survey.questions.map(function (question) {
    var raw = findRawAnswer_(rawAnswers, question.id);
    var visible = isQuestionVisible_(question, rawAnswerMap, survey);
    var required = isQuestionRequired_(question, visible, survey);
    if (question.type === "photo") {
      var requiredPhotoCount = getPhotoQuestionRequiredCount_(question, visible, survey);
      var existingAnswer = answerMap[question.id] || { files: [] };
      var files = visible
        ? syncPhotoFiles_(
            existingAnswer.files || [],
            raw.files || [],
            existing.id,
            question.id,
            existing.customerName,
            getPhotoQuestionMaxFiles_(question, survey),
            survey,
            rawAnswerMap
          )
        : [];
      if (visible && requiredPhotoCount && files.length < requiredPhotoCount) {
        throw new Error(requiredPhotoCount === 1 ? "未回答の質問があります。" : "写真を" + requiredPhotoCount + "枚添付してください。");
      }
      return {
        questionId: question.id,
        label: question.label,
        type: question.type,
        value: files.length ? files.map(function (file) { return file.url; }).join(", ") : "",
        files: files,
      };
    }

    var values = Array.isArray(raw.value) ? raw.value : [raw.value];
    values = values.map(normalizeText_).filter(Boolean);
    var value = visible ? values.join(", ") : "";

    if (visible && required && !value) throw new Error("未回答の質問があります。");
    if (visible && question.type === "rating" && value && ["1", "2", "3", "4", "5"].indexOf(value) === -1) {
      throw new Error("評価は1から5で回答してください。");
    }
    if (visible && question.type === "choice" && value && question.options.indexOf(value) === -1) {
      throw new Error("選択肢から回答してください。");
    }
    if (visible && question.type === "checkbox" && values.some(function (item) { return question.options.indexOf(item) === -1; })) {
      throw new Error("選択肢から回答してください。");
    }

    return {
      questionId: question.id,
      label: question.label,
      type: question.type,
      value: value,
    };
  });

  return Object.assign({}, existing, {
    answers: answers,
    managedAt: new Date().toISOString(),
  });
}

function findDuplicateResponse_(survey, customerName, customerClientId, answers) {
  var signature = makeResponseSignature_(survey.id, customerName, answers);
  var nowTime = Date.now();
  var responses = getResponses_({});
  for (var i = 0; i < responses.length; i += 1) {
    var response = responses[i];
    if (response.surveyId !== survey.id) continue;
    if (normalizeText_(response.customerName) !== normalizeText_(customerName)) continue;
    if (String(response.customerClientId || "") !== String(customerClientId || "")) continue;
    if (nowTime - new Date(response.submittedAt).getTime() > DUPLICATE_RESPONSE_WINDOW_MS) continue;
    if (makeResponseSignature_(response.surveyId, response.customerName, response.answers) === signature) {
      return response;
    }
  }
  return null;
}

function findLatestSurveyResponseWithinCooldown_(surveyId, customerName) {
  var normalizedSurveyId = normalizeText_(surveyId);
  var normalizedCustomerName = normalizeText_(customerName);
  if (!normalizedSurveyId || !normalizedCustomerName) return null;
  var nowTime = Date.now();
  var responses = getResponses_({ customerName: normalizedCustomerName, matchByNameOnly: true });
  for (var i = 0; i < responses.length; i += 1) {
    var response = responses[i];
    if (normalizeText_(response.surveyId) !== normalizedSurveyId) continue;
    var submittedAt = new Date(response.submittedAt);
    if (isNaN(submittedAt.getTime())) continue;
    if (nowTime - submittedAt.getTime() <= RESPONSE_RESUBMIT_COOLDOWN_MS) {
      return response;
    }
  }
  return null;
}

function formatDateTimeJa_(value) {
  if (!value) return "";
  var date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
}

function buildResponseResubmissionCooldownMessage_(response) {
  var submittedAt = new Date(response && response.submittedAt);
  if (isNaN(submittedAt.getTime())) {
    return "このアンケートは送信後24時間たたないと再送信できません。";
  }
  var availableAt = new Date(submittedAt.getTime() + RESPONSE_RESUBMIT_COOLDOWN_MS);
  return "このアンケートは送信後24時間たたないと再送信できません。次は " +
    formatDateTimeJa_(availableAt) + " 以降に送信できます。";
}

function makeResponseSignature_(surveyId, customerName, answers) {
  return JSON.stringify({
    surveyId: surveyId,
    customerName: normalizeText_(customerName),
    answers: (Array.isArray(answers) ? answers : []).map(function (answer) {
      return {
        questionId: answer.questionId,
        value: normalizeText_(answer.value),
        files: Array.isArray(answer.files)
          ? answer.files.map(function (file) { return normalizeText_(file.fileId || file.name || file.url); })
          : [],
      };
    }),
  });
}

function getRawAnswerValue_(rawAnswerMap, questionId) {
  var values = rawAnswerMap && rawAnswerMap[questionId];
  return Array.isArray(values) && values.length ? normalizeText_(values[0]) : "";
}

function getQuestionPhotoBaseName_(question, survey) {
  if (!question) return "写真";
  var questionId = normalizeText_(question.id);
  if (
    [
      "q_bijiris_session_monitor_photos_6",
      "q_bijiris_session_monitor_photos_10",
      "q_bijiris_session_monitor_photos",
    ].indexOf(questionId) >= 0
  ) {
    return "モニター";
  }
  if (
    [
      "q_bijiris_session_ticket_end_photos_6",
      "q_bijiris_session_ticket_end_photos_10",
      "q_bijiris_session_ticket_end_photos",
      "q_ticket_end_photo_last",
    ].indexOf(questionId) >= 0
  ) {
    return "回数券終了";
  }
  var label = normalizeText_(question.label);
  if (!label) return "写真";
  label = label.replace(/写真\d*枚?/g, "");
  label = label.replace(/\s+/g, " ");
  return sanitizeFileName_(label || "写真");
}

function buildPhotoStorageContext_(survey, question, rawAnswerMap) {
  var surveyTitle = sanitizeFolderName_(survey && survey.title || "アンケート");
  var sessionType = getRawAnswerValue_(rawAnswerMap, "q_bijiris_session_type");
  var ticketPlan = getRawAnswerValue_(rawAnswerMap, "q_bijiris_session_ticket_plan");
  var ticketSheet = getRawAnswerValue_(rawAnswerMap, "q_bijiris_session_ticket_sheet");
  var legacyTicketPlan = getRawAnswerValue_(rawAnswerMap, "q_ticket_end_ticket_size");
  var legacyTicketSheet = getRawAnswerValue_(rawAnswerMap, "q_ticket_end_ticket_sheet");
  var folderLabel = surveyTitle || "アンケート";

  if (survey && survey.id === "survey_bijiris_session") {
    if (sessionType === "回数券") {
      folderLabel = sanitizeFolderName_("回数券" + ticketPlan + " " + ticketSheet) || folderLabel;
    } else if (sessionType) {
      folderLabel = sanitizeFolderName_(sessionType) || folderLabel;
    }
  } else if (survey && survey.id === "survey_bijiris_ticket_end") {
    folderLabel = sanitizeFolderName_("回数券" + legacyTicketPlan + " " + legacyTicketSheet) || folderLabel;
  }

  return {
    folderName: folderLabel || "アンケート",
    fileBaseName: getQuestionPhotoBaseName_(question, survey),
  };
}

function getChildFolderByName_(parentFolder, folderName) {
  var normalizedFolderName = sanitizeFolderName_(folderName) || "未分類";
  var folders = parentFolder.getFoldersByName(normalizedFolderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(normalizedFolderName);
}

function savePhotoFiles_(files, responseId, questionId, customerName, survey, rawAnswerMap) {
  if (!Array.isArray(files) || !files.length) return [];
  var customerFolder = getCustomerPhotoFolder_(customerName);
  var question = survey && Array.isArray(survey.questions)
    ? survey.questions.filter(function (item) { return item && item.id === questionId; })[0]
    : null;
  var storageContext = buildPhotoStorageContext_(survey, question, rawAnswerMap);
  var folder = getChildFolderByName_(customerFolder, storageContext.folderName);
  var folderName = folder.getName();
  var folderUrl = folder.getUrl();
  var customerFolderName = customerFolder.getName();
  var customerFolderUrl = customerFolder.getUrl();
  return files.slice(0, 6).map(function (file, index) {
    var dataUrl = String(file.dataUrl || "");
    var match = dataUrl.match(/^data:(image\/(?:jpeg|jpg|png|webp));base64,(.+)$/i);
    if (!match) throw new Error("写真データの形式が正しくありません。");

    var mimeType = match[1].replace("image/jpg", "image/jpeg");
    var extension = mimeType.split("/")[1].replace("jpeg", "jpg");
    var fileBaseName = sanitizeFileName_(storageContext.fileBaseName || "写真");
    var fileName = sanitizeFileName_(fileBaseName + String(index + 1).padStart(2, "0") + "." + extension);
    var blob = Utilities.newBlob(Utilities.base64Decode(match[2]), mimeType, fileName);
    var driveFile = folder.createFile(blob);
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = driveFile.getId();
    return {
      name: fileName,
      type: mimeType,
      capturedAt: normalizeText_(file.capturedAt),
      fileId: fileId,
      customerFolderName: customerFolderName,
      customerFolderUrl: customerFolderUrl,
      folderName: folderName,
      folderUrl: folderUrl,
      url: driveFile.getUrl(),
      previewUrl: "https://drive.google.com/uc?export=view&id=" + fileId,
      downloadUrl: "https://drive.google.com/uc?export=download&id=" + fileId,
      thumbnailUrl: "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1200",
    };
  });
}

function syncPhotoFiles_(existingFiles, nextFiles, responseId, questionId, customerName, maxFiles, survey, rawAnswerMap) {
  var currentFiles = Array.isArray(existingFiles) ? existingFiles : [];
  var requestedFiles = Array.isArray(nextFiles) ? nextFiles : [];
  var existingById = {};
  var keptIds = {};
  var keptFiles = [];
  var newFiles = [];
  var normalizedMaxFiles = Math.max(1, Number(maxFiles || 6));

  currentFiles.forEach(function (file) {
    if (file && file.fileId) existingById[String(file.fileId)] = file;
  });

  requestedFiles.forEach(function (file) {
    if (!file) return;
    if (file.dataUrl) {
      newFiles.push(file);
      return;
    }
    var fileId = normalizeText_(file.fileId);
    if (!fileId || !existingById[fileId] || keptIds[fileId]) return;
    keptIds[fileId] = true;
    keptFiles.push(existingById[fileId]);
  });

  if (keptFiles.length + newFiles.length > normalizedMaxFiles) {
    throw new Error("写真は" + normalizedMaxFiles + "枚まで添付できます。");
  }

  currentFiles.forEach(function (file) {
    var fileId = normalizeText_(file && file.fileId);
    if (!fileId || keptIds[fileId]) return;
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch (error) {
      // Ignore already-deleted or inaccessible files.
    }
  });

  return keptFiles.concat(savePhotoFiles_(newFiles, responseId, questionId, customerName, survey, rawAnswerMap));
}

function sanitizeFileName_(value) {
  return String(value).replace(/[\\/:*?"<>|#%{}]/g, "_").slice(0, 180);
}

function stripFileExtension_(value) {
  return String(value || "").replace(/\.[^.]+$/, "");
}

function sanitizeFolderName_(value) {
  return normalizeText_(value).replace(/[\\/:*?"<>|#%{}]/g, "_").slice(0, 120);
}

function getRootPhotoFolder_() {
  var properties = PropertiesService.getScriptProperties();
  var folderId = normalizeText_(properties.getProperty("PHOTO_ROOT_FOLDER_ID"));
  if (folderId) {
    try {
      var existingFolder = DriveApp.getFolderById(folderId);
      if (existingFolder.getName() !== ROOT_DRIVE_FOLDER_NAME) {
        try {
          existingFolder.setName(ROOT_DRIVE_FOLDER_NAME);
        } catch (error) {
          // Ignore rename failures and keep using the accessible folder.
        }
      }
      return existingFolder;
    } catch (error) {
      // Recreate the folder if the saved id is no longer accessible.
    }
  }

  var driveRoot = DriveApp.getRootFolder();
  var folders = driveRoot.getFoldersByName(ROOT_DRIVE_FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : null;
  if (!folder) {
    for (var i = 0; i < LEGACY_ROOT_DRIVE_FOLDER_NAMES.length; i += 1) {
      var legacyFolders = driveRoot.getFoldersByName(LEGACY_ROOT_DRIVE_FOLDER_NAMES[i]);
      if (!legacyFolders.hasNext()) continue;
      folder = legacyFolders.next();
      try {
        folder.setName(ROOT_DRIVE_FOLDER_NAME);
      } catch (error) {
        // Keep using the legacy folder name if rename is not allowed.
      }
      break;
    }
  }
  if (!folder) {
    folder = driveRoot.createFolder(ROOT_DRIVE_FOLDER_NAME);
  }
  properties.setProperty("PHOTO_ROOT_FOLDER_ID", folder.getId());
  return folder;
}

function getCustomerPhotoFolder_(customerName) {
  var rootFolder = getRootPhotoFolder_();
  var folderName = sanitizeFolderName_(customerName) || "お名前未設定";
  var folders = rootFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : rootFolder.createFolder(folderName);
}

function getChildFolderByName_(parentFolder, folderName) {
  var normalizedName = sanitizeFolderName_(folderName) || "名称未設定";
  var folders = parentFolder.getFoldersByName(normalizedName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(normalizedName);
}

function getAnalysisSheetRootFolder_() {
  return getChildFolderByName_(getRootPhotoFolder_(), "分析シート");
}

function uploadAnalysisSheetFile_(payload) {
  var customerName = normalizeText_(payload && payload.customerName);
  var fileName = normalizeText_(payload && payload.fileName);
  var base64Data = normalizeText_(payload && payload.base64Data);
  var mimeType = normalizeText_(payload && payload.mimeType) || MimeType.PDF;
  var replaceExisting = payload && payload.replaceExisting === false ? false : true;
  if (!customerName) throw new Error("顧客名が必要です。");
  if (!fileName) throw new Error("ファイル名が必要です。");
  if (!base64Data) throw new Error("ファイルデータが必要です。");

  var analysisRoot = getAnalysisSheetRootFolder_();
  var customerFolder = getChildFolderByName_(analysisRoot, customerName);

  if (replaceExisting) {
    var existingFiles = customerFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      try {
        existingFiles.next().setTrashed(true);
      } catch (error) {
        // Ignore inaccessible files.
      }
    }
  }

  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  var file = customerFolder.createFile(blob);
  appendAuditLog_("analysis_sheet.upload", {
    customerName: customerName,
    fileName: fileName,
    fileId: file.getId(),
  });
  return {
    customerName: customerName,
    fileName: file.getName(),
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    folderId: customerFolder.getId(),
    folderUrl: customerFolder.getUrl(),
  };
}

function normalizeBijirisPostStatus_(value) {
  var status = normalizeText_(value);
  return ["published", "draft", "archived"].indexOf(status) >= 0 ? status : "draft";
}

function normalizeBijirisPostAttachmentRecord_(record, kind, index) {
  var attachmentKind = normalizeText_(record && record.kind || kind);
  var fileId = normalizeText_(record && record.fileId);
  var url = normalizeText_(record && record.url);
  var name = normalizeText_(record && record.name);
  if (!fileId && !url && !name) return null;
  return {
    kind: attachmentKind || "photo",
    sortOrder: Number(record && record.sortOrder || index || 0),
    name: name || ("添付" + String(Number(index || 0) + 1)),
    title: normalizeText_(record && record.title) || (attachmentKind === "pdf" ? name.replace(/\.pdf$/i, "") : ""),
    type: normalizeText_(record && record.type),
    fileId: fileId,
    url: url,
    previewUrl: normalizeText_(record && record.previewUrl),
    downloadUrl: normalizeText_(record && record.downloadUrl),
    thumbnailUrl: normalizeText_(record && record.thumbnailUrl),
    thumbnailFileId: normalizeText_(record && record.thumbnailFileId),
  };
}

function normalizeBijirisPostRecord_(record, fallbackId) {
  var id = normalizeText_(record && record.id || fallbackId);
  var title = normalizeText_(record && record.title);
  if (!id || !title) return null;
  var photos = (Array.isArray(record && record.photos) ? record.photos : [])
    .map(function (file, index) {
      return normalizeBijirisPostAttachmentRecord_(file, "photo", index);
    })
    .filter(Boolean);
  var documents = (Array.isArray(record && record.documents) ? record.documents : [])
    .map(function (file, index) {
      return normalizeBijirisPostAttachmentRecord_(file, "pdf", index);
    })
    .filter(Boolean);
  return {
    id: id,
    title: title,
    category: normalizeText_(record && record.category),
    summary: normalizeText_(record && record.summary),
    body: normalizeText_(record && record.body),
    status: normalizeBijirisPostStatus_(record && record.status),
    pinned: record && record.pinned === true,
    createdAt: normalizeText_(record && record.createdAt) || new Date().toISOString(),
    updatedAt: normalizeText_(record && record.updatedAt) || new Date().toISOString(),
    publishedAt: normalizeText_(record && record.publishedAt),
    photos: photos,
    documents: documents,
  };
}

function ensureBijirisPostsSheet_() {
  return ensureSheet_(getSpreadsheet_(), BIJIRIS_POSTS_SHEET_NAME, BIJIRIS_POST_HEADERS);
}

function ensureBijirisPostAttachmentsSheet_() {
  return ensureSheet_(getSpreadsheet_(), BIJIRIS_POST_ATTACHMENTS_SHEET_NAME, BIJIRIS_POST_ATTACHMENT_HEADERS);
}

function buildBijirisPostRowValues_(post) {
  return [
    post.createdAt || "",
    post.updatedAt || "",
    post.publishedAt || "",
    post.id || "",
    post.title || "",
    post.category || "",
    post.summary || "",
    post.body || "",
    post.status || "draft",
    post.pinned ? "1" : "",
  ];
}

function buildBijirisPostAttachmentRowValues_(postId, file) {
  return [
    postId || "",
    file.kind || "",
    Number(file.sortOrder || 0),
    file.name || "",
    file.type || "",
    file.fileId || "",
    file.url || "",
    file.previewUrl || "",
    file.downloadUrl || "",
    file.thumbnailUrl || "",
    file.title || "",
    file.thumbnailFileId || "",
  ];
}

function readBijirisPostRows_() {
  var sheet = ensureBijirisPostsSheet_();
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, BIJIRIS_POST_HEADERS.length).getValues();
  return values
    .map(function (row) {
      return normalizeBijirisPostRecord_({
        createdAt: stringifyDate_(row[0]),
        updatedAt: stringifyDate_(row[1]),
        publishedAt: stringifyDate_(row[2]),
        id: String(row[3] || ""),
        title: String(row[4] || ""),
        category: String(row[5] || ""),
        summary: String(row[6] || ""),
        body: String(row[7] || ""),
        status: String(row[8] || ""),
        pinned: String(row[9] || "") === "1",
      }, row[3]);
    })
    .filter(Boolean);
}

function readBijirisPostAttachmentRows_() {
  var sheet = ensureBijirisPostAttachmentsSheet_();
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, BIJIRIS_POST_ATTACHMENT_HEADERS.length).getValues();
  return values
    .map(function (row) {
      var attachment = normalizeBijirisPostAttachmentRecord_({
        kind: String(row[1] || ""),
        sortOrder: row[2],
        name: String(row[3] || ""),
        type: String(row[4] || ""),
        fileId: String(row[5] || ""),
        url: String(row[6] || ""),
        previewUrl: String(row[7] || ""),
        downloadUrl: String(row[8] || ""),
        thumbnailUrl: String(row[9] || ""),
        title: String(row[10] || ""),
        thumbnailFileId: String(row[11] || ""),
      }, row[1], row[2]);
      if (!attachment) return null;
      attachment.postId = String(row[0] || "");
      return attachment;
    })
    .filter(Boolean);
}

function publicBijirisPost_(post) {
  var normalized = normalizeBijirisPostRecord_(post, post && post.id);
  if (!normalized) return null;
  return {
    id: normalized.id,
    title: normalized.title,
    category: normalized.category,
    summary: normalized.summary,
    body: normalized.body,
    status: normalized.status,
    pinned: normalized.pinned,
    createdAt: normalized.createdAt,
    updatedAt: normalized.updatedAt,
    publishedAt: normalized.publishedAt,
    photos: normalized.photos,
    documents: normalized.documents,
  };
}

function compareBijirisPostRecords_(a, b) {
  if (Boolean(b && b.pinned) !== Boolean(a && a.pinned)) {
    return b && b.pinned ? 1 : -1;
  }
  var aTime = new Date(a && (a.publishedAt || a.updatedAt || a.createdAt) || 0).getTime();
  var bTime = new Date(b && (b.publishedAt || b.updatedAt || b.createdAt) || 0).getTime();
  return bTime - aTime;
}

function getBijirisPosts_(filter) {
  ensureSpreadsheet_();
  var attachmentsByPostId = {};
  readBijirisPostAttachmentRows_().forEach(function (attachment) {
    var postId = normalizeText_(attachment && attachment.postId);
    if (!postId) return;
    if (!attachmentsByPostId[postId]) {
      attachmentsByPostId[postId] = { photos: [], documents: [] };
    }
    if (attachment.kind === "pdf") {
      attachmentsByPostId[postId].documents.push(attachment);
    } else {
      attachmentsByPostId[postId].photos.push(attachment);
    }
  });

  return readBijirisPostRows_()
    .map(function (post) {
      var attachments = attachmentsByPostId[post.id] || { photos: [], documents: [] };
      return publicBijirisPost_(Object.assign({}, post, {
        photos: attachments.photos.sort(function (a, b) { return Number(a.sortOrder || 0) - Number(b.sortOrder || 0); }),
        documents: attachments.documents.sort(function (a, b) { return Number(a.sortOrder || 0) - Number(b.sortOrder || 0); }),
      }));
    })
    .filter(function (post) {
      if (!post) return false;
      if (filter && filter.publishedOnly) return post.status === "published";
      if (filter && filter.includeDrafts) return true;
      return post.status !== "archived";
    })
    .sort(compareBijirisPostRecords_);
}

function rewriteBijirisPostsSheets_(posts) {
  var normalizedPosts = (Array.isArray(posts) ? posts : [])
    .map(function (post) { return normalizeBijirisPostRecord_(post, post && post.id); })
    .filter(Boolean);
  var postSheet = ensureBijirisPostsSheet_();
  var attachmentSheet = ensureBijirisPostAttachmentsSheet_();
  if (postSheet.getLastRow() > 1) {
    postSheet.deleteRows(2, postSheet.getLastRow() - 1);
  }
  if (attachmentSheet.getLastRow() > 1) {
    attachmentSheet.deleteRows(2, attachmentSheet.getLastRow() - 1);
  }
  if (normalizedPosts.length) {
    postSheet.getRange(2, 1, normalizedPosts.length, BIJIRIS_POST_HEADERS.length).setValues(
      normalizedPosts.map(buildBijirisPostRowValues_)
    );
  }
  var attachmentRows = [];
  normalizedPosts.forEach(function (post) {
    (post.photos || []).forEach(function (file, index) {
      attachmentRows.push(buildBijirisPostAttachmentRowValues_(post.id, Object.assign({}, file, {
        kind: "photo",
        sortOrder: index,
      })));
    });
    (post.documents || []).forEach(function (file, index) {
      attachmentRows.push(buildBijirisPostAttachmentRowValues_(post.id, Object.assign({}, file, {
        kind: "pdf",
        sortOrder: index,
      })));
    });
  });
  if (attachmentRows.length) {
    attachmentSheet.getRange(2, 1, attachmentRows.length, BIJIRIS_POST_ATTACHMENT_HEADERS.length).setValues(attachmentRows);
  }
}

function getBijirisPostsRootFolder_() {
  return getChildFolderByName_(getRootPhotoFolder_(), BIJIRIS_POSTS_FOLDER_NAME);
}

function getBijirisPostFolder_(postId) {
  return getChildFolderByName_(getBijirisPostsRootFolder_(), "post_" + sanitizeFolderName_(postId));
}

function trashDriveFilesByRecords_(files) {
  (Array.isArray(files) ? files : []).forEach(function (file) {
    [normalizeText_(file && file.fileId), normalizeText_(file && file.thumbnailFileId)].forEach(function (fileId) {
      if (!fileId) return;
      try {
        DriveApp.getFileById(fileId).setTrashed(true);
      } catch (error) {
        // Ignore already-deleted or inaccessible files.
      }
    });
  });
}

function saveBijirisPdfCoverFile_(dataUrl, postId, fileName, existingThumbnailFileId) {
  var normalizedDataUrl = String(dataUrl || "");
  if (!normalizedDataUrl) {
    if (existingThumbnailFileId) {
      try {
        DriveApp.getFileById(existingThumbnailFileId).setTrashed(true);
      } catch (error) {
        // Ignore already-deleted or inaccessible files.
      }
    }
    return { thumbnailUrl: "", thumbnailFileId: "" };
  }
  var match = normalizedDataUrl.match(/^data:(image\/(?:jpeg|jpg|png|webp));base64,(.+)$/i);
  if (!match) throw new Error("PDF 表紙画像の形式が正しくありません。");
  if (existingThumbnailFileId) {
    try {
      DriveApp.getFileById(existingThumbnailFileId).setTrashed(true);
    } catch (error) {
      // Ignore already-deleted or inaccessible files.
    }
  }
  var folder = getChildFolderByName_(getBijirisPostFolder_(postId), "pdf-covers");
  var mimeType = match[1].replace("image/jpg", "image/jpeg");
  var extension = mimeType.split("/")[1].replace("jpeg", "jpg");
  var coverName = sanitizeFileName_(stripFileExtension_(fileName) + "_cover." + extension);
  var blob = Utilities.newBlob(Utilities.base64Decode(match[2]), mimeType, coverName);
  var driveFile = folder.createFile(blob);
  driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var coverFileId = driveFile.getId();
  return {
    thumbnailUrl: "https://drive.google.com/thumbnail?id=" + coverFileId + "&sz=w1200",
    thumbnailFileId: coverFileId,
  };
}

function saveBijirisPostImageFiles_(files, postId) {
  var folder = getChildFolderByName_(getBijirisPostFolder_(postId), "photos");
  return (Array.isArray(files) ? files : []).slice(0, 8).map(function (file, index) {
    var dataUrl = String(file && file.dataUrl || "");
    var match = dataUrl.match(/^data:(image\/(?:jpeg|jpg|png|webp));base64,(.+)$/i);
    if (!match) throw new Error("写真データの形式が正しくありません。");
    var mimeType = match[1].replace("image/jpg", "image/jpeg");
    var extension = mimeType.split("/")[1].replace("jpeg", "jpg");
    var fileName = sanitizeFileName_(normalizeText_(file && file.name) || ("photo_" + String(index + 1).padStart(2, "0") + "." + extension));
    var blob = Utilities.newBlob(Utilities.base64Decode(match[2]), mimeType, fileName);
    var driveFile = folder.createFile(blob);
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = driveFile.getId();
    return {
      kind: "photo",
      sortOrder: index,
      name: fileName,
      type: mimeType,
      fileId: fileId,
      url: driveFile.getUrl(),
      previewUrl: "https://drive.google.com/uc?export=view&id=" + fileId,
      downloadUrl: "https://drive.google.com/uc?export=download&id=" + fileId,
      thumbnailUrl: "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1200",
    };
  });
}

function saveBijirisPostDocumentFiles_(files, postId) {
  var folder = getChildFolderByName_(getBijirisPostFolder_(postId), "pdf");
  return (Array.isArray(files) ? files : []).slice(0, 5).map(function (file, index) {
    var base64Data = normalizeText_(file && file.base64Data);
    var mimeType = normalizeText_(file && (file.mimeType || file.type)) || MimeType.PDF;
    if (!base64Data) throw new Error("PDF データの形式が正しくありません。");
    if (mimeType !== MimeType.PDF && mimeType !== "application/pdf") {
      throw new Error("PDF のみ添付できます。");
    }
    var fileName = sanitizeFileName_(normalizeText_(file && file.name) || ("document_" + String(index + 1).padStart(2, "0") + ".pdf"));
    if (!/\.pdf$/i.test(fileName)) fileName += ".pdf";
    var displayTitle = normalizeText_(file && file.title) || stripFileExtension_(fileName);
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.PDF, fileName);
    var driveFile = folder.createFile(blob);
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = driveFile.getId();
    var coverDataUrl = normalizeText_(file && file.thumbnailDataUrl);
    if (!coverDataUrl && /^data:image\//i.test(normalizeText_(file && file.thumbnailUrl))) {
      coverDataUrl = normalizeText_(file && file.thumbnailUrl);
    }
    var coverMeta = saveBijirisPdfCoverFile_(coverDataUrl, postId, fileName, "");
    return {
      kind: "pdf",
      sortOrder: index,
      name: fileName,
      title: displayTitle,
      type: MimeType.PDF,
      fileId: fileId,
      url: driveFile.getUrl(),
      previewUrl: "https://drive.google.com/file/d/" + fileId + "/preview",
      downloadUrl: "https://drive.google.com/uc?export=download&id=" + fileId,
      thumbnailUrl: coverMeta.thumbnailUrl,
      thumbnailFileId: coverMeta.thumbnailFileId,
    };
  });
}

function syncBijirisPostFiles_(existingFiles, nextFiles, postId, kind) {
  var currentFiles = Array.isArray(existingFiles) ? existingFiles : [];
  var requestedFiles = Array.isArray(nextFiles) ? nextFiles : [];
  var existingById = {};
  var keptIds = {};
  var keptFiles = [];
  var newFiles = [];

  currentFiles.forEach(function (file) {
    var fileId = normalizeText_(file && file.fileId);
    if (fileId) existingById[fileId] = file;
  });

  requestedFiles.forEach(function (file) {
    var fileId = normalizeText_(file && file.fileId);
    if ((file && file.dataUrl) || (file && file.base64Data)) {
      newFiles.push(file);
      return;
    }
    if (!fileId || !existingById[fileId] || keptIds[fileId]) return;
    keptIds[fileId] = true;
    if (kind === "pdf") {
      var existingFile = existingById[fileId];
      var nextTitle = normalizeText_(file && file.title) || existingFile.title || stripFileExtension_(existingFile.name);
      var nextName = existingFile.name || file.name || "";
      if (nextTitle && (!existingFile.title || existingFile.title !== nextTitle)) {
        try {
          var nextDriveName = sanitizeFileName_(nextTitle);
          if (!/\.pdf$/i.test(nextDriveName)) nextDriveName += ".pdf";
          DriveApp.getFileById(fileId).setName(nextDriveName);
          nextName = nextDriveName;
        } catch (error) {
          // Ignore rename failures and keep current name.
        }
      }
      var nextThumbnailUrl = existingFile.thumbnailUrl || "";
      var nextThumbnailFileId = existingFile.thumbnailFileId || "";
      var nextThumbnailDataUrl = normalizeText_(file && file.thumbnailDataUrl);
      if (!nextThumbnailDataUrl && /^data:image\//i.test(normalizeText_(file && file.thumbnailUrl))) {
        nextThumbnailDataUrl = normalizeText_(file && file.thumbnailUrl);
      }
      if (file && file.thumbnailRemoved === true) {
        var removedCoverMeta = saveBijirisPdfCoverFile_("", postId, nextName, existingFile.thumbnailFileId || "");
        nextThumbnailUrl = removedCoverMeta.thumbnailUrl;
        nextThumbnailFileId = removedCoverMeta.thumbnailFileId;
      } else if (nextThumbnailDataUrl) {
        var updatedCoverMeta = saveBijirisPdfCoverFile_(nextThumbnailDataUrl, postId, nextName, existingFile.thumbnailFileId || "");
        nextThumbnailUrl = updatedCoverMeta.thumbnailUrl;
        nextThumbnailFileId = updatedCoverMeta.thumbnailFileId;
      } else if (normalizeText_(file && file.thumbnailUrl)) {
        nextThumbnailUrl = normalizeText_(file && file.thumbnailUrl);
      }
      keptFiles.push(Object.assign({}, existingFile, {
        name: nextName,
        title: nextTitle,
        thumbnailUrl: nextThumbnailUrl,
        thumbnailFileId: nextThumbnailFileId,
      }));
      return;
    }
    keptFiles.push(existingById[fileId]);
  });

  currentFiles.forEach(function (file) {
    var fileId = normalizeText_(file && file.fileId);
    if (!fileId || keptIds[fileId]) return;
    trashDriveFilesByRecords_([file]);
  });

  return keptFiles.concat(
    kind === "pdf"
      ? saveBijirisPostDocumentFiles_(newFiles, postId)
      : saveBijirisPostImageFiles_(newFiles, postId)
  );
}

function createBijirisPost_(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSpreadsheet_();
    var now = new Date().toISOString();
    var status = normalizeBijirisPostStatus_(payload && payload.status);
    var postId = normalizeText_(payload && payload.id) || makeId_("bijiris");
    var record = normalizeBijirisPostRecord_({
      id: postId,
      title: payload && payload.title,
      category: payload && payload.category,
      summary: payload && payload.summary,
      body: payload && payload.body,
      status: status,
      pinned: payload && payload.pinned === true,
      createdAt: now,
      updatedAt: now,
      publishedAt: status === "published" ? normalizeText_(payload && payload.publishedAt) || now : "",
      photos: [],
      documents: [],
    }, postId);
    if (!record) throw new Error("タイトルを入力してください。");
    record.photos = syncBijirisPostFiles_([], payload && payload.photos, record.id, "photo");
    record.documents = syncBijirisPostFiles_([], payload && payload.documents, record.id, "pdf");
    if (!record.summary && !record.body && !record.photos.length && !record.documents.length) {
      throw new Error("本文、要約、写真、PDF のいずれかを入力してください。");
    }
    var posts = getBijirisPosts_({ includeDrafts: true }).concat([record]);
    rewriteBijirisPostsSheets_(posts);
    notifyBijirisPostIfNeeded_(record, payload, "create");
    appendAuditLog_("bijiris_post.create", {
      postId: record.id,
      title: record.title,
      status: record.status,
    });
    return publicBijirisPost_(record);
  } finally {
    lock.releaseLock();
  }
}

function updateBijirisPost_(postId, payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSpreadsheet_();
    var posts = getBijirisPosts_({ includeDrafts: true });
    var existing = posts.filter(function (post) { return post.id === normalizeText_(postId); })[0];
    if (!existing) throw new Error("投稿が見つかりません。");
    var now = new Date().toISOString();
    var status = normalizeBijirisPostStatus_(Object.prototype.hasOwnProperty.call(payload || {}, "status") ? payload.status : existing.status);
    var updated = normalizeBijirisPostRecord_({
      id: existing.id,
      title: Object.prototype.hasOwnProperty.call(payload || {}, "title") ? payload.title : existing.title,
      category: Object.prototype.hasOwnProperty.call(payload || {}, "category") ? payload.category : existing.category,
      summary: Object.prototype.hasOwnProperty.call(payload || {}, "summary") ? payload.summary : existing.summary,
      body: Object.prototype.hasOwnProperty.call(payload || {}, "body") ? payload.body : existing.body,
      status: status,
      pinned: Object.prototype.hasOwnProperty.call(payload || {}, "pinned") ? payload.pinned === true : existing.pinned,
      createdAt: existing.createdAt,
      updatedAt: now,
      publishedAt: status === "published" ? (normalizeText_(payload && payload.publishedAt) || existing.publishedAt || now) : existing.publishedAt,
      photos: [],
      documents: [],
    }, existing.id);
    if (!updated) throw new Error("タイトルを入力してください。");
    updated.photos = syncBijirisPostFiles_(
      existing.photos,
      Object.prototype.hasOwnProperty.call(payload || {}, "photos") ? payload.photos : existing.photos,
      updated.id,
      "photo"
    );
    updated.documents = syncBijirisPostFiles_(
      existing.documents,
      Object.prototype.hasOwnProperty.call(payload || {}, "documents") ? payload.documents : existing.documents,
      updated.id,
      "pdf"
    );
    if (!updated.summary && !updated.body && !updated.photos.length && !updated.documents.length) {
      throw new Error("本文、要約、写真、PDF のいずれかを入力してください。");
    }
    rewriteBijirisPostsSheets_(
      posts.map(function (post) {
        return post.id === updated.id ? updated : post;
      })
    );
    notifyBijirisPostIfNeeded_(updated, payload, "update");
    appendAuditLog_("bijiris_post.update", {
      postId: updated.id,
      title: updated.title,
      status: updated.status,
    });
    return publicBijirisPost_(updated);
  } finally {
    lock.releaseLock();
  }
}

function deleteBijirisPost_(postId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSpreadsheet_();
    var posts = getBijirisPosts_({ includeDrafts: true });
    var existing = posts.filter(function (post) { return post.id === normalizeText_(postId); })[0];
    if (!existing) throw new Error("投稿が見つかりません。");
    trashDriveFilesByRecords_((existing.photos || []).concat(existing.documents || []));
    rewriteBijirisPostsSheets_(
      posts.filter(function (post) { return post.id !== existing.id; })
    );
    appendAuditLog_("bijiris_post.delete", {
      postId: existing.id,
      title: existing.title,
    });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function replaceBijirisPosts_(posts) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSpreadsheet_();
    rewriteBijirisPostsSheets_(Array.isArray(posts) ? posts : []);
    appendAuditLog_("bijiris_post.replace", {
      count: Array.isArray(posts) ? posts.length : 0,
    });
    return getBijirisPosts_({ includeDrafts: true });
  } finally {
    lock.releaseLock();
  }
}

function normalizeMeasurementRecord_(record, fallbackId) {
  var id = normalizeText_(record && record.id || fallbackId);
  var customerName = normalizeText_(record && record.customerName);
  var measuredAt = normalizeMeasurementDate_(record && record.measuredAt);
  if (!id || !customerName || !measuredAt) return null;
  var waist = normalizeMeasurementValue_(record && record.waist);
  var hip = normalizeMeasurementValue_(record && record.hip);
  var thighRight = normalizeMeasurementValue_(record && record.thighRight);
  var thighLeft = normalizeMeasurementValue_(record && record.thighLeft);
  return {
    id: id,
    customerName: customerName,
    memberNumber: normalizeMemberNumber_(record && record.memberNumber),
    measuredAt: measuredAt,
    waist: waist,
    hip: hip,
    thighRight: thighRight,
    thighLeft: thighLeft,
    whr: computeMeasurementWhr_(waist, hip),
    memo: normalizeText_(record && record.memo),
    createdAt: normalizeText_(record && record.createdAt) || new Date().toISOString(),
    updatedAt: normalizeText_(record && record.updatedAt) || new Date().toISOString(),
  };
}

function buildMeasurementRowValues_(record) {
  return [
    record.createdAt || "",
    record.updatedAt || "",
    record.id || "",
    record.customerName || "",
    record.memberNumber || "",
    record.measuredAt || "",
    record.waist === "" ? "" : record.waist,
    record.hip === "" ? "" : record.hip,
    record.thighRight === "" ? "" : record.thighRight,
    record.thighLeft === "" ? "" : record.thighLeft,
    record.whr === "" ? "" : record.whr,
    record.memo || "",
  ];
}

function hasMeasurementValues_(record) {
  return ["waist", "hip", "thighRight", "thighLeft"].some(function (key) {
    return record && record[key] !== "";
  });
}

function ensureMeasurementsSheet_() {
  return ensureSheet_(getSpreadsheet_(), MEASUREMENTS_SHEET_NAME, MEASUREMENT_HEADERS);
}

function readMeasurementRows_() {
  var sheet = ensureMeasurementsSheet_();
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, MEASUREMENT_HEADERS.length).getValues();
  return values
    .map(function (row) {
      return normalizeMeasurementRecord_({
        createdAt: stringifyDate_(row[0]),
        updatedAt: stringifyDate_(row[1]),
        id: String(row[2] || ""),
        customerName: String(row[3] || ""),
        memberNumber: String(row[4] || ""),
        measuredAt: row[5],
        waist: row[6],
        hip: row[7],
        thighRight: row[8],
        thighLeft: row[9],
        whr: row[10],
        memo: String(row[11] || ""),
      }, row[2]);
    })
    .filter(Boolean);
}

function appendMeasurementRow_(record) {
  ensureMeasurementsSheet_().appendRow(buildMeasurementRowValues_(record));
}

function writeMeasurementRow_(rowIndex, record) {
  var sheet = ensureMeasurementsSheet_();
  if (!sheet || !rowIndex) return;
  sheet.getRange(rowIndex, 1, 1, MEASUREMENT_HEADERS.length).setValues([buildMeasurementRowValues_(record)]);
}

function findMeasurementRowIndex_(measurementId) {
  return findRowIndexByColumn_(ensureMeasurementsSheet_(), 3, measurementId);
}

function getMeasurementById_(measurementId) {
  var list = readMeasurementRows_();
  for (var i = 0; i < list.length; i += 1) {
    if (list[i].id === measurementId) return list[i];
  }
  return null;
}

function decorateMeasurementWithCustomerProfile_(measurement, profiles) {
  var profileMatch = findCustomerProfileByName_(profiles || {}, measurement && measurement.customerName, "");
  var profile = profileMatch && profileMatch.profile
    ? normalizeCustomerProfileRecord_(profileMatch.profile, profileMatch.key)
    : null;
  return Object.assign({}, measurement, {
    memberNumber: normalizeMemberNumber_(measurement && measurement.memberNumber) ||
      normalizeMemberNumber_(profile && profile.memberNumber),
    target: publicMeasurementTargets_(profile && profile.measurementTargets),
  });
}

function getMeasurements_(filter) {
  ensureSpreadsheet_();
  var profiles = getCustomerProfiles_();
  return readMeasurementRows_()
    .filter(function (measurement) {
      return !filter || !filter.customerName || measurement.customerName === normalizeText_(filter.customerName);
    })
    .map(function (measurement) {
      return decorateMeasurementWithCustomerProfile_(measurement, profiles);
    })
    .sort(function (a, b) {
      var measuredDiff = new Date(b.measuredAt).getTime() - new Date(a.measuredAt).getTime();
      if (measuredDiff !== 0) return measuredDiff;
      return new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime();
    });
}

function touchCustomerProfileUpdatedAt_(customerName) {
  var name = normalizeText_(customerName);
  if (!name) return null;
  var profiles = getCustomerProfiles_();
  var match = findCustomerProfileByName_(profiles, name, "");
  if (!match) return null;
  var record = normalizeCustomerProfileRecord_(match.profile, match.key);
  if (!record) return null;
  record.updatedAt = new Date().toISOString();
  var saved = saveCustomerProfileRecord_(profiles, match.key, record, {
    replaceActiveTicketCard: true,
    replaceMeasurementTargets: true,
  });
  saveCustomerProfiles_(profiles);
  return saved;
}

function renameMeasurementsForCustomer_(currentName, nextName, memberNumber) {
  var fromName = normalizeText_(currentName);
  var toName = normalizeText_(nextName);
  if (!fromName || !toName) return 0;
  var normalizedMemberNumber = normalizeMemberNumber_(memberNumber);
  var rows = readMeasurementRows_();
  var updated = 0;
  rows.forEach(function (measurement) {
    if (measurement.customerName !== fromName) return;
    measurement.customerName = toName;
    if (normalizedMemberNumber) {
      measurement.memberNumber = normalizedMemberNumber;
    }
    measurement.updatedAt = new Date().toISOString();
    writeMeasurementRow_(findMeasurementRowIndex_(measurement.id), measurement);
    updated += 1;
  });
  if (updated) touchCustomerProfileUpdatedAt_(toName);
  return updated;
}

function syncMeasurementMemberNumberForCustomer_(customerName, memberNumber) {
  var name = normalizeText_(customerName);
  var normalizedMemberNumber = normalizeMemberNumber_(memberNumber);
  if (!name || !normalizedMemberNumber) return 0;
  var rows = readMeasurementRows_();
  var updated = 0;
  rows.forEach(function (measurement) {
    if (measurement.customerName !== name) return;
    if (normalizeMemberNumber_(measurement.memberNumber) === normalizedMemberNumber) return;
    measurement.memberNumber = normalizedMemberNumber;
    measurement.updatedAt = new Date().toISOString();
    writeMeasurementRow_(findMeasurementRowIndex_(measurement.id), measurement);
    updated += 1;
  });
  if (updated) touchCustomerProfileUpdatedAt_(name);
  return updated;
}

function deleteMeasurementsByCustomer_(customerName) {
  var name = normalizeText_(customerName);
  if (!name) return 0;
  var sheet = ensureMeasurementsSheet_();
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var rows = readMeasurementRows_();
  var deleted = 0;
  for (var i = rows.length - 1; i >= 0; i -= 1) {
    if (rows[i].customerName !== name) continue;
    var rowIndex = findMeasurementRowIndex_(rows[i].id);
    if (!rowIndex) continue;
    sheet.deleteRow(rowIndex);
    deleted += 1;
  }
  return deleted;
}

function createMeasurement_(customerName, payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSpreadsheet_();
    var normalizedCustomerName = normalizeText_(customerName || payload && payload.customerName);
    if (!normalizedCustomerName) throw new Error("顧客名が必要です。");
    var profiles = getCustomerProfiles_();
    var profileMatch = findCustomerProfileByName_(profiles, normalizedCustomerName, "");
    var profile = profileMatch && profileMatch.profile
      ? normalizeCustomerProfileRecord_(profileMatch.profile, profileMatch.key)
      : null;
    var responses = getResponses_({ includeTrashed: true }).filter(function (response) {
      return normalizeText_(response.customerName) === normalizedCustomerName;
    });
    if (!profile && !responses.length) throw new Error("顧客が見つかりません。");

    var record = normalizeMeasurementRecord_({
      id: normalizeText_(payload && payload.id) || Utilities.getUuid(),
      customerName: profile && profile.name || normalizedCustomerName,
      memberNumber: normalizeMemberNumber_(payload && payload.memberNumber) ||
        normalizeMemberNumber_(profile && profile.memberNumber),
      measuredAt: payload && payload.measuredAt,
      waist: payload && payload.waist,
      hip: payload && payload.hip,
      thighRight: payload && payload.thighRight,
      thighLeft: payload && payload.thighLeft,
      memo: payload && payload.memo,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
    if (!record) throw new Error("測定日と顧客情報を確認してください。");
    if (!hasMeasurementValues_(record)) throw new Error("測定値を1つ以上入力してください。");
    appendMeasurementRow_(record);
    touchCustomerProfileUpdatedAt_(record.customerName);
    appendAuditLog_("measurement.create", {
      measurementId: record.id,
      customerName: record.customerName,
      measuredAt: record.measuredAt,
    });
    return decorateMeasurementWithCustomerProfile_(record, getCustomerProfiles_());
  } finally {
    lock.releaseLock();
  }
}

function updateMeasurement_(measurementId, payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSpreadsheet_();
    var existing = getMeasurementById_(measurementId);
    if (!existing) throw new Error("測定データが見つかりません。");
    var profileMatch = findCustomerProfileByName_(getCustomerProfiles_(), existing.customerName, "");
    var profile = profileMatch && profileMatch.profile
      ? normalizeCustomerProfileRecord_(profileMatch.profile, profileMatch.key)
      : null;
    var updated = normalizeMeasurementRecord_({
      id: existing.id,
      customerName: existing.customerName,
      memberNumber: normalizeMemberNumber_(payload && payload.memberNumber) ||
        normalizeMemberNumber_(profile && profile.memberNumber) ||
        existing.memberNumber,
      measuredAt: payload && payload.measuredAt || existing.measuredAt,
      waist: Object.prototype.hasOwnProperty.call(payload || {}, "waist") ? payload.waist : existing.waist,
      hip: Object.prototype.hasOwnProperty.call(payload || {}, "hip") ? payload.hip : existing.hip,
      thighRight: Object.prototype.hasOwnProperty.call(payload || {}, "thighRight") ? payload.thighRight : existing.thighRight,
      thighLeft: Object.prototype.hasOwnProperty.call(payload || {}, "thighLeft") ? payload.thighLeft : existing.thighLeft,
      memo: Object.prototype.hasOwnProperty.call(payload || {}, "memo") ? payload.memo : existing.memo,
      createdAt: existing.createdAt,
      updatedAt: new Date().toISOString(),
    });
    if (!updated) throw new Error("測定データを更新できませんでした。");
    if (!hasMeasurementValues_(updated)) throw new Error("測定値を1つ以上入力してください。");
    writeMeasurementRow_(findMeasurementRowIndex_(measurementId), updated);
    touchCustomerProfileUpdatedAt_(updated.customerName);
    appendAuditLog_("measurement.update", {
      measurementId: updated.id,
      customerName: updated.customerName,
      measuredAt: updated.measuredAt,
    });
    return decorateMeasurementWithCustomerProfile_(updated, getCustomerProfiles_());
  } finally {
    lock.releaseLock();
  }
}

function deleteMeasurement_(measurementId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSpreadsheet_();
    var existing = getMeasurementById_(measurementId);
    if (!existing) throw new Error("測定データが見つかりません。");
    var rowIndex = findMeasurementRowIndex_(measurementId);
    if (!rowIndex) throw new Error("測定データが見つかりません。");
    ensureMeasurementsSheet_().deleteRow(rowIndex);
    touchCustomerProfileUpdatedAt_(existing.customerName);
    appendAuditLog_("measurement.delete", {
      measurementId: existing.id,
      customerName: existing.customerName,
    });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function replaceMeasurements_(measurements) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSpreadsheet_();
    var sheet = ensureMeasurementsSheet_();
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
    var normalized = (Array.isArray(measurements) ? measurements : [])
      .map(function (measurement) {
        return normalizeMeasurementRecord_({
          id: measurement && measurement.id,
          customerName: measurement && measurement.customerName,
          memberNumber: measurement && measurement.memberNumber,
          measuredAt: measurement && measurement.measuredAt,
          waist: measurement && measurement.waist,
          hip: measurement && measurement.hip,
          thighRight: measurement && measurement.thighRight,
          thighLeft: measurement && measurement.thighLeft,
          memo: measurement && measurement.memo,
          createdAt: measurement && measurement.createdAt,
          updatedAt: measurement && measurement.updatedAt,
        }, measurement && measurement.id);
      })
      .filter(Boolean);
    if (normalized.length) {
      sheet.getRange(2, 1, normalized.length, MEASUREMENT_HEADERS.length).setValues(
        normalized.map(buildMeasurementRowValues_)
      );
    }
    appendAuditLog_("measurement.replace", {
      count: normalized.length,
    });
    return getMeasurements_({});
  } finally {
    lock.releaseLock();
  }
}

function decorateResponseWithCustomerProfile_(response, profiles) {
  var profileMatch = findCustomerProfileByClientId_(profiles || {}, response && response.customerClientId) ||
    findCustomerProfileByName_(profiles || {}, response && response.customerName, "");
  var profile = profileMatch && profileMatch.profile
    ? normalizeCustomerProfileRecord_(profileMatch.profile, profileMatch.key)
    : null;
  return Object.assign({}, response, {
    customerMemberNumber: normalizeMemberNumber_(profile && profile.memberNumber),
    customerNameKana: normalizeKana_(profile && profile.nameKana),
  });
}

function getResponses_(filter) {
  ensureSpreadsheet_();
  var rows = readMasterRows_();
  var profiles = getCustomerProfiles_();
  return rows
    .filter(function (response) {
      var matchByNameOnly = Boolean(filter && filter.matchByNameOnly);
      var clientMatches = matchByNameOnly || !filter.clientId || response.customerClientId === filter.clientId;
      return clientMatches &&
        (!filter.customerName || response.customerName === String(filter.customerName)) &&
        (filter.includeTrashed ? true : response.status !== "trash");
    })
    .map(function (response) {
      return decorateResponseWithCustomerProfile_(response, profiles);
    })
    .sort(function (a, b) {
      return new Date(b.submittedAt).getTime() - new Date(a.submittedAt).getTime();
    });
}

function getResponseById_(responseId) {
  var rows = readMasterRows_();
  var profiles = getCustomerProfiles_();
  for (var i = 0; i < rows.length; i += 1) {
    if (rows[i].id === responseId) return decorateResponseWithCustomerProfile_(rows[i], profiles);
  }
  return null;
}

function readMasterRows_() {
  var sheet = getSpreadsheet_().getSheetByName(MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, MASTER_HEADERS.length).getValues();
  return values
    .filter(function (row) { return row[1]; })
    .map(function (row) {
      return {
        submittedAt: stringifyDate_(row[0]),
        id: String(row[1]),
        surveyId: String(row[2]),
        surveyTitle: String(row[3]),
        customerClientId: String(row[4]),
        customerName: String(row[5]),
        customerEmail: String(row[6]),
        status: normalizeStatus_(row[7]),
        adminMemo: String(row[8] || ""),
        answers: parseJson_(row[9], []),
        files: parseJson_(row[10], []),
        managedAt: stringifyDate_(row[11]),
      };
    });
}

function appendMasterRow_(response) {
  var sheet = getSpreadsheet_().getSheetByName(MASTER_SHEET_NAME);
  var files = [];
  response.answers.forEach(function (answer) {
    if (Array.isArray(answer.files)) files = files.concat(answer.files);
  });
  sheet.appendRow([
    response.submittedAt,
    response.id,
    response.surveyId,
    response.surveyTitle,
    response.customerClientId,
    response.customerName,
    response.customerEmail,
    response.status,
    response.adminMemo,
    JSON.stringify(response.answers),
    JSON.stringify(files),
    "",
  ]);
}

function writeMasterResponseRow_(rowIndex, response) {
  var sheet = getSpreadsheet_().getSheetByName(MASTER_SHEET_NAME);
  if (!sheet || !rowIndex) return;
  sheet.getRange(rowIndex, 1, 1, MASTER_HEADERS.length).setValues([[
    response.submittedAt,
    response.id,
    response.surveyId,
    response.surveyTitle,
    response.customerClientId,
    response.customerName,
    response.customerEmail,
    response.status,
    response.adminMemo,
    JSON.stringify(response.answers || []),
    JSON.stringify(collectFilesFromAnswers_(response.answers || [])),
    response.managedAt || "",
  ]]);
}

function appendSurveyRow_(survey, response) {
  var sheet = ensureSurveySheet_(survey);
  var answerMap = {};
  response.answers.forEach(function (answer) {
    answerMap[answer.questionId] = answer.value || "";
  });
  var row = [
    response.submittedAt,
    response.id,
    response.customerClientId,
    response.customerName,
    response.customerEmail,
    response.status,
    response.adminMemo,
  ];
  survey.questions.forEach(function (question) {
    row.push(answerMap[question.id] || "");
  });
  sheet.appendRow(row);
}

function updateResponse_(responseId, status, adminMemo, answers) {
  ensureSpreadsheet_();
  var rowIndex = findMasterRowIndex_(responseId);
  if (!rowIndex) throw new Error("回答が見つかりません。");
  var existing = getResponseById_(responseId);
  if (!existing) throw new Error("回答が見つかりません。");
  var survey = findSurvey_(existing.surveyId);
  var normalizedStatus = normalizeStatus_(status);
  var memo = normalizeText_(adminMemo);
  var normalizedAnswers = normalizeAdminAnswers_(survey, existing.answers, answers);
  var managedAt = new Date().toISOString();
  writeMasterResponseRow_(rowIndex, Object.assign({}, existing, {
    status: normalizedStatus,
    adminMemo: memo,
    answers: normalizedAnswers,
    managedAt: managedAt,
  }));

  updateSurveySheetResponse_(survey, Object.assign({}, existing, {
    status: normalizedStatus,
    adminMemo: memo,
    answers: normalizedAnswers,
  }));
  syncCustomerProfileTicketCardFromResponse_(Object.assign({}, existing, {
    status: normalizedStatus,
    adminMemo: memo,
    answers: normalizedAnswers,
  }));
  appendAuditLog_("response.admin_update", {
    responseId: responseId,
    surveyId: existing.surveyId,
    customerName: existing.customerName,
    status: normalizedStatus,
  });
  return getResponseById_(responseId);
}

function deleteResponse_(responseId) {
  return trashResponse_(responseId);
}

function updateCustomerProfile_(customerName, payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSpreadsheet_();
    var currentName = normalizeText_(customerName);
    var nextName = normalizeText_(payload && payload.name);
    var memberNumber = normalizeMemberNumber_(payload && payload.memberNumber);
    var ticketPlan = normalizeText_(payload && payload.ticketPlan);
    var ticketSheet = normalizeText_(payload && payload.ticketSheet);
    var ticketRound = normalizeText_(payload && payload.ticketRound);
    var shouldReplaceMeasurementTargets = Boolean(
      payload && Object.prototype.hasOwnProperty.call(payload, "measurementTargets")
    );
    var measurementTargets = normalizeMeasurementTargets_(payload && payload.measurementTargets);
    if (!currentName) throw new Error("顧客名が必要です。");
    if (!nextName) throw new Error("お名前を入力してください。");

    var shouldUpdateTicket = Boolean(ticketPlan || ticketSheet || ticketRound);
    if (shouldUpdateTicket && !(ticketPlan && ticketSheet && ticketRound)) {
      throw new Error("回数券情報は種類・何枚目・何回目をすべて入力してください。");
    }

    var profileMatch = findCustomerProfileByName_(getCustomerProfiles_(), currentName, "");
    var responses = getResponses_({ includeTrashed: true }).filter(function (response) {
      return normalizeText_(response.customerName) === currentName;
    });
    if (!responses.length && !profileMatch) throw new Error("顧客が見つかりません。");

    var latestTicketResponse = responses
      .filter(function (response) {
        return response.status !== "trash" && hasResponseTicketInfo_(response);
      })
      .sort(function (a, b) {
        return new Date(b.submittedAt).getTime() - new Date(a.submittedAt).getTime();
      })[0];

    responses.forEach(function (response) {
      var nextAnswers = Array.isArray(response.answers) ? response.answers.slice() : [];
      if (shouldUpdateTicket && latestTicketResponse && response.id === latestTicketResponse.id) {
        nextAnswers = updateAnswerValueByQuestionIds_(nextAnswers, CUSTOMER_TICKET_INFO_QUESTION_IDS.plan, ticketPlan);
        nextAnswers = updateAnswerValueByQuestionIds_(nextAnswers, CUSTOMER_TICKET_INFO_QUESTION_IDS.sheet, ticketSheet);
        nextAnswers = updateAnswerValueByQuestionIds_(nextAnswers, CUSTOMER_TICKET_INFO_QUESTION_IDS.round, ticketRound);
      }
      var updated = Object.assign({}, response, {
        customerName: nextName,
        answers: nextAnswers,
        managedAt: new Date().toISOString(),
      });
      writeMasterResponseRow_(findMasterRowIndex_(response.id), updated);
      updateSurveySheetResponse_(findSurvey_(response.surveyId), updated);
    });

    if (currentName !== nextName) {
      renameCustomerMemo_(currentName, nextName);
      renameCustomerPhotoFolder_(currentName, nextName);
    }
    var profileUpdateOptions = {};
    if (memberNumber) {
      profileUpdateOptions.memberNumber = memberNumber;
    }
    if (shouldUpdateTicket) {
      profileUpdateOptions.activeTicketCard = {
        plan: ticketPlan,
        sheetLabel: ticketSheet,
        roundLabel: ticketRound,
      };
      profileUpdateOptions.activeTicketCardSource = "admin";
    }
    if (shouldReplaceMeasurementTargets) {
      profileUpdateOptions.measurementTargets = measurementTargets;
    }
    var savedProfile = updateAdminCustomerProfileRecord_(currentName, nextName, responses, profileUpdateOptions);
    var updatedMeasurements = 0;
    if (savedProfile && currentName !== nextName) {
      updatedMeasurements = renameMeasurementsForCustomer_(currentName, savedProfile.name, savedProfile.memberNumber);
    } else if (savedProfile && savedProfile.memberNumber) {
      updatedMeasurements = syncMeasurementMemberNumberForCustomer_(savedProfile.name, savedProfile.memberNumber);
    }

    appendAuditLog_("customer.profile.update", {
      customerName: currentName,
      nextCustomerName: nextName,
      updatedResponses: responses.length,
      ticketUpdated: shouldUpdateTicket,
      memberNumber: memberNumber,
      measurementTargetsUpdated: shouldReplaceMeasurementTargets,
      updatedMeasurements: updatedMeasurements,
    });
    return {
      ok: true,
      customerName: nextName,
      customerProfile: publicCustomerProfile_(savedProfile),
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteCustomerProfile_(customerName) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSpreadsheet_();
    var name = normalizeText_(customerName);
    if (!name) throw new Error("顧客名が必要です。");
    var hasProfile = Boolean(findCustomerProfileByName_(getCustomerProfiles_(), name, ""));
    var responses = getResponses_({ includeTrashed: true }).filter(function (response) {
      return normalizeText_(response.customerName) === name;
    });
    var hadMemo = Boolean(getCustomerMemos_()[name]);
    if (!responses.length && !hadMemo && !hasProfile) throw new Error("顧客が見つかりません。");

    responses.forEach(function (response) {
      purgeResponse_(response.id);
    });
    var deletedMeasurements = deleteMeasurementsByCustomer_(name);
    deleteCustomerMemo_(name);
    deleteCustomerProfileRecord_(name);
    trashCustomerPhotoFolder_(name);

    appendAuditLog_("customer.profile.delete", {
      customerName: name,
      deletedResponses: responses.length,
      deletedMeasurements: deletedMeasurements,
    });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function trashResponse_(responseId) {
  ensureSpreadsheet_();
  var response = getResponseById_(responseId);
  if (!response) throw new Error("回答が見つかりません。");
  updateResponse_(responseId, "trash", response.adminMemo, response.answers);
  appendAuditLog_("response.trash", {
    responseId: responseId,
    surveyId: response.surveyId,
    customerName: response.customerName,
  });
  return getResponseById_(responseId);
}

function purgeResponse_(responseId) {
  ensureSpreadsheet_();
  var response = getResponseById_(responseId);
  if (!response) return false;
  deleteResponseFiles_(response);
  deleteRowByResponseId_(getSpreadsheet_().getSheetByName(response.surveyTitle), responseId, 2);
  deleteRowByResponseId_(getSpreadsheet_().getSheetByName(MASTER_SHEET_NAME), responseId, 2);
  appendAuditLog_("response.purge", {
    responseId: responseId,
    surveyId: response.surveyId,
    customerName: response.customerName,
  });
  return true;
}

function deleteResponseFiles_(response) {
  var files = Array.isArray(response.files) ? response.files : [];
  files.forEach(function (file) {
    var fileId = normalizeText_(file.fileId);
    if (!fileId) return;
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch (error) {
      // Ignore already-deleted or inaccessible files.
    }
  });
}

function collectFilesFromAnswers_(answers) {
  var files = [];
  (Array.isArray(answers) ? answers : []).forEach(function (answer) {
    if (Array.isArray(answer.files)) files = files.concat(answer.files);
  });
  return files;
}

function normalizeAdminAnswers_(survey, existingAnswers, answers) {
  var existingMap = {};
  (Array.isArray(existingAnswers) ? existingAnswers : []).forEach(function (answer) {
    existingMap[answer.questionId] = answer;
  });

  var answerMap = {};
  (Array.isArray(answers) ? answers : []).forEach(function (answer) {
    answerMap[answer.questionId] = answer;
  });

  var rawAnswerMap = buildRawAnswerMap_(answers);
  return survey.questions.map(function (question) {
    var existing = existingMap[question.id] || {
      questionId: question.id,
      label: question.label,
      type: question.type,
      value: "",
    };
    var visible = isQuestionVisible_(question, rawAnswerMap, survey);
    var required = isQuestionRequired_(question, visible, survey);
    if (question.type === "photo") {
      return visible ? existing : Object.assign({}, existing, { value: "", files: [] });
    }

    var raw = answerMap[question.id] || {};
    var values = Array.isArray(raw.value) ? raw.value : [raw.value];
    values = values.map(normalizeText_).filter(Boolean);
    var value = visible ? values.join(", ") : "";

    if (required && !value) {
      throw new Error("未回答の質問があります。");
    }
    if (visible && question.type === "rating" && value && ["1", "2", "3", "4", "5"].indexOf(value) === -1) {
      throw new Error("評価は1から5で回答してください。");
    }
    if (visible && question.type === "choice" && value && question.options.indexOf(value) === -1) {
      throw new Error("選択肢から回答してください。");
    }
    if (visible && question.type === "checkbox" && values.some(function (item) { return question.options.indexOf(item) === -1; })) {
      throw new Error("選択肢から回答してください。");
    }

    return {
      questionId: question.id,
      label: question.label,
      type: question.type,
      value: value,
    };
  });
}

function updateSurveySheetManagement_(sheetName, responseId, status, adminMemo) {
  var sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) return;
  var rowIndex = findRowIndexByColumn_(sheet, 2, responseId);
  if (!rowIndex) return;
  sheet.getRange(rowIndex, 6).setValue(status);
  sheet.getRange(rowIndex, 7).setValue(adminMemo);
}

function updateSurveySheetResponse_(survey, response) {
  var sheet = ensureSurveySheet_(survey);
  var rowIndex = findRowIndexByColumn_(sheet, 2, response.id);
  if (!rowIndex) return;

  var answerMap = {};
  response.answers.forEach(function (answer) {
    answerMap[answer.questionId] = answer.value || "";
  });

  var row = [
    response.submittedAt,
    response.id,
    response.customerClientId,
    response.customerName,
    response.customerEmail,
    response.status,
    response.adminMemo,
  ];
  survey.questions.forEach(function (question) {
    row.push(answerMap[question.id] || "");
  });
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
}

function deleteRowByResponseId_(sheet, responseId, column) {
  if (!sheet) return;
  var rowIndex = findRowIndexByColumn_(sheet, column, responseId);
  if (rowIndex) sheet.deleteRow(rowIndex);
}

function findMasterRowIndex_(responseId) {
  return findRowIndexByColumn_(getSpreadsheet_().getSheetByName(MASTER_SHEET_NAME), 2, responseId);
}

function findRowIndexByColumn_(sheet, column, value) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var values = sheet.getRange(2, column, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < values.length; i += 1) {
    if (String(values[i][0]) === String(value)) return i + 2;
  }
  return 0;
}

function ensureSpreadsheet_() {
  var spreadsheet = getSpreadsheet_();
  ensureSheet_(spreadsheet, MASTER_SHEET_NAME, MASTER_HEADERS);
  ensureMeasurementsSheet_();
  ensureBijirisPostsSheet_();
  ensureBijirisPostAttachmentsSheet_();
  getSurveys_().forEach(ensureSurveySheet_);
}

function ensureSurveySheet_(survey) {
  var headers = ["送信日時", "回答ID", "端末ID", "お名前", "メールアドレス", "対応状況", "管理メモ"];
  survey.questions.forEach(function (question) {
    headers.push(question.label);
  });
  return ensureSheet_(getSpreadsheet_(), survey.title, headers);
}

function ensureSheet_(spreadsheet, name, headers) {
  var sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function getAdminInfo_() {
  var spreadsheet = getSpreadsheet_();
  var rootFolder = getRootPhotoFolder_();
  var credentials = getAdminCredentials_();
  var backupFolder = getBackupFolder_();
  var bijirisPostsFolder = getBijirisPostsRootFolder_();
  return {
    backend: "gas",
    adminUsername: credentials.username,
    adminUsers: getAdminUsers_().map(publicAdminUser_),
    ownerEmail: getOwnerEmail_(),
    pushAppId: getPushAppId_(),
    pushConfigured: isPushNotificationConfigured_(),
    customerAppUrl: getCustomerAppUrl_(),
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    masterSheetName: MASTER_SHEET_NAME,
    photoRootFolderName: ROOT_DRIVE_FOLDER_NAME,
    photoRootFolderUrl: rootFolder.getUrl(),
    bijirisPostsFolderUrl: bijirisPostsFolder.getUrl(),
    backupFolderUrl: backupFolder.getUrl(),
    customerProfiles: getAdminCustomerProfiles_(),
    latestBackup: getBackupMeta_(),
    lastMaintenance: getLastMaintenanceMeta_(),
    version: VERSION,
  };
}

function getSpreadsheet_() {
  var properties = PropertiesService.getScriptProperties();
  var candidateIds = uniqueValues_([
    properties.getProperty("ACTIVE_SPREADSHEET_ID"),
    properties.getProperty("SPREADSHEET_ID"),
    SPREADSHEET_ID,
  ]);

  for (var i = 0; i < candidateIds.length; i += 1) {
    try {
      var spreadsheet = SpreadsheetApp.openById(candidateIds[i]);
      rememberSpreadsheetId_(spreadsheet.getId());
      return spreadsheet;
    } catch (error) {
      // Try the next configured spreadsheet id.
    }
  }

  var created = SpreadsheetApp.create(DEFAULT_SPREADSHEET_TITLE);
  rememberSpreadsheetId_(created.getId());
  return created;
}

function rememberSpreadsheetId_(spreadsheetId) {
  PropertiesService.getScriptProperties().setProperty("ACTIVE_SPREADSHEET_ID", String(spreadsheetId));
}

function uniqueValues_(values) {
  var seen = {};
  return values.filter(function (value) {
    var normalized = normalizeText_(value);
    if (!normalized || seen[normalized]) return false;
    seen[normalized] = true;
    return true;
  });
}

function adminLogin_(loginId, password) {
  var normalizedLoginId = normalizeText_(loginId);
  ensureLoginAllowed_(normalizedLoginId);
  var user = findAdminUserByUsername_(normalizedLoginId);
  if (!user || String(password || "") !== user.password) {
    recordLoginFailure_(normalizedLoginId);
    appendErrorLog_("adminLogin", "ログイン失敗", {
      loginId: normalizedLoginId,
    });
    throw new Error("ログインIDまたはパスワードが違います。");
  }
  clearLoginFailures_(normalizedLoginId);
  var expiresAt = Date.now() + 8 * 60 * 60 * 1000;
  appendAuditLog_("admin.login", {
    loginId: normalizedLoginId,
  });
  return {
    token: makeToken_(user.username, expiresAt),
    expiresAt: new Date(expiresAt).toISOString(),
  };
}

function getAdminCredentials_() {
  var users = getAdminUsers_();
  var primary = users[0];
  return {
    username: primary.username,
    password: primary.password,
  };
}

function updateAdminCredentials_(loginId, password) {
  var properties = PropertiesService.getScriptProperties();
  var users = getAdminUsers_();
  var current = users[0];
  var nextLoginId = normalizeText_(loginId);
  var nextPassword = String(password || "");

  if (!nextLoginId) throw new Error("ログインIDを入力してください。");
  if (!nextPassword) nextPassword = current.password;
  if (String(nextPassword).length < 4) {
    throw new Error("パスワードは4文字以上で入力してください。");
  }

  users[0] = Object.assign({}, current, {
    username: nextLoginId,
    password: nextPassword,
  });
  properties.setProperty(ADMIN_USERS_PROPERTY_KEY, JSON.stringify(users));
  properties.setProperty("ADMIN_USERNAME", nextLoginId);
  properties.setProperty("ADMIN_PASSWORD", nextPassword);
  clearLoginFailures_(current.username);
  clearLoginFailures_(nextLoginId);
  appendAuditLog_("admin.credentials.update", {
    loginId: nextLoginId,
  });

  return {
    ok: true,
    adminInfo: getAdminInfo_(),
  };
}

function requireAdmin_(token) {
  if (!verifyToken_(token)) throw new Error("管理者ログインが必要です。");
}

function makeToken_(username, expiresAt) {
  var base = username + "|" + expiresAt;
  return Utilities.base64EncodeWebSafe(JSON.stringify({
    u: username,
    exp: expiresAt,
    sig: sign_(base),
  }));
}

function verifyToken_(token) {
  try {
    var raw = Utilities.newBlob(Utilities.base64DecodeWebSafe(String(token || ""))).getDataAsString();
    var payload = JSON.parse(raw);
    if (!payload.u || !payload.exp || Number(payload.exp) < Date.now()) return false;
    return payload.sig === sign_(payload.u + "|" + payload.exp);
  } catch (error) {
    return false;
  }
}

function sign_(value) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(String(value), getConfig_("TOKEN_SECRET", DEFAULT_TOKEN_SECRET))
  );
}

function getOwnerEmail_() {
  try {
    return Session.getEffectiveUser().getEmail();
  } catch (error) {
    return "";
  }
}

function getConfig_(key, fallback) {
  return PropertiesService.getScriptProperties().getProperty(key) || fallback;
}

function getPushAppId_() {
  return normalizeText_(getConfig_("ONESIGNAL_APP_ID", DEFAULT_ONESIGNAL_APP_ID));
}

function getPushRestApiKey_() {
  return normalizeText_(getConfig_("ONESIGNAL_REST_API_KEY", ""));
}

function isPushNotificationConfigured_() {
  return Boolean(getPushAppId_() && getPushRestApiKey_());
}

function getCustomerAppUrl_() {
  return normalizeText_(getConfig_("CUSTOMER_APP_URL", DEFAULT_CUSTOMER_APP_URL)) || DEFAULT_CUSTOMER_APP_URL;
}

function configurePushNotifications(pushAppId, restApiKey, customerAppUrl) {
  return updatePushConfig_(
    {
      pushAppId: pushAppId,
      restApiKey: restApiKey,
      customerAppUrl: customerAppUrl,
    },
    { source: "clasp" }
  );
}

function buildBijirisNotificationUrl_(post) {
  var base = getCustomerAppUrl_();
  var separator = base.indexOf("?") >= 0 ? "&" : "?";
  return base + separator + "page=bijiris&postId=" + encodeURIComponent(normalizeText_(post && post.id));
}

function buildBijirisNotificationPayload_(post, payload, mode) {
  var title = normalizeText_(payload && payload.notificationTitle) || (mode === "update" ? "豆知識を更新しました" : "新しい豆知識を追加しました");
  var body = normalizeText_(payload && payload.notificationBody) || normalizeText_(post && post.summary) || normalizeText_(post && post.title) || "新しい豆知識があります。";
  return {
    app_id: getPushAppId_(),
    filters: [
      { field: "tag", key: ONE_SIGNAL_APP_SCOPE_KEY, relation: "=", value: ONE_SIGNAL_APP_SCOPE_VALUE },
    ],
    headings: { en: title, ja: title },
    contents: { en: body, ja: body },
    url: buildBijirisNotificationUrl_(post),
    data: {
      page: "bijiris",
      postId: normalizeText_(post && post.id),
    },
    ios_badgeType: "Increase",
    ios_badgeCount: 1,
  };
}

function sendBijirisPushNotification_(post, payload, mode) {
  if (!isPushNotificationConfigured_()) return false;
  var response = UrlFetchApp.fetch("https://onesignal.com/api/v1/notifications", {
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    headers: {
      Authorization: "Basic " + getPushRestApiKey_(),
    },
    payload: JSON.stringify(buildBijirisNotificationPayload_(post, payload, mode)),
  });
  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error("OneSignal送信失敗: " + response.getContentText());
  }
  return true;
}

function notifyBijirisPostIfNeeded_(post, payload, mode) {
  if (!(post && normalizeText_(post.status) === "published" && payload && payload.notifyCustomers === true)) {
    return;
  }
  try {
    var sent = sendBijirisPushNotification_(post, payload, mode);
    appendAuditLog_("bijiris_post.notify", {
      postId: post.id,
      title: post.title,
      mode: mode,
      sent: sent === true,
    });
  } catch (error) {
    appendAuditLog_("bijiris_post.notify_error", {
      postId: post && post.id,
      title: post && post.title,
      mode: mode,
      error: error && error.message ? error.message : String(error),
    });
  }
}

function normalizeText_(value) {
  return String(value || "").trim();
}

function normalizeKana_(value) {
  return String(value || "").replace(/\s+/g, "").trim();
}

function parseTicketLabelNumber_(value) {
  var matched = normalizeText_(value).match(/\d+/);
  return matched ? Number(matched[0]) : 0;
}

function parseTicketPlanCount_(plan) {
  var text = normalizeText_(plan);
  if (text.indexOf("10") >= 0) return 10;
  if (text.indexOf("6") >= 0) return 6;
  return parseTicketLabelNumber_(text);
}

// 完了枚数: 種類(6回券/10回券)を問わず、使い切ったカードの通算枚数。
// 1枚の完了 = あるカード(種類＋何枚目)について最終回(種類の上限)の回答が送信されていること。
// 管理者の手動カード設定(activeTicketCard override)は含めず、履歴(回答)ベースで算出する。
function getCompletedTicketCardCountsByCustomer_(responses) {
  var list = Array.isArray(responses) ? responses : getResponses_({});
  var maxRoundByCustomerCard = {};
  list.forEach(function (response) {
    var name = normalizeText_(response && response.customerName);
    if (!name) return;
    var answers = response && response.answers;
    var plan = getAnswerValueByQuestionIds_(answers, CUSTOMER_TICKET_INFO_QUESTION_IDS.plan);
    var sheet = parseTicketLabelNumber_(getAnswerValueByQuestionIds_(answers, CUSTOMER_TICKET_INFO_QUESTION_IDS.sheet));
    var round = parseTicketLabelNumber_(getAnswerValueByQuestionIds_(answers, CUSTOMER_TICKET_INFO_QUESTION_IDS.round));
    if (!plan || sheet <= 0 || round <= 0) return;
    if (!maxRoundByCustomerCard[name]) maxRoundByCustomerCard[name] = {};
    var cardKey = plan + "|" + sheet;
    var prev = maxRoundByCustomerCard[name][cardKey];
    if (!prev || round > prev.round) {
      maxRoundByCustomerCard[name][cardKey] = { plan: plan, round: round };
    }
  });
  var counts = {};
  Object.keys(maxRoundByCustomerCard).forEach(function (name) {
    var cards = maxRoundByCustomerCard[name];
    var completed = 0;
    Object.keys(cards).forEach(function (cardKey) {
      var card = cards[cardKey];
      var planCount = parseTicketPlanCount_(card.plan);
      if (planCount > 0 && card.round >= planCount) completed += 1;
    });
    counts[name] = completed;
  });
  return counts;
}

function getCompletedTicketCardCountForCustomer_(customerName) {
  var name = normalizeText_(customerName);
  if (!name) return 0;
  var counts = getCompletedTicketCardCountsByCustomer_(getResponses_({ customerName: name }));
  return counts[name] || 0;
}

function updateCustomerRewardRedemption_(customerName, threshold, handed) {
  var name = normalizeText_(customerName);
  var thresholdNum = Math.floor(Number(threshold));
  if (!name) throw new Error("お客様を指定してください。");
  if (!isFinite(thresholdNum) || thresholdNum <= 0) throw new Error("しきい値が正しくありません。");

  var profiles = getCustomerProfiles_();
  var match = findCustomerProfileByName_(profiles, name, "");
  var record = match ? normalizeCustomerProfileRecord_(match.profile, match.key) : null;
  if (!record) throw new Error("お客様が見つかりません。");

  var redemptions = record.rewardRedemptions
    ? JSON.parse(JSON.stringify(record.rewardRedemptions))
    : {};
  if (handed === true) {
    redemptions[String(thresholdNum)] = { handed: true, handedAt: new Date().toISOString() };
  } else {
    delete redemptions[String(thresholdNum)];
  }
  record.rewardRedemptions = normalizeRewardRedemptions_(redemptions);
  record.updatedAt = new Date().toISOString();

  var saved = saveCustomerProfileRecord_(profiles, match.key, record, {
    replaceRewardRedemptions: true,
  });
  saveCustomerProfiles_(profiles);

  appendAuditLog_("customer.reward_redemption.update", {
    customerName: saved.name,
    threshold: thresholdNum,
    handed: handed === true,
  });

  var publicProfile = publicCustomerProfile_(saved);
  if (publicProfile) {
    publicProfile.completedTicketCardCount = getCompletedTicketCardCountForCustomer_(saved.name);
  }
  return { customerProfile: publicProfile };
}

function normalizeActiveTicketCard_(value) {
  if (!value || typeof value !== "object") return null;
  var plan = normalizeText_(value.plan || value.ticketPlan || value.size);
  var sheetNumber = Math.floor(
    Number(value.sheetNumber || value.sheetNumberValue) ||
    parseTicketLabelNumber_(value.sheetLabel || value.ticketSheet || value.sheet)
  );
  var round = Math.max(
    0,
    Math.floor(
      Number(value.round || value.roundValue) ||
      parseTicketLabelNumber_(value.roundLabel || value.ticketRound || value.round)
    )
  );
  if (!plan || sheetNumber <= 0) return null;
  return {
    plan: plan,
    sheetNumber: sheetNumber,
    round: round,
  };
}

function publicActiveTicketCard_(value) {
  var normalized = normalizeActiveTicketCard_(value);
  if (!normalized) return null;
  return {
    plan: normalized.plan,
    sheetNumber: normalized.sheetNumber,
    sheetLabel: normalized.sheetNumber + "枚目",
    round: normalized.round,
    roundLabel: normalized.round + "回目",
  };
}

function formatMemberNumber_(index) {
  return "M" + Utilities.formatString("%04d", Math.max(1, Math.floor(Number(index) || 0)));
}

function parseMemberNumberIndex_(value) {
  var matched = normalizeText_(value).toUpperCase().match(/(\d+)/);
  return matched ? Math.max(0, Number(matched[1]) || 0) : 0;
}

function normalizeMemberNumber_(value) {
  var index = parseMemberNumberIndex_(value);
  return index > 0 ? formatMemberNumber_(index) : "";
}

function getStoredNextMemberNumberIndex_() {
  var value = Number(PropertiesService.getScriptProperties().getProperty(NEXT_MEMBER_NUMBER_PROPERTY_KEY));
  return Number.isFinite(value) && value > 0 ? Math.floor(value) : 1;
}

function saveNextMemberNumberIndex_(value) {
  PropertiesService.getScriptProperties().setProperty(
    NEXT_MEMBER_NUMBER_PROPERTY_KEY,
    String(Math.max(1, Math.floor(Number(value) || 1)))
  );
}

function getHighestMemberNumberIndex_(profiles) {
  var highest = 0;
  Object.keys(profiles || {}).forEach(function (key) {
    var profile = normalizeCustomerProfileRecord_(profiles[key], key);
    if (!profile) return;
    highest = Math.max(highest, parseMemberNumberIndex_(profile.memberNumber));
  });
  return highest;
}

function ensureCustomerProfileMemberNumbers_(profiles) {
  var nextIndex = Math.max(getStoredNextMemberNumberIndex_(), getHighestMemberNumberIndex_(profiles) + 1, 1);
  var changed = false;
  Object.keys(profiles || {})
    .sort(function (left, right) {
      return normalizeText_(left).localeCompare(normalizeText_(right));
    })
    .forEach(function (key) {
      var profile = normalizeCustomerProfileRecord_(profiles[key], key);
      if (!profile) return;
      if (!profile.memberNumber) {
        profile.memberNumber = formatMemberNumber_(nextIndex);
        nextIndex += 1;
        profiles[key] = profile;
        changed = true;
      }
    });
  if (nextIndex !== getStoredNextMemberNumberIndex_()) {
    saveNextMemberNumberIndex_(nextIndex);
  }
  return changed;
}

function getBackupMeta_() {
  var meta = parseJson_(PropertiesService.getScriptProperties().getProperty(BACKUP_META_PROPERTY_KEY), {});
  return {
    at: normalizeText_(meta && meta.at),
    fileId: normalizeText_(meta && meta.fileId),
    fileName: normalizeText_(meta && meta.fileName),
    fileUrl: normalizeText_(meta && meta.fileUrl),
  };
}

function saveBackupMeta_(meta) {
  PropertiesService.getScriptProperties().setProperty(
    BACKUP_META_PROPERTY_KEY,
    JSON.stringify({
      at: normalizeText_(meta && meta.at),
      fileId: normalizeText_(meta && meta.fileId),
      fileName: normalizeText_(meta && meta.fileName),
      fileUrl: normalizeText_(meta && meta.fileUrl),
    })
  );
}

function getLastMaintenanceMeta_() {
  var meta = parseJson_(PropertiesService.getScriptProperties().getProperty(LAST_MAINTENANCE_META_PROPERTY_KEY), {});
  return {
    at: normalizeText_(meta && meta.at),
    purged: Number(meta && meta.purged || 0),
    autoBackupEnabled: meta && meta.autoBackupEnabled === true,
    backupInfo: meta && meta.backupInfo ? {
      at: normalizeText_(meta.backupInfo.at),
      fileId: normalizeText_(meta.backupInfo.fileId),
      fileName: normalizeText_(meta.backupInfo.fileName),
      fileUrl: normalizeText_(meta.backupInfo.fileUrl),
    } : null,
  };
}

function saveLastMaintenanceMeta_(meta) {
  PropertiesService.getScriptProperties().setProperty(
    LAST_MAINTENANCE_META_PROPERTY_KEY,
    JSON.stringify({
      at: normalizeText_(meta && meta.at),
      purged: Number(meta && meta.purged || 0),
      autoBackupEnabled: meta && meta.autoBackupEnabled === true,
      backupInfo: meta && meta.backupInfo ? {
        at: normalizeText_(meta.backupInfo.at),
        fileId: normalizeText_(meta.backupInfo.fileId),
        fileName: normalizeText_(meta.backupInfo.fileName),
        fileUrl: normalizeText_(meta.backupInfo.fileUrl),
      } : null,
    })
  );
}

function buildActiveTicketCardFromAnswers_(answers) {
  var ticketPlan = getAnswerValueByQuestionIds_(answers, CUSTOMER_TICKET_INFO_QUESTION_IDS.plan);
  var ticketSheet = getAnswerValueByQuestionIds_(answers, CUSTOMER_TICKET_INFO_QUESTION_IDS.sheet);
  var ticketRound = getAnswerValueByQuestionIds_(answers, CUSTOMER_TICKET_INFO_QUESTION_IDS.round);
  return normalizeActiveTicketCard_({
    plan: ticketPlan,
    sheetLabel: ticketSheet,
    roundLabel: ticketRound,
  });
}

function normalizeEmail_(value) {
  return normalizeText_(value).toLowerCase();
}

function normalizeStatus_(status) {
  return ["new", "checked", "done", "trash"].indexOf(String(status)) >= 0 ? String(status) : "new";
}

function normalizeSurveyStatus_(status) {
  var value = normalizeText_(status);
  return ["published", "draft", "archived"].indexOf(value) >= 0 ? value : "published";
}

function buildRawAnswerMap_(answers) {
  var map = {};
  (Array.isArray(answers) ? answers : []).forEach(function (answer) {
    var questionId = normalizeText_(answer && answer.questionId);
    if (!questionId) return;
    if (Array.isArray(answer.value)) {
      map[questionId] = answer.value.map(normalizeText_).filter(Boolean);
      return;
    }
    if (Array.isArray(answer.files)) {
      map[questionId] = answer.files;
      return;
    }
    var value = normalizeText_(answer && answer.value);
    map[questionId] = value ? [value] : [];
  });
  return map;
}

function isLegacyTicketEndLastPhotoQuestion_(question, survey) {
  return survey && survey.id === "survey_bijiris_ticket_end" && question && question.id === "q_ticket_end_photo_last";
}

function isLegacyTicketEndLastPhotoVisible_(rawAnswerMap) {
  var ticketSizeValues = rawAnswerMap && rawAnswerMap.q_ticket_end_ticket_size;
  var ticketRoundValues = rawAnswerMap && rawAnswerMap.q_ticket_end_ticket_round;
  var ticketSize = Array.isArray(ticketSizeValues) && ticketSizeValues.length ? normalizeText_(ticketSizeValues[0]) : "";
  var ticketRound = Array.isArray(ticketRoundValues) && ticketRoundValues.length ? normalizeText_(ticketRoundValues[0]) : "";
  return (
    (ticketSize === "6回券" && ticketRound === "6回目") ||
    (ticketSize === "10回券" && ticketRound === "10回目")
  );
}

function getBijirisSessionPhotoConfig_(question, survey) {
  if (!(survey && survey.id === "survey_bijiris_session" && question)) return null;
  if (question.id === "q_bijiris_session_ticket_photos") {
    return { maxFiles: 4, requiredCount: 4 };
  }
  if (
    [
      "q_bijiris_session_monitor_photos_6",
      "q_bijiris_session_monitor_photos_10",
      "q_bijiris_session_ticket_end_photos_6",
      "q_bijiris_session_ticket_end_photos_10",
      "q_bijiris_session_monitor_photos",
      "q_bijiris_session_ticket_end_photos"
    ].indexOf(question.id) >= 0
  ) {
    return { maxFiles: 2, requiredCount: 2 };
  }
  return null;
}

function isLegacyBijirisSessionPhotoQuestion_(question, survey) {
  return Boolean(
    survey && survey.id === "survey_bijiris_session" &&
      question &&
      ["q_bijiris_session_monitor_photos", "q_bijiris_session_ticket_end_photos"].indexOf(question.id) >= 0
  );
}

function isBijirisSessionFinalPhotoVisible_(rawAnswerMap) {
  var sessionTypeValues = rawAnswerMap && rawAnswerMap.q_bijiris_session_type;
  var ticketPlanValues = rawAnswerMap && rawAnswerMap.q_bijiris_session_ticket_plan;
  var ticketRoundValues = rawAnswerMap && rawAnswerMap.q_bijiris_session_ticket_round;
  var sessionType = Array.isArray(sessionTypeValues) && sessionTypeValues.length ? normalizeText_(sessionTypeValues[0]) : "";
  var ticketPlan = Array.isArray(ticketPlanValues) && ticketPlanValues.length ? normalizeText_(ticketPlanValues[0]) : "";
  var ticketRound = Array.isArray(ticketRoundValues) && ticketRoundValues.length ? normalizeText_(ticketRoundValues[0]) : "";
  return (
    sessionType === "回数券" &&
    ((ticketPlan === "6回券" && ticketRound === "6回目") ||
      (ticketPlan === "10回券" && ticketRound === "10回目"))
  );
}

function getPhotoQuestionMaxFiles_(question, survey) {
  var bijirisSessionPhotoConfig = getBijirisSessionPhotoConfig_(question, survey);
  if (bijirisSessionPhotoConfig) return bijirisSessionPhotoConfig.maxFiles;
  return 6;
}

function getPhotoQuestionRequiredCount_(question, visible, survey) {
  if (!visible) return 0;
  var bijirisSessionPhotoConfig = getBijirisSessionPhotoConfig_(question, survey);
  if (bijirisSessionPhotoConfig) return bijirisSessionPhotoConfig.requiredCount;
  if (isLegacyTicketEndLastPhotoQuestion_(question, survey) && !getQuestionVisibilityConditions_(question).length) {
    return 1;
  }
  return question && question.required === false ? 0 : 1;
}

function isQuestionVisible_(question, rawAnswerMap, survey) {
  if (isLegacyBijirisSessionPhotoQuestion_(question, survey)) {
    return isBijirisSessionFinalPhotoVisible_(rawAnswerMap || {});
  }
  var conditions = getQuestionVisibilityConditions_(question);
  if (conditions.length) {
    return conditions.every(function (condition) {
      var values = rawAnswerMap && rawAnswerMap[condition.questionId];
      if (!Array.isArray(values) || !values.length) return false;
      var expected = normalizeText_(condition.value);
      return values.some(function (value) {
        return normalizeText_(value) === expected;
      });
    });
  }
  if (isLegacyTicketEndLastPhotoQuestion_(question, survey)) {
    return isLegacyTicketEndLastPhotoVisible_(rawAnswerMap || {});
  }
  if (!question) return true;
  if (!question.visibleWhen) return true;
  var map = rawAnswerMap || {};
  var values = map[question.visibleWhen.questionId];
  if (!Array.isArray(values) || !values.length) return false;
  var expected = normalizeText_(question.visibleWhen.value);
  return values.some(function (value) {
    return normalizeText_(value) === expected;
  });
}

function isQuestionRequired_(question, visible, survey) {
  return getPhotoQuestionRequiredCount_(question, visible, survey) > 0 || (
    visible && !(question && question.type === "photo") && !(question && question.required === false)
  );
}

function notifyNewResponse_(response) {
  var preferences = getPreferences_();
  if (!preferences.notificationEnabled || !preferences.notificationEmail) return;
  var templateData = {
    customerName: normalizeText_(response && response.customerName),
    surveyTitle: normalizeText_(response && response.surveyTitle),
    submittedAt: normalizeText_(response && response.submittedAt),
    responseId: normalizeText_(response && response.id),
  };
  var subject = renderTemplate_(preferences.notificationSubject, templateData);
  var body = renderTemplate_(preferences.notificationBody, templateData);

  try {
    MailApp.sendEmail(preferences.notificationEmail, subject, body);
  } catch (error) {
    appendErrorLog_("notifyNewResponse", error.message || "通知メール送信エラー", {
      responseId: response && response.id,
      notificationEmail: preferences.notificationEmail,
    });
  }
}

function getLoginAttempts_() {
  return parseJson_(PropertiesService.getScriptProperties().getProperty(LOGIN_ATTEMPTS_PROPERTY_KEY), {});
}

function saveLoginAttempts_(attempts) {
  PropertiesService.getScriptProperties().setProperty(
    LOGIN_ATTEMPTS_PROPERTY_KEY,
    JSON.stringify(attempts || {})
  );
}

function pruneLoginAttempts_(attempts) {
  var nowTime = Date.now();
  Object.keys(attempts || {}).forEach(function (key) {
    var item = attempts[key];
    if (!item || nowTime - Number(item.lastAt || 0) > LOGIN_LOCK_WINDOW_MS) {
      delete attempts[key];
    }
  });
  return attempts;
}

function ensureLoginAllowed_(loginId) {
  var key = normalizeText_(loginId) || "_default";
  var attempts = pruneLoginAttempts_(getLoginAttempts_());
  var item = attempts[key];
  if (item && Number(item.count || 0) >= LOGIN_MAX_ATTEMPTS) {
    throw new Error("ログイン失敗が続いたため、一時的にログインを停止しています。15分後に再試行してください。");
  }
  saveLoginAttempts_(attempts);
}

function recordLoginFailure_(loginId) {
  var key = normalizeText_(loginId) || "_default";
  var attempts = pruneLoginAttempts_(getLoginAttempts_());
  var nowTime = Date.now();
  var current = attempts[key] || { count: 0, firstAt: nowTime, lastAt: nowTime };
  current.count = Number(current.count || 0) + 1;
  current.lastAt = nowTime;
  attempts[key] = current;
  saveLoginAttempts_(attempts);
}

function clearLoginFailures_(loginId) {
  var key = normalizeText_(loginId) || "_default";
  var attempts = pruneLoginAttempts_(getLoginAttempts_());
  delete attempts[key];
  saveLoginAttempts_(attempts);
}

function parseJson_(value, fallback) {
  try {
    return value ? JSON.parse(String(value)) : fallback;
  } catch (error) {
    return fallback;
  }
}

function stringifyDate_(value) {
  if (!value) return "";
  return value instanceof Date ? value.toISOString() : String(value);
}
