const crypto = require("node:crypto");
const fs = require("node:fs/promises");
const http = require("node:http");
const os = require("node:os");
const path = require("node:path");
const { URL } = require("node:url");
const { makeDefaultSurveys } = require("./default-surveys");

const ROOT_DIR = __dirname;
const STATIC_APPS = {
  "/customer": path.join(ROOT_DIR, "customer-app"),
  "/admin": path.join(ROOT_DIR, "admin-app"),
};
const DATA_DIR = process.env.DATA_DIR || path.join(ROOT_DIR, "data");
const DB_FILE = process.env.DB_FILE || path.join(DATA_DIR, "db.json");
const PORT = Number(process.env.PORT || 3000);
const HOST = process.env.HOST || "0.0.0.0";
const ADMIN_USERNAME = process.env.ADMIN_USERNAME || "mayumi2026";
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "3939";
const SESSION_SECRET =
  process.env.SESSION_SECRET || "change-this-secret-before-production";
const TOKEN_TTL_SECONDS = 60 * 60 * 8;
const BODY_LIMIT_BYTES = 12 * 1024 * 1024;
const PHOTO_FILE_LIMIT = 6;
const PHOTO_DATA_URL_LIMIT_BYTES = 2.5 * 1024 * 1024;
const QUESTION_TYPES = ["text", "textarea", "rating", "choice", "checkbox", "photo"];

let writeQueue = Promise.resolve();

function sendJson(res, status, data) {
  const body = JSON.stringify(data);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    "Access-Control-Allow-Origin": process.env.CORS_ORIGIN || "*",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
  });
  res.end(body);
}

function sendError(res, status, message) {
  sendJson(res, status, { error: message });
}

function id(prefix) {
  return `${prefix}_${crypto.randomUUID()}`;
}

function now() {
  return new Date().toISOString();
}

function getLanUrls() {
  return Object.values(os.networkInterfaces())
    .flat()
    .filter((network) => network && network.family === "IPv4" && !network.internal)
    .map((network) => `http://${network.address}:${PORT}`);
}

function base64Url(input) {
  return Buffer.from(input).toString("base64url");
}

function signToken(payload) {
  const encodedPayload = base64Url(JSON.stringify(payload));
  const signature = crypto
    .createHmac("sha256", SESSION_SECRET)
    .update(encodedPayload)
    .digest("base64url");
  return `${encodedPayload}.${signature}`;
}

function verifyToken(token) {
  if (!token || !token.includes(".")) return false;
  const [encodedPayload, signature] = token.split(".");
  const expected = crypto
    .createHmac("sha256", SESSION_SECRET)
    .update(encodedPayload)
    .digest("base64url");
  const signatureBuffer = Buffer.from(signature);
  const expectedBuffer = Buffer.from(expected);
  if (
    signatureBuffer.length !== expectedBuffer.length ||
    !crypto.timingSafeEqual(signatureBuffer, expectedBuffer)
  ) {
    return false;
  }

  try {
    const payload = JSON.parse(Buffer.from(encodedPayload, "base64url").toString("utf8"));
    return payload.sub === "admin" && payload.exp > Math.floor(Date.now() / 1000);
  } catch {
    return false;
  }
}

function requireAdmin(req, res) {
  const auth = req.headers.authorization || "";
  const token = auth.startsWith("Bearer ") ? auth.slice(7) : "";
  if (!verifyToken(token)) {
    sendError(res, 401, "管理者ログインが必要です。");
    return false;
  }
  return true;
}

function seedDb() {
  return {
    surveys: makeDefaultSurveys(now()),
    responses: [],
    measurements: [],
    bijirisPosts: [],
    settings: {
      adminUsername: ADMIN_USERNAME,
      adminPassword: ADMIN_PASSWORD,
    },
  };
}

function getAdminSettings(db) {
  if (!db.settings || typeof db.settings !== "object") {
    db.settings = {};
  }

  const storedUsername = normalizeText(db.settings.adminUsername);
  const storedPassword = String(db.settings.adminPassword || "");
  db.settings.adminUsername =
    !storedUsername || storedUsername === "admin" ? ADMIN_USERNAME : storedUsername;
  db.settings.adminPassword =
    !storedPassword || storedPassword === "admin123" ? ADMIN_PASSWORD : storedPassword;

  return {
    adminUsername: db.settings.adminUsername,
    adminPassword: db.settings.adminPassword,
  };
}

function buildAdminInfo(db) {
  const settings = getAdminSettings(db);
  return {
    backend: "rest",
    adminUsername: settings.adminUsername,
    ownerEmail: "",
    spreadsheetId: "",
    spreadsheetUrl: "",
    masterSheetName: "data/db.json",
    photoRootFolderName: "ローカル確認",
    photoRootFolderUrl: "",
    bijirisPostsFolderUrl: "",
  };
}

function surveyComparable(survey) {
  const { createdAt, updatedAt, ...rest } = survey;
  return rest;
}

function hasCurrentDefaultSurveys(db) {
  const defaults = makeDefaultSurveys("default");
  return (
    Array.isArray(db.surveys) &&
    JSON.stringify(db.surveys.map(surveyComparable)) ===
      JSON.stringify(defaults.map(surveyComparable))
  );
}

async function syncDefaultSurveys() {
  const raw = await fs.readFile(DB_FILE, "utf8");
  const db = JSON.parse(raw);
  if (hasCurrentDefaultSurveys(db)) return;

  db.surveys = makeDefaultSurveys(now());
  db.responses = Array.isArray(db.responses) ? db.responses : [];
  db.bijirisPosts = Array.isArray(db.bijirisPosts) ? db.bijirisPosts : [];
  await fs.writeFile(DB_FILE, `${JSON.stringify(db, null, 2)}\n`);
}

async function ensureDb() {
  await fs.mkdir(DATA_DIR, { recursive: true });
  try {
    await fs.access(DB_FILE);
  } catch {
    await fs.writeFile(DB_FILE, `${JSON.stringify(seedDb(), null, 2)}\n`);
    return;
  }
  await syncDefaultSurveys();
}

async function readDb() {
  await ensureDb();
  const raw = await fs.readFile(DB_FILE, "utf8");
  const db = JSON.parse(raw);
  db.surveys = Array.isArray(db.surveys) ? db.surveys : [];
  db.responses = Array.isArray(db.responses) ? db.responses : [];
  db.measurements = Array.isArray(db.measurements) ? db.measurements : [];
  db.bijirisPosts = Array.isArray(db.bijirisPosts) ? db.bijirisPosts : [];
  return db;
}

async function writeDb(db) {
  await fs.writeFile(DB_FILE, `${JSON.stringify(db, null, 2)}\n`);
}

async function updateDb(mutator) {
  const run = async () => {
    const db = await readDb();
    const result = await mutator(db);
    await writeDb(db);
    return result;
  };
  writeQueue = writeQueue.then(run, run);
  return writeQueue;
}

async function readBody(req) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > BODY_LIMIT_BYTES) {
      throw new Error("送信データが大きすぎます。");
    }
    chunks.push(chunk);
  }
  if (!chunks.length) return {};
  return JSON.parse(Buffer.concat(chunks).toString("utf8"));
}

function normalizeText(value) {
  return String(value ?? "").trim();
}

function normalizeEmail(value) {
  return normalizeText(value).toLowerCase();
}

function normalizeResponseStatus(status) {
  return ["new", "checked", "done"].includes(status) ? status : "new";
}

function normalizeBijirisPostStatus(status) {
  return ["published", "draft", "archived"].includes(normalizeText(status))
    ? normalizeText(status)
    : "draft";
}

function normalizeBijirisPostFile(file, kind = "photo", index = 0) {
  const name = normalizeText(file?.name) || (kind === "pdf" ? `資料${index + 1}.pdf` : `写真${index + 1}.jpg`);
  if (kind === "pdf") {
    const base64Data = normalizeText(file?.base64Data);
    const thumbnailDataUrl = normalizeText(file?.thumbnailDataUrl);
    const dataUrl = /^data:application\/pdf;base64,/i.test(String(file?.url || ""))
      ? String(file.url)
      : base64Data
        ? `data:application/pdf;base64,${base64Data}`
        : "";
    const url = normalizeText(file?.url) || dataUrl;
    if (!url) return null;
    return {
      kind,
      name,
      type: "application/pdf",
      fileId: normalizeText(file?.fileId),
      title: normalizeText(file?.title) || name.replace(/\.pdf$/i, ""),
      url,
      previewUrl: normalizeText(file?.previewUrl) || url,
      downloadUrl: normalizeText(file?.downloadUrl) || url,
      thumbnailUrl: normalizeText(file?.thumbnailUrl) || thumbnailDataUrl,
    };
  }
  const dataUrl = String(file?.dataUrl || file?.url || "");
  if (!/^data:image\/(jpeg|jpg|png|webp);base64,/i.test(dataUrl)) return null;
  return {
    kind,
    name,
    type: normalizeText(file?.type) || "image/jpeg",
    fileId: normalizeText(file?.fileId),
    url: normalizeText(file?.url) || dataUrl,
    previewUrl: normalizeText(file?.previewUrl) || dataUrl,
    downloadUrl: normalizeText(file?.downloadUrl) || dataUrl,
    thumbnailUrl: normalizeText(file?.thumbnailUrl) || dataUrl,
  };
}

function validateBijirisPostPayload(payload, existing) {
  const title = normalizeText(payload?.title);
  if (!title) throw new Error("タイトルを入力してください。");
  const summary = normalizeText(payload?.summary);
  const body = normalizeText(payload?.body);
  const photos = (Array.isArray(payload?.photos) ? payload.photos : [])
    .map((file, index) => normalizeBijirisPostFile(file, "photo", index))
    .filter(Boolean);
  const documents = (Array.isArray(payload?.documents) ? payload.documents : [])
    .map((file, index) => normalizeBijirisPostFile(file, "pdf", index))
    .filter(Boolean);
  if (!summary && !body && !photos.length && !documents.length) {
    throw new Error("本文、要約、写真、PDF のいずれかを入力してください。");
  }
  return {
    id: existing?.id || normalizeText(payload?.id) || id("bijiris"),
    title,
    category: normalizeText(payload?.category) || "豆知識",
    summary,
    body,
    status: normalizeBijirisPostStatus(payload?.status),
    pinned: payload?.pinned === true,
    createdAt: existing?.createdAt || now(),
    updatedAt: now(),
    publishedAt:
      normalizeBijirisPostStatus(payload?.status) === "published"
        ? normalizeText(payload?.publishedAt) || existing?.publishedAt || now()
        : existing?.publishedAt || "",
    photos,
    documents,
  };
}

function isChoiceType(type) {
  return type === "choice" || type === "checkbox";
}

function normalizeVisibilityCondition(condition) {
  const questionId = normalizeText(condition?.questionId);
  const value = normalizeText(condition?.value);
  if (!questionId || !value) return null;
  return { questionId, value };
}

function getQuestionVisibilityConditions(question) {
  const conditions = Array.isArray(question?.visibilityConditions)
    ? question.visibilityConditions.map(normalizeVisibilityCondition).filter(Boolean)
    : [];
  if (conditions.length) return conditions;
  const fallback = normalizeVisibilityCondition(question?.visibleWhen);
  return fallback ? [fallback] : [];
}

function buildRawAnswerMap(answers) {
  const map = {};
  (Array.isArray(answers) ? answers : []).forEach((answer) => {
    const questionId = normalizeText(answer?.questionId);
    if (!questionId) return;
    if (Array.isArray(answer?.files)) {
      map[questionId] = answer.files
        .map((file) => normalizeText(file?.name || file?.url || file?.dataUrl))
        .filter(Boolean);
      return;
    }
    if (Array.isArray(answer?.value)) {
      map[questionId] = answer.value.map(normalizeText).filter(Boolean);
      return;
    }
    const value = normalizeText(answer?.value);
    map[questionId] = value ? [value] : [];
  });
  return map;
}

function isLegacyTicketEndLastPhotoQuestion(question, survey) {
  return survey?.id === "survey_bijiris_ticket_end" && question?.id === "q_ticket_end_photo_last";
}

function getBijirisSessionPhotoConfig(question, survey) {
  if (!(survey?.id === "survey_bijiris_session" && question)) return null;
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
      "q_bijiris_session_ticket_end_photos",
    ].includes(question.id)
  ) {
    return { maxFiles: 2, requiredCount: 2 };
  }
  return null;
}

function isLegacyBijirisSessionPhotoQuestion(question, survey) {
  return Boolean(
    survey?.id === "survey_bijiris_session" &&
      question &&
      ["q_bijiris_session_monitor_photos", "q_bijiris_session_ticket_end_photos"].includes(question.id),
  );
}

function isBijirisSessionFinalPhotoVisible(answerMap = {}) {
  const sessionType = normalizeText(answerMap?.q_bijiris_session_type?.[0]);
  const ticketPlan = normalizeText(answerMap?.q_bijiris_session_ticket_plan?.[0]);
  const ticketRound = normalizeText(answerMap?.q_bijiris_session_ticket_round?.[0]);
  return (
    sessionType === "回数券" &&
    ((ticketPlan === "6回券" && ticketRound === "6回目") ||
      (ticketPlan === "10回券" && ticketRound === "10回目"))
  );
}

function isLegacyTicketEndLastPhotoVisible(answerMap) {
  const ticketSize = normalizeText(answerMap?.q_ticket_end_ticket_size?.[0]);
  const ticketRound = normalizeText(answerMap?.q_ticket_end_ticket_round?.[0]);
  return (
    (ticketSize === "6回券" && ticketRound === "6回目") ||
    (ticketSize === "10回券" && ticketRound === "10回目")
  );
}

function getPhotoQuestionMaxFiles(question, survey) {
  const bijirisSessionPhotoConfig = getBijirisSessionPhotoConfig(question, survey);
  if (bijirisSessionPhotoConfig) return bijirisSessionPhotoConfig.maxFiles;
  return PHOTO_FILE_LIMIT;
}

function getPhotoQuestionRequiredCount(question, visible, survey) {
  if (!visible) return 0;
  const bijirisSessionPhotoConfig = getBijirisSessionPhotoConfig(question, survey);
  if (bijirisSessionPhotoConfig) return bijirisSessionPhotoConfig.requiredCount;
  if (isLegacyTicketEndLastPhotoQuestion(question, survey) && !getQuestionVisibilityConditions(question).length) {
    return 1;
  }
  return question?.required === false ? 0 : 1;
}

function isQuestionVisible(question, answerMap, survey) {
  if (isLegacyBijirisSessionPhotoQuestion(question, survey)) {
    return isBijirisSessionFinalPhotoVisible(answerMap);
  }
  const conditions = getQuestionVisibilityConditions(question);
  if (conditions.length) {
    return conditions.every((condition) => {
      const values = Array.isArray(answerMap?.[condition.questionId]) ? answerMap[condition.questionId] : [];
      return values.some((value) => normalizeText(value) === condition.value);
    });
  }
  if (isLegacyTicketEndLastPhotoQuestion(question, survey)) {
    return isLegacyTicketEndLastPhotoVisible(answerMap);
  }
  return true;
}

function isQuestionRequired(question, visible, survey) {
  if (question?.type === "photo") {
    return getPhotoQuestionRequiredCount(question, visible, survey) > 0;
  }
  if (!visible) return false;
  return question?.required === false ? false : true;
}

function validateSurveyPayload(payload, existing) {
  const title = normalizeText(payload.title);
  const description = normalizeText(payload.description);
  const status = payload.status === "draft" ? "draft" : "published";
  const questions = Array.isArray(payload.questions) ? payload.questions : [];

  if (!title) throw new Error("タイトルを入力してください。");
  if (!description) throw new Error("説明文を入力してください。");
  if (!questions.length) throw new Error("質問は1つ以上必要です。");

  const normalizedQuestions = questions.map((question) => {
    const type = QUESTION_TYPES.includes(question.type)
      ? question.type
      : "text";
    const label = normalizeText(question.label);
    const options = Array.isArray(question.options)
      ? question.options.map(normalizeText).filter(Boolean)
      : [];
    const visibilityConditions = getQuestionVisibilityConditions(question);

    if (!label) throw new Error("質問文を入力してください。");
    if (isChoiceType(type) && options.length < 2) {
      throw new Error("選択式の質問は選択肢を2つ以上入力してください。");
    }

    return {
      id: question.id || id("question"),
      label,
      type,
      required: question.required === false ? false : true,
      options: isChoiceType(type) ? options : [],
      visibilityConditions,
      visibleWhen: visibilityConditions[0] || null,
    };
  });

  return {
    id: existing?.id || id("survey"),
    title,
    description,
    status,
    questions: normalizedQuestions,
    createdAt: existing?.createdAt || now(),
    updatedAt: now(),
  };
}

function normalizePhotoFiles(files, question, survey) {
  const maxFiles = getPhotoQuestionMaxFiles(question, survey);
  if (!Array.isArray(files)) return [];
  if (files.length > maxFiles) {
    throw new Error(`写真は${maxFiles}枚まで添付できます。`);
  }
  return files
    .slice(0, maxFiles)
    .map((file) => {
      const name = normalizeText(file?.name).slice(0, 140) || "photo.jpg";
      const type = normalizeText(file?.type) || "image/jpeg";
      const dataUrl = String(file?.dataUrl || "");
      if (!/^data:image\/(jpeg|jpg|png|webp);base64,/i.test(dataUrl)) {
        throw new Error("写真データの形式が正しくありません。");
      }
      if (Buffer.byteLength(dataUrl, "utf8") > PHOTO_DATA_URL_LIMIT_BYTES) {
        throw new Error("写真データが大きすぎます。");
      }
      return { name, type, dataUrl };
    });
}

function validateResponsePayload(db, payload) {
  const survey = db.surveys.find(
    (item) => item.id === payload.surveyId && item.status === "published",
  );
  if (!survey) throw new Error("回答できるアンケートが見つかりません。");

  const customerName = normalizeText(payload.customer?.name);
  if (!customerName) throw new Error("お名前を入力してください。");

  const answerMap = new Map(
    (Array.isArray(payload.answers) ? payload.answers : []).map((answer) => [
      answer.questionId,
      answer,
    ]),
  );
  const rawAnswerMap = buildRawAnswerMap(payload.answers);

  const answers = survey.questions.map((question) => {
    const answer = answerMap.get(question.id) || {};
    const visible = isQuestionVisible(question, rawAnswerMap, survey);
    const required = isQuestionRequired(question, visible, survey);
    if (question.type === "photo") {
      const requiredPhotoCount = getPhotoQuestionRequiredCount(question, visible, survey);
      const files = visible ? normalizePhotoFiles(answer.files, question, survey) : [];
      if (required && files.length < requiredPhotoCount) {
        throw new Error(requiredPhotoCount === 1 ? "未回答の質問があります。" : `写真を${requiredPhotoCount}枚添付してください。`);
      }
      return {
        questionId: question.id,
        label: question.label,
        type: question.type,
        value: files.length ? `${files.length}枚の写真` : "",
        files,
      };
    }

    const rawValue = answer.value;
    const values = Array.isArray(rawValue)
      ? rawValue.map(normalizeText).filter(Boolean)
      : [normalizeText(rawValue)].filter(Boolean);
    const value = visible ? values.join(", ") : "";

    if (required && !value) throw new Error("未回答の質問があります。");
    if (!value) {
      return {
        questionId: question.id,
        label: question.label,
        type: question.type,
        value: "",
      };
    }

    if (question.type === "rating" && !["1", "2", "3", "4", "5"].includes(value)) {
      throw new Error("評価は1から5で回答してください。");
    }
    if (question.type === "choice" && !question.options.includes(value)) {
      throw new Error("選択肢から回答してください。");
    }
    if (
      question.type === "checkbox" &&
      values.some((item) => !question.options.includes(item))
    ) {
      throw new Error("選択肢から回答してください。");
    }
    return {
      questionId: question.id,
      label: question.label,
      type: question.type,
      value,
    };
  });

  return {
    id: id("response"),
    surveyId: survey.id,
    surveyTitle: survey.title,
    customerName,
    customerEmail: "",
    answers,
    status: "new",
    adminMemo: "",
    submittedAt: now(),
  };
}

function normalizeAdminEditedAnswers(survey, existingAnswers, payloadAnswers) {
  const existingMap = new Map(
    (Array.isArray(existingAnswers) ? existingAnswers : []).map((answer) => [
      answer.questionId,
      answer,
    ]),
  );
  const answerMap = new Map(
    (Array.isArray(payloadAnswers) ? payloadAnswers : []).map((answer) => [
      answer.questionId,
      answer,
    ]),
  );
  const rawAnswerMap = buildRawAnswerMap(payloadAnswers);

  return survey.questions.map((question) => {
    const existingAnswer = existingMap.get(question.id) || {
      questionId: question.id,
      label: question.label,
      type: question.type,
      value: "",
    };
    const visible = isQuestionVisible(question, rawAnswerMap, survey);
    const required = isQuestionRequired(question, visible, survey);

    if (question.type === "photo") {
      return visible ? existingAnswer : { ...existingAnswer, value: "", files: [] };
    }

    const rawAnswer = answerMap.get(question.id) || {};
    const rawValue = rawAnswer.value;
    const values = Array.isArray(rawValue)
      ? rawValue.map(normalizeText).filter(Boolean)
      : [normalizeText(rawValue)].filter(Boolean);
    const value = visible ? values.join(", ") : "";

    if (required && !value) {
      throw new Error("未回答の質問があります。");
    }
    if (question.type === "rating" && value && !["1", "2", "3", "4", "5"].includes(value)) {
      throw new Error("評価は1から5で回答してください。");
    }
    if (question.type === "choice" && value && !question.options.includes(value)) {
      throw new Error("選択肢から回答してください。");
    }
    if (
      question.type === "checkbox" &&
      values.some((item) => !question.options.includes(item))
    ) {
      throw new Error("選択肢から回答してください。");
    }

    return {
      questionId: question.id,
      label: question.label,
      type: question.type,
      value,
    };
  });
}

function getRouteId(pathname, prefix) {
  if (!pathname.startsWith(prefix)) return "";
  return decodeURIComponent(pathname.slice(prefix.length));
}

async function handleApi(req, res, pathname, searchParams) {
  if (req.method === "OPTIONS") {
    sendJson(res, 204, {});
    return;
  }

  try {
    if (req.method === "GET" && pathname === "/api/health") {
      sendJson(res, 200, { ok: true });
      return;
    }

    if (req.method === "GET" && pathname === "/api/public/surveys") {
      const db = await readDb();
      sendJson(res, 200, {
        surveys: db.surveys.filter((survey) => survey.status === "published"),
      });
      return;
    }

    if (req.method === "GET" && pathname === "/api/public/bijiris-posts") {
      const db = await readDb();
      const posts = db.bijirisPosts
        .filter((post) => normalizeBijirisPostStatus(post.status) === "published")
        .sort((a, b) => {
          if (Boolean(b?.pinned) !== Boolean(a?.pinned)) return b?.pinned ? 1 : -1;
          return new Date(b?.publishedAt || b?.updatedAt || 0) - new Date(a?.publishedAt || a?.updatedAt || 0);
        });
      sendJson(res, 200, { posts });
      return;
    }

    if (req.method === "GET" && pathname === "/api/public/responses") {
      const name = normalizeText(searchParams.get("name"));
      const db = await readDb();
      sendJson(res, 200, {
        responses: name
          ? db.responses.filter((response) => response.customerName === name)
          : [],
        measurements: [],
        customerProfile: null,
      });
      return;
    }

    if (req.method === "POST" && pathname === "/api/public/responses") {
      const payload = await readBody(req);
      const result = await updateDb((db) => {
        const response = validateResponsePayload(db, payload);
        db.responses.push(response);
        return { response };
      });
      sendJson(res, 201, result);
      return;
    }

    if (req.method === "POST" && pathname === "/api/admin/login") {
      const payload = await readBody(req);
      const db = await readDb();
      const settings = getAdminSettings(db);
      if (
        normalizeText(payload.loginId) !== settings.adminUsername ||
        String(payload.password ?? "") !== settings.adminPassword
      ) {
        sendError(res, 401, "ログインIDまたはパスワードが違います。");
        return;
      }

      const exp = Math.floor(Date.now() / 1000) + TOKEN_TTL_SECONDS;
      const token = signToken({ sub: "admin", exp });
      sendJson(res, 200, { token, expiresAt: new Date(exp * 1000).toISOString() });
      return;
    }

    if (pathname.startsWith("/api/admin/") && !requireAdmin(req, res)) {
      return;
    }

    if (req.method === "GET" && pathname === "/api/admin/info") {
      const db = await readDb();
      sendJson(res, 200, buildAdminInfo(db));
      return;
    }

    if (req.method === "GET" && pathname === "/api/admin/surveys") {
      const db = await readDb();
      sendJson(res, 200, { surveys: db.surveys });
      return;
    }

    if (req.method === "GET" && pathname === "/api/admin/bijiris-posts") {
      const db = await readDb();
      sendJson(res, 200, { posts: db.bijirisPosts });
      return;
    }

    if (req.method === "POST" && pathname === "/api/admin/bijiris-posts") {
      const payload = await readBody(req);
      const result = await updateDb((db) => {
        const post = validateBijirisPostPayload(payload);
        db.bijirisPosts.unshift(post);
        return { post };
      });
      sendJson(res, 201, result);
      return;
    }

    if (req.method === "POST" && pathname === "/api/admin/surveys") {
      const payload = await readBody(req);
      const result = await updateDb((db) => {
        const survey = validateSurveyPayload(payload);
        db.surveys.unshift(survey);
        return { survey };
      });
      sendJson(res, 201, result);
      return;
    }

    if (req.method === "PUT" && pathname.startsWith("/api/admin/surveys/")) {
      const surveyId = getRouteId(pathname, "/api/admin/surveys/");
      const payload = await readBody(req);
      const result = await updateDb((db) => {
        const existing = db.surveys.find((survey) => survey.id === surveyId);
        if (!existing) throw new Error("アンケートが見つかりません。");
        const survey = validateSurveyPayload(payload, existing);
        db.surveys = db.surveys.map((item) => (item.id === surveyId ? survey : item));
        return { survey };
      });
      sendJson(res, 200, result);
      return;
    }

    if (req.method === "DELETE" && pathname.startsWith("/api/admin/surveys/")) {
      const surveyId = getRouteId(pathname, "/api/admin/surveys/");
      const result = await updateDb((db) => {
        const before = db.surveys.length;
        db.surveys = db.surveys.filter((survey) => survey.id !== surveyId);
        if (db.surveys.length === before) throw new Error("アンケートが見つかりません。");
        return { ok: true };
      });
      sendJson(res, 200, result);
      return;
    }

    if (req.method === "POST" && pathname === "/api/admin/bijiris-posts/replace") {
      const payload = await readBody(req);
      const result = await updateDb((db) => {
        const posts = Array.isArray(payload.posts) ? payload.posts : [];
        db.bijirisPosts = posts
          .map((post) => validateBijirisPostPayload(post, post))
          .filter(Boolean);
        return { posts: db.bijirisPosts };
      });
      sendJson(res, 200, result);
      return;
    }

    if (req.method === "PUT" && pathname.startsWith("/api/admin/bijiris-posts/")) {
      const postId = getRouteId(pathname, "/api/admin/bijiris-posts/");
      const payload = await readBody(req);
      const result = await updateDb((db) => {
        const existing = db.bijirisPosts.find((post) => post.id === postId);
        if (!existing) throw new Error("投稿が見つかりません。");
        const post = validateBijirisPostPayload(payload, existing);
        db.bijirisPosts = db.bijirisPosts.map((item) => (item.id === postId ? post : item));
        return { post };
      });
      sendJson(res, 200, result);
      return;
    }

    if (req.method === "DELETE" && pathname.startsWith("/api/admin/bijiris-posts/")) {
      const postId = getRouteId(pathname, "/api/admin/bijiris-posts/");
      const result = await updateDb((db) => {
        const before = db.bijirisPosts.length;
        db.bijirisPosts = db.bijirisPosts.filter((post) => post.id !== postId);
        if (db.bijirisPosts.length === before) throw new Error("投稿が見つかりません。");
        return { ok: true };
      });
      sendJson(res, 200, result);
      return;
    }

    if (req.method === "GET" && pathname === "/api/admin/responses") {
      const db = await readDb();
      sendJson(res, 200, {
        responses: db.responses.map((response) => ({
          status: "new",
          adminMemo: "",
          ...response,
        })),
      });
      return;
    }

    if (req.method === "PUT" && pathname.startsWith("/api/admin/responses/")) {
      const responseId = getRouteId(pathname, "/api/admin/responses/");
      const payload = await readBody(req);
      const result = await updateDb((db) => {
        const existing = db.responses.find((response) => response.id === responseId);
        if (!existing) throw new Error("回答が見つかりません。");
        const survey = db.surveys.find((item) => item.id === existing.surveyId);
        if (!survey) throw new Error("アンケートが見つかりません。");
        const updated = {
          ...existing,
          status: normalizeResponseStatus(payload.status),
          adminMemo: normalizeText(payload.adminMemo),
          answers: normalizeAdminEditedAnswers(survey, existing.answers, payload.answers),
          managedAt: now(),
        };
        db.responses = db.responses.map((response) =>
          response.id === responseId ? updated : response,
        );
        return { response: updated };
      });
      sendJson(res, 200, result);
      return;
    }

    if (req.method === "DELETE" && pathname.startsWith("/api/admin/responses/")) {
      const responseId = getRouteId(pathname, "/api/admin/responses/");
      const result = await updateDb((db) => {
        const before = db.responses.length;
        db.responses = db.responses.filter((response) => response.id !== responseId);
        if (db.responses.length === before) throw new Error("回答が見つかりません。");
        return { ok: true };
      });
      sendJson(res, 200, result);
      return;
    }

    if (req.method === "GET" && pathname === "/api/admin/export") {
      const db = await readDb();
      sendJson(res, 200, { ...db, exportedAt: now() });
      return;
    }

    if (req.method === "PUT" && pathname === "/api/admin/credentials") {
      const payload = await readBody(req);
      const result = await updateDb((db) => {
        const settings = getAdminSettings(db);
        const loginId = normalizeText(payload.loginId);
        const password = String(payload.password || "");

        if (!loginId) throw new Error("ログインIDを入力してください。");
        settings.adminUsername = loginId;
        settings.adminPassword = password || settings.adminPassword;

        if (String(settings.adminPassword).length < 4) {
          throw new Error("パスワードは4文字以上で入力してください。");
        }

        db.settings = settings;
        return { ok: true, adminInfo: buildAdminInfo(db) };
      });
      sendJson(res, 200, result);
      return;
    }

    sendError(res, 404, "API が見つかりません。");
  } catch (error) {
    const message =
      error instanceof SyntaxError
        ? "JSON の形式が正しくありません。"
        : error.message || "サーバーエラーが発生しました。";
    sendError(res, 400, message);
  }
}

function contentTypeFor(filePath) {
  const ext = path.extname(filePath);
  return (
    {
      ".html": "text/html; charset=utf-8",
      ".js": "text/javascript; charset=utf-8",
      ".css": "text/css; charset=utf-8",
      ".json": "application/json; charset=utf-8",
      ".webmanifest": "application/manifest+json; charset=utf-8",
      ".svg": "image/svg+xml",
      ".png": "image/png",
      ".ico": "image/x-icon",
    }[ext] || "application/octet-stream"
  );
}

async function serveStatic(req, res, pathname) {
  if (pathname === "/customer") {
    res.writeHead(302, { Location: "/customer/" });
    res.end();
    return;
  }

  if (pathname === "/admin") {
    res.writeHead(302, { Location: "/admin/" });
    res.end();
    return;
  }

  const appEntry = Object.entries(STATIC_APPS).find(([prefix]) =>
    pathname === `${prefix}/` || pathname.startsWith(`${prefix}/`),
  );

  if (appEntry) {
    const [prefix, appDir] = appEntry;
    const relativePath = pathname.slice(prefix.length) || "/";
    await serveStaticFromDir(res, appDir, relativePath);
    return;
  }

  const requestedPath = pathname === "/" ? "/index.html" : pathname;
  const decodedPath = decodeURIComponent(requestedPath);
  if (
    decodedPath.startsWith("/data/") || decodedPath.includes("..")
  ) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  const filePath = path.join(ROOT_DIR, decodedPath);
  if (!filePath.startsWith(ROOT_DIR)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  try {
    const file = await fs.readFile(filePath);
    res.writeHead(200, {
      "Content-Type": contentTypeFor(filePath),
      "Cache-Control": filePath.endsWith("index.html")
        ? "no-cache"
        : "public, max-age=3600",
    });
    res.end(file);
  } catch {
    const index = await fs.readFile(path.join(ROOT_DIR, "index.html"));
    res.writeHead(200, {
      "Content-Type": "text/html; charset=utf-8",
      "Cache-Control": "no-cache",
    });
    res.end(index);
  }
}

async function serveStaticFromDir(res, baseDir, requestedPath) {
  const relative = requestedPath === "/" ? "/index.html" : requestedPath;
  const decodedPath = decodeURIComponent(relative);
  if (decodedPath.includes("..")) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  const filePath = path.join(baseDir, decodedPath);
  if (!filePath.startsWith(baseDir)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  try {
    const file = await fs.readFile(filePath);
    res.writeHead(200, {
      "Content-Type": contentTypeFor(filePath),
      "Cache-Control": filePath.endsWith("index.html")
        ? "no-cache"
        : "public, max-age=3600",
    });
    res.end(file);
  } catch {
    const index = await fs.readFile(path.join(baseDir, "index.html"));
    res.writeHead(200, {
      "Content-Type": "text/html; charset=utf-8",
      "Cache-Control": "no-cache",
    });
    res.end(index);
  }
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);
  if (url.pathname.startsWith("/api/")) {
    await handleApi(req, res, url.pathname, url.searchParams);
    return;
  }
  await serveStatic(req, res, url.pathname);
});

ensureDb()
  .then(() => {
    server.listen(PORT, HOST, () => {
      const localUrl = `http://localhost:${PORT}`;
      const lanUrls = getLanUrls();
      console.log("Mayumi Bijiris app is running.");
      console.log(`Customer app: ${localUrl}/customer-app/`);
      console.log(`Admin app: ${localUrl}/admin-app/`);
      if (lanUrls.length) {
        console.log("Same Wi-Fi customer app URLs:");
        lanUrls.forEach((url) => {
          console.log(`- ${url}/customer-app/`);
        });
        console.log("Same Wi-Fi admin app URLs:");
        lanUrls.forEach((url) => {
          console.log(`- ${url}/admin-app/`);
        });
      }
    });
  })
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });
