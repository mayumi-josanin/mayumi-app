const CUSTOMER_KEY = "mayumi_survey_customer";
const DRAFTS_KEY = "mayumi_survey_drafts";
const PENDING_KEY = "mayumi_survey_pending_submission";
const PHOTO_FILE_LIMIT = 6;
const PHOTO_MAX_SIZE = 1400;
const PHOTO_JPEG_QUALITY = 0.74;
const RESPONSE_EDIT_WINDOW_MS = 24 * 60 * 60 * 1000;
const APP_VERSION = "20260411-6";
const TICKET_END_SURVEY_ID = "survey_bijiris_ticket_end";
const TICKET_END_COUNT_QUESTION_ID = "q_ticket_end_ticket_size";
const TICKET_END_SHEET_QUESTION_ID = "q_ticket_end_ticket_sheet";
const TICKET_END_ROUND_QUESTION_ID = "q_ticket_end_ticket_round";

const appState = {
  customer: loadLocal(CUSTOMER_KEY, { name: "" }),
  drafts: loadLocal(DRAFTS_KEY, {}),
  pendingSubmission: loadLocal(PENDING_KEY, null),
  surveys: [],
  history: [],
  publicInfo: {
    dataPolicyText: "",
    version: "",
  },
  selectedSurveyId: "",
  installPrompt: null,
  panelMode: "form",
  confirmPayload: null,
  editingResponseId: "",
  lastSubmittedResponse: null,
  lastSubmissionWasEdit: false,
};

const api = window.MayumiSurveyApi;
const toast = document.querySelector("#toast");
const surveyList = document.querySelector("#surveyList");
const answerPanel = document.querySelector("#answerPanel");
const historyList = document.querySelector("#historyList");
const customerForm = document.querySelector("#customerForm");
const installButton = document.querySelector("#installButton");

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

function getSelectedSurvey() {
  return appState.surveys.find((survey) => survey.id === appState.selectedSurveyId) || null;
}

function getTicketEndCountSelection(surveyId) {
  return normalizeText(getDraftValue(surveyId, TICKET_END_COUNT_QUESTION_ID));
}

function getTicketEndSheetSelection(surveyId) {
  return normalizeText(getDraftValue(surveyId, TICKET_END_SHEET_QUESTION_ID));
}

function getTicketEndRoundSelection(surveyId) {
  return normalizeText(getDraftValue(surveyId, TICKET_END_ROUND_QUESTION_ID));
}

function clearTicketEndSelections(surveyId) {
  const draft = getSurveyDraft(surveyId);
  delete draft.values[TICKET_END_COUNT_QUESTION_ID];
  delete draft.values[TICKET_END_SHEET_QUESTION_ID];
  delete draft.values[TICKET_END_ROUND_QUESTION_ID];
  setSurveyDraft(surveyId, draft);
}

function getTicketEndCountQuestion(survey) {
  return survey.questions.find((question) => question.id === TICKET_END_COUNT_QUESTION_ID);
}

function getTicketEndSheetQuestion(survey) {
  return survey.questions.find((question) => question.id === TICKET_END_SHEET_QUESTION_ID);
}

function getTicketEndRoundQuestion(survey) {
  return survey.questions.find((question) => question.id === TICKET_END_ROUND_QUESTION_ID);
}

function getTicketRoundOptions(ticketCount) {
  const max = ticketCount === "10回券" ? 10 : ticketCount === "6回券" ? 6 : 0;
  return Array.from({ length: max }, (_, index) => `${index + 1}回目`);
}

function getQuestionLabel(question, surveyId = appState.selectedSurveyId) {
  if (question.id === "q_ticket_end_photo_last") {
    const ticketRound = getTicketEndRoundSelection(surveyId);
    if (ticketRound) return `計測写真(${ticketRound})`;
    const ticketCount = getTicketEndCountSelection(surveyId);
    if (ticketCount === "6回券") return "計測写真(6回目)";
    if (ticketCount === "10回券") return "計測写真(10回目)";
  }
  return question.label;
}

function getSurveyAvailability(survey) {
  const nowTime = Date.now();
  if (survey?.acceptingResponses === false) {
    return { open: false, label: "受付停止中", detail: "現在このアンケートは停止しています。" };
  }
  if (survey?.startAt && new Date(survey.startAt).getTime() > nowTime) {
    return {
      open: false,
      label: "受付開始前",
      detail: `開始予定: ${formatDate(survey.startAt)}`,
    };
  }
  if (survey?.endAt && new Date(survey.endAt).getTime() < nowTime) {
    return {
      open: false,
      label: "受付終了",
      detail: `終了日時: ${formatDate(survey.endAt)}`,
    };
  }
  return {
    open: true,
    label: "受付中",
    detail: survey?.endAt ? `受付終了: ${formatDate(survey.endAt)}` : "期限指定なし",
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

function isQuestionVisible(question, answerMap) {
  if (!question?.visibleWhen) return true;
  const values = Array.isArray(answerMap[question.visibleWhen.questionId])
    ? answerMap[question.visibleWhen.questionId]
    : [];
  return values.includes(normalizeText(question.visibleWhen.value));
}

function getVisibleQuestions(survey, draft) {
  const answerMap = buildDraftAnswerMap(survey, draft);
  return survey.questions.filter((question) => {
    if (
      question.id === TICKET_END_COUNT_QUESTION_ID ||
      question.id === TICKET_END_SHEET_QUESTION_ID ||
      question.id === TICKET_END_ROUND_QUESTION_ID
    ) {
      return false;
    }
    return isQuestionVisible(question, answerMap);
  });
}

function isQuestionAnswered(question, surveyId, draft) {
  if (question.type === "photo") {
    return getDraftPhotos(surveyId, question.id).length > 0;
  }
  const raw = draft.values[question.id];
  if (Array.isArray(raw)) return raw.some((item) => normalizeText(item));
  return Boolean(normalizeText(raw));
}

function getProgress(survey, draft) {
  const visibleQuestions = getVisibleQuestions(survey, draft);
  const requiredQuestions = visibleQuestions.filter((question) => question.required !== false);
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
  const navPage = page === "survey" ? "home" : page;
  document.querySelectorAll(".page").forEach((node) => {
    node.classList.toggle("active", node.id === `page-${page}`);
  });
  document.querySelectorAll(".nav-button").forEach((button) => {
    button.classList.toggle("active", button.dataset.page === navPage);
  });
  if (page === "history") {
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
  if (!appState.customer.name) {
    historyList.innerHTML = `<div class="empty">先にお客様情報を保存してください。</div>`;
    return;
  }

  historyList.innerHTML = `<div class="empty">読み込み中です。</div>`;
  try {
    const result = await api.request(
      `/api/public/responses?name=${encodeURIComponent(appState.customer.name)}`,
    );
    appState.history = Array.isArray(result.responses) ? result.responses : [];
    renderHistory();
  } catch (error) {
    reportClientError("customer.loadHistory", error, { name: appState.customer.name });
    historyList.innerHTML = `<div class="empty">${escapeHtml(error.message || "履歴を読み込めませんでした。")}</div>`;
  }
}

function renderCustomerForm() {
  customerForm.elements.name.value = appState.customer.name || "";
}

function selectSurvey(surveyId) {
  appState.selectedSurveyId = surveyId;
  appState.panelMode = "form";
  appState.confirmPayload = null;
  appState.editingResponseId = "";
  appState.lastSubmittedResponse = null;
  appState.lastSubmissionWasEdit = false;
  renderSurveys();
  renderAnswerPanel();
  setPage("survey");
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
            class="survey-card ${survey.id === appState.selectedSurveyId ? "active" : ""} ${availability.open ? "" : "disabled"}"
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
  const required = question.required !== false;
  const label = `
    <span>${index + 1}. ${escapeHtml(getQuestionLabel(question, surveyId))}</span>
    ${required ? "" : `<span class="meta">任意</span>`}
  `;

  if (question.type === "textarea") {
    return `
      <label class="question-block">
        ${label}
        <textarea name="${name}" data-question-id="${question.id}">${escapeHtml(draft.values[question.id] || "")}</textarea>
      </label>
    `;
  }

  if (question.type === "rating") {
    return `
      <fieldset class="question-block">
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
      <fieldset class="question-block">
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
    const selected = new Set(Array.isArray(draft.values[question.id]) ? draft.values[question.id] : []);
    return `
      <fieldset class="question-block">
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
    return `
      <div class="question-block">
        <div class="question-label">${label}</div>
        <input type="file" name="${name}" accept="image/*" multiple hidden data-photo-input="${question.id}" />
        <div class="photo-actions">
          <button class="secondary-button" type="button" data-photo-pick="${question.id}" data-photo-mode="append">
            ${files.length ? "写真を追加" : "写真を選ぶ"}
          </button>
          ${files.length ? `<button class="ghost-button" type="button" data-photo-pick="${question.id}" data-photo-mode="replace">撮り直す</button>` : ""}
        </div>
        <span class="field-help">スマホ内の写真を選択してください。最大 ${PHOTO_FILE_LIMIT} 枚まで添付できます。</span>
        ${
          files.length
            ? `
              <div class="photo-list editable">
                ${files
                  .map((file, fileIndex) => {
                    const preview =
                      file.previewUrl || file.thumbnailUrl || file.dataUrl || file.url || "";
                    return `
                      <div class="photo-thumb editable-thumb">
                        ${preview ? `<img src="${escapeHtml(preview)}" alt="${escapeHtml(file.name || "photo")}" />` : ""}
                        <span>${escapeHtml(file.name || `写真${fileIndex + 1}`)}</span>
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
    <label class="question-block">
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
            const preview = file.previewUrl || file.thumbnailUrl || file.dataUrl || file.url || "";
            return `
              <div class="photo-thumb">
                ${preview ? `<img src="${escapeHtml(preview)}" alt="${escapeHtml(file.name || "photo")}" />` : ""}
                <span>${escapeHtml(file.name || "写真")}</span>
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
      <p>${appState.lastSubmissionWasEdit ? "回答を更新しました。" : "回答を送信しました。"} 履歴からいつでも確認できます。</p>
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
      <span class="badge ${availability.open ? "open" : "closed"}">${escapeHtml(availability.label)}</span>
      <span class="meta">${escapeHtml(availability.detail)}</span>
    </div>
    ${renderProgressBar(progress)}
    ${
      isTicketEndSurvey(survey)
        ? `
          <div class="question-block prefilled-answer">
            <strong>回数券の種類</strong>
            <div>${escapeHtml(getTicketEndCountSelection(surveyId))}</div>
            <strong>回数券の何枚目</strong>
            <div>${escapeHtml(getTicketEndSheetSelection(surveyId))}</div>
            <strong>回数券の何回目</strong>
            <div>${escapeHtml(getTicketEndRoundSelection(surveyId))}</div>
            <button id="changeTicketCountButton" class="ghost-button" type="button">変更する</button>
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
  document.querySelector("#changeTicketCountButton")?.addEventListener("click", () => {
    clearTicketEndSelections(surveyId);
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

  if (isTicketEndSurvey(survey) && !getTicketEndCountSelection(surveyId)) {
    renderTicketStepCount(survey, surveyId);
    return;
  }
  if (isTicketEndSurvey(survey) && !getTicketEndSheetSelection(surveyId)) {
    renderTicketStepSheet(survey, surveyId);
    return;
  }
  if (isTicketEndSurvey(survey) && !getTicketEndRoundSelection(surveyId)) {
    renderTicketStepRound(survey, surveyId);
    return;
  }

  renderFormPanel(survey, draft);
}

function attachAnswerFormHandlers(form, survey) {
  if (!form) return;

  form.addEventListener("input", (event) => {
    const input = event.target;
    const questionId = input.dataset.questionId;
    if (!questionId) return;

    if (input.type === "checkbox") {
      const values = Array.from(form.querySelectorAll(`[data-question-id="${questionId}"]:checked`)).map(
        (node) => node.value,
      );
      updateDraftValue(survey.id, questionId, values);
      return;
    }

    if (input.type === "radio") {
      updateDraftValue(survey.id, questionId, input.value);
      renderAnswerPanel();
      return;
    }

    updateDraftValue(survey.id, questionId, input.value);
    const progress = getProgress(survey, getSurveyDraft(survey.id));
    document.querySelector(".progress-fill")?.setAttribute("style", `width: ${progress.percent}%`);
    const progressHead = document.querySelector(".progress-head span");
    if (progressHead) {
      progressHead.textContent = `${progress.answeredRequired}/${progress.requiredTotal} 必須回答`;
    }
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

  form.querySelectorAll("[data-photo-input]").forEach((input) => {
    input.addEventListener("change", async () => {
      const questionId = input.dataset.photoInput;
      const files = Array.from(input.files || []);
      if (!files.length) return;

      const existing = getDraftPhotos(survey.id, questionId);
      const nextFiles = input.dataset.photoMode === "replace" ? [] : existing.slice();

      try {
        for (const file of files.slice(0, PHOTO_FILE_LIMIT)) {
          nextFiles.push(await preparePhotoFile(file));
        }
        updateDraftPhotos(survey.id, questionId, nextFiles.slice(0, PHOTO_FILE_LIMIT));
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
      showToast(error.message || "回答内容を確認できませんでした。");
    }
  });
}

function prepareSubmissionFromDraft(survey, draft) {
  const answerMap = buildDraftAnswerMap(survey, draft);
  const answers = [];
  const summary = [];

  for (const question of survey.questions) {
    const visible =
      question.id === TICKET_END_COUNT_QUESTION_ID ||
      question.id === TICKET_END_SHEET_QUESTION_ID ||
      question.id === TICKET_END_ROUND_QUESTION_ID
        ? true
        : isQuestionVisible(question, answerMap);

    if (question.type === "photo") {
      const files = visible ? getDraftPhotos(survey.id, question.id) : [];
      if (visible && question.required !== false && !files.length) {
        throw new Error("未回答の質問があります。");
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
      if (visible && question.required !== false && !value.length) {
        throw new Error("未回答の質問があります。");
      }
      if (visible && value.some((item) => !question.options.includes(item))) {
        throw new Error("選択肢から回答してください。");
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
    if (visible && question.required !== false && !value) {
      throw new Error("未回答の質問があります。");
    }
    if (visible && question.type === "rating" && value && !["1", "2", "3", "4", "5"].includes(value)) {
      throw new Error("5段階評価の選択が正しくありません。");
    }
    if (visible && question.type === "choice" && value && !question.options.includes(value)) {
      throw new Error("選択肢から回答してください。");
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

  return { answers, summary };
}

async function submitPreparedAnswer() {
  const survey = getSelectedSurvey();
  if (!survey || !appState.confirmPayload) return;
  if (!appState.customer.name) {
    showToast("先にお客様情報を保存してください。");
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
            const href = file.previewUrl || file.url || file.dataUrl || "#";
            const preview = file.previewUrl || file.thumbnailUrl || file.dataUrl || file.url || "";
            return `
              <a class="photo-thumb" href="${escapeHtml(href)}" target="_blank" rel="noopener">
                ${preview ? `<img src="${escapeHtml(preview)}" alt="${escapeHtml(file.name || "photo")}" />` : ""}
                <span>${escapeHtml(file.name || "写真")}</span>
              </a>
            `;
          })
          .join("")}
      </div>
    `;
  }
  return escapeHtml(answer.value || "未回答");
}

function renderHistory() {
  if (!appState.history.length) {
    historyList.innerHTML = `<div class="empty">まだ回答履歴はありません。</div>`;
    return;
  }

  historyList.innerHTML = `
    ${renderPendingNotice()}
    ${appState.history
      .slice()
      .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))
      .map(
        (response) => `
          <article class="history-card">
            <div class="section-head">
              <div>
                <strong>${escapeHtml(response.surveyTitle)}</strong>
                <div class="meta">${formatDate(response.submittedAt)}</div>
              </div>
              ${canEditResponse(response) ? `<button class="secondary-button" type="button" data-edit-response="${response.id}">修正する</button>` : ""}
            </div>
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
                ? `<div class="meta">回答後24時間以内は修正できます。</div>`
                : `<div class="meta">修正可能期間を過ぎています。</div>`
            }
          </article>
        `,
      )
      .join("")}
  `;

  historyList.querySelectorAll("[data-edit-response]").forEach((button) => {
    button.addEventListener("click", () => startResponseEdit(button.dataset.editResponse));
  });
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

function setupInstall() {
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker
        .register("./sw.js?v=20260411-6", { updateViaCache: "none" })
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

customerForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const formData = new FormData(event.currentTarget);
  appState.customer = {
    name: normalizeText(formData.get("name")),
  };
  saveLocal(CUSTOMER_KEY, appState.customer);
  showToast("お客様情報を保存しました。");
  void loadHistory();
});

document.querySelectorAll("[data-page]").forEach((button) => {
  button.addEventListener("click", () => setPage(button.dataset.page));
});

document.querySelector("#refreshButton").addEventListener("click", () => {
  void loadSurveys();
});

document.querySelector("#historyRefreshButton").addEventListener("click", () => {
  void loadHistory();
});

setupInstall();
window.addEventListener("error", (event) => {
  reportClientError("customer.window", event.error || event.message || "window error");
});
window.addEventListener("unhandledrejection", (event) => {
  reportClientError("customer.promise", event.reason || "unhandled rejection");
});
renderCustomerForm();
renderSurveys();
renderAnswerPanel();
void loadSurveys();
if (appState.customer.name) {
  void loadHistory();
}
