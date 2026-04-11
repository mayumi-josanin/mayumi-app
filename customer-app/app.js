const CUSTOMER_KEY = "mayumi_survey_customer";
const PHOTO_FILE_LIMIT = 6;
const PHOTO_MAX_SIZE = 1400;
const PHOTO_JPEG_QUALITY = 0.74;
const TICKET_END_SURVEY_ID = "survey_bijiris_ticket_end";
const TICKET_END_COUNT_QUESTION_ID = "q_ticket_end_ticket_size";
const TICKET_END_SHEET_QUESTION_ID = "q_ticket_end_ticket_sheet";
const TICKET_END_ROUND_QUESTION_ID = "q_ticket_end_ticket_round";

const appState = {
  customer: loadLocal(CUSTOMER_KEY, { name: "" }),
  surveys: [],
  history: [],
  selectedSurveyId: "",
  prefilledAnswers: {},
  installPrompt: null,
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
  localStorage.setItem(key, JSON.stringify(value));
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
      reject(new Error("JPEG、PNG、WEBPの写真を選択してください。")),
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

function getFallbackSurveys() {
  const makeSurveys = window.MayumiDefaultSurveys;
  return typeof makeSurveys === "function"
    ? makeSurveys(new Date().toISOString()).filter((survey) => survey.status === "published")
    : [];
}

function isTicketEndSurvey(survey) {
  return survey?.id === TICKET_END_SURVEY_ID;
}

function getTicketEndCountSelection() {
  return appState.prefilledAnswers[TICKET_END_COUNT_QUESTION_ID] || "";
}

function getTicketEndSheetSelection() {
  return appState.prefilledAnswers[TICKET_END_SHEET_QUESTION_ID] || "";
}

function getTicketEndRoundSelection() {
  return appState.prefilledAnswers[TICKET_END_ROUND_QUESTION_ID] || "";
}

function clearTicketEndSelections() {
  delete appState.prefilledAnswers[TICKET_END_COUNT_QUESTION_ID];
  delete appState.prefilledAnswers[TICKET_END_SHEET_QUESTION_ID];
  delete appState.prefilledAnswers[TICKET_END_ROUND_QUESTION_ID];
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

function getQuestionLabel(question) {
  if (question.id === "q_ticket_end_photo_last") {
    const ticketRound = getTicketEndRoundSelection();
    if (ticketRound) return `計測写真(${ticketRound})`;
    const ticketCount = getTicketEndCountSelection();
    if (ticketCount === "6回券") return "計測写真(6回目)";
    if (ticketCount === "10回券") return "計測写真(10回目)";
  }
  return question.label;
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
    if (
      !appState.selectedSurveyId ||
      !appState.surveys.some((survey) => survey.id === appState.selectedSurveyId)
    ) {
      appState.selectedSurveyId = "";
    }
    renderSurveys();
    renderAnswerPanel();
  } catch (error) {
    const fallbackSurveys = getFallbackSurveys();
    if (fallbackSurveys.length) {
      appState.surveys = fallbackSurveys;
      appState.selectedSurveyId = "";
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
    appState.history = result.responses || [];
    renderHistory();
  } catch (error) {
    historyList.innerHTML = `<div class="empty">${escapeHtml(error.message || "履歴を読み込めませんでした。")}</div>`;
  }
}

function renderCustomerForm() {
  customerForm.elements.name.value = appState.customer.name || "";
}

function renderSurveys() {
  if (!appState.surveys.length) {
    surveyList.innerHTML = `<div class="empty">現在回答できるアンケートはありません。</div>`;
    return;
  }

  surveyList.innerHTML = appState.surveys
    .map(
      (survey) => `
        <button class="survey-card ${survey.id === appState.selectedSurveyId ? "active" : ""}" type="button" data-survey-id="${survey.id}">
          <strong>${escapeHtml(survey.title)}</strong>
          <span>${escapeHtml(survey.description)}</span>
          <span class="badge">${survey.questions.length}問</span>
        </button>
      `,
    )
    .join("");

  document.querySelectorAll("[data-survey-id]").forEach((button) => {
    button.addEventListener("click", () => {
      const survey = appState.surveys.find((item) => item.id === button.dataset.surveyId);
      appState.selectedSurveyId = button.dataset.surveyId;
      if (isTicketEndSurvey(survey)) {
        clearTicketEndSelections();
      }
      renderSurveys();
      renderAnswerPanel();
      setPage("survey");
    });
  });
}

function renderAnswerPanel() {
  const survey = appState.surveys.find((item) => item.id === appState.selectedSurveyId);
  if (!survey) {
    answerPanel.innerHTML = `
      <div class="empty">
        ホームに戻って、回答したいアンケートを選択してください。
      </div>
    `;
    return;
  }

  if (isTicketEndSurvey(survey) && !getTicketEndCountSelection()) {
    const ticketQuestion = getTicketEndCountQuestion(survey);
    answerPanel.innerHTML = `
      <div class="section-head survey-toolbar">
        <div>
          <h2>${escapeHtml(survey.title)}</h2>
          <p>最初に回数券の種類を選択してください。</p>
        </div>
        <button id="backToHomeButton" class="ghost-button" type="button">一覧へ戻る</button>
      </div>
      <form id="ticketCountForm" class="question-list">
        <fieldset class="question-block">
          <legend>
            <span>${escapeHtml(ticketQuestion.label)}</span>
          </legend>
          <div class="choice-row">
            ${ticketQuestion.options
              .map(
                (option) => `
                  <label>
                    <input type="radio" name="ticketCount" value="${escapeHtml(option)}" />
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

    document.querySelector("#backToHomeButton").addEventListener("click", () => {
      setPage("home");
    });
    document.querySelector("#ticketCountForm").addEventListener("submit", (event) => {
      event.preventDefault();
      const formData = new FormData(event.currentTarget);
      const ticketCount = String(formData.get("ticketCount") || "");
      if (!ticketCount) {
        showToast("回数券の種類を選択してください。");
        return;
      }
      appState.prefilledAnswers[TICKET_END_COUNT_QUESTION_ID] = ticketCount;
      delete appState.prefilledAnswers[TICKET_END_SHEET_QUESTION_ID];
      delete appState.prefilledAnswers[TICKET_END_ROUND_QUESTION_ID];
      renderAnswerPanel();
    });
    return;
  }

  if (isTicketEndSurvey(survey) && !getTicketEndSheetSelection()) {
    const ticketSheetQuestion = getTicketEndSheetQuestion(survey);
    const selectedTicketCount = getTicketEndCountSelection();
    answerPanel.innerHTML = `
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
          <legend>
            <span>${escapeHtml(ticketSheetQuestion.label)}</span>
          </legend>
          <div class="choice-row">
            ${ticketSheetQuestion.options
              .map(
                (option) => `
                  <label>
                    <input type="radio" name="ticketSheet" value="${escapeHtml(option)}" />
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
      clearTicketEndSelections();
      renderAnswerPanel();
    });
    document.querySelector("#ticketSheetForm").addEventListener("submit", (event) => {
      event.preventDefault();
      const formData = new FormData(event.currentTarget);
      const ticketSheet = String(formData.get("ticketSheet") || "");
      if (!ticketSheet) {
        showToast("回数券の何枚目かを選択してください。");
        return;
      }
      appState.prefilledAnswers[TICKET_END_SHEET_QUESTION_ID] = ticketSheet;
      delete appState.prefilledAnswers[TICKET_END_ROUND_QUESTION_ID];
      renderAnswerPanel();
    });
    return;
  }

  if (isTicketEndSurvey(survey) && !getTicketEndRoundSelection()) {
    const ticketRoundQuestion = getTicketEndRoundQuestion(survey);
    const selectedTicketCount = getTicketEndCountSelection();
    const selectedTicketSheet = getTicketEndSheetSelection();
    const roundOptions = getTicketRoundOptions(selectedTicketCount);
    answerPanel.innerHTML = `
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
          <legend>
            <span>${escapeHtml(ticketRoundQuestion.label)}</span>
          </legend>
          <div class="choice-row">
            ${roundOptions
              .map(
                (option) => `
                  <label>
                    <input type="radio" name="ticketRound" value="${escapeHtml(option)}" />
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
      delete appState.prefilledAnswers[TICKET_END_SHEET_QUESTION_ID];
      delete appState.prefilledAnswers[TICKET_END_ROUND_QUESTION_ID];
      renderAnswerPanel();
    });
    document.querySelector("#ticketRoundForm").addEventListener("submit", (event) => {
      event.preventDefault();
      const formData = new FormData(event.currentTarget);
      const ticketRound = String(formData.get("ticketRound") || "");
      if (!ticketRound) {
        showToast("回数券の何回目かを選択してください。");
        return;
      }
      appState.prefilledAnswers[TICKET_END_ROUND_QUESTION_ID] = ticketRound;
      renderAnswerPanel();
    });
    return;
  }

  const visibleQuestions = survey.questions.filter(
    (question) =>
      question.id !== TICKET_END_COUNT_QUESTION_ID &&
      question.id !== TICKET_END_SHEET_QUESTION_ID &&
      question.id !== TICKET_END_ROUND_QUESTION_ID,
  );

  answerPanel.innerHTML = `
    <div class="section-head survey-toolbar">
      <div>
        <h2>${escapeHtml(survey.title)}</h2>
        <p>${escapeHtml(survey.description)}</p>
      </div>
      <button id="backToHomeButton" class="ghost-button" type="button">一覧へ戻る</button>
    </div>
    ${
      isTicketEndSurvey(survey)
        ? `
          <div class="question-block prefilled-answer">
            <strong>回数券の種類</strong>
            <div>${escapeHtml(getTicketEndCountSelection())}</div>
            <strong>回数券の何枚目</strong>
            <div>${escapeHtml(getTicketEndSheetSelection())}</div>
            <strong>回数券の何回目</strong>
            <div>${escapeHtml(getTicketEndRoundSelection())}</div>
            <button id="changeTicketCountButton" class="ghost-button" type="button">変更する</button>
          </div>
        `
        : ""
    }
    <form id="answerForm" class="question-list">
      ${visibleQuestions.map((question, index) => renderQuestion(question, index)).join("")}
      <button class="primary-button" type="submit">回答を送信する</button>
    </form>
  `;

  document.querySelector("#backToHomeButton").addEventListener("click", () => {
    setPage("home");
  });
  document.querySelector("#changeTicketCountButton")?.addEventListener("click", () => {
    clearTicketEndSelections();
    renderAnswerPanel();
  });
  document.querySelector("#answerForm").addEventListener("submit", submitAnswer);
}

function renderQuestion(question, index) {
  const name = `question-${question.id}`;
  const required = question.required !== false;
  const requiredAttr = required ? "required" : "";
  const label = `
    <span>${index + 1}. ${escapeHtml(getQuestionLabel(question))}</span>
    ${required ? "" : `<span class="meta">任意</span>`}
  `;

  if (question.type === "textarea") {
    return `
      <label class="question-block">
        ${label}
        <textarea name="${name}" ${requiredAttr}></textarea>
      </label>
    `;
  }

  if (question.type === "rating") {
    return `
      <fieldset class="question-block">
        <legend>${label}</legend>
        <div class="rating-row">
          ${[1, 2, 3, 4, 5]
            .map((value) => `<label><input type="radio" name="${name}" value="${value}" ${requiredAttr} />${value}</label>`)
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
            .map((option) => `<label><input type="radio" name="${name}" value="${escapeHtml(option)}" ${requiredAttr} />${escapeHtml(option)}</label>`)
            .join("")}
        </div>
      </fieldset>
    `;
  }

  if (question.type === "checkbox") {
    return `
      <fieldset class="question-block">
        <legend>${label}</legend>
        <div class="checkbox-row">
          ${question.options
            .map((option) => `<label><input type="checkbox" name="${name}" value="${escapeHtml(option)}" />${escapeHtml(option)}</label>`)
            .join("")}
        </div>
      </fieldset>
    `;
  }

  if (question.type === "photo") {
    return `
      <label class="question-block">
        ${label}
        <input type="file" name="${name}" accept="image/*" multiple ${requiredAttr} />
        <span class="field-help">スマホ内の写真を選択してください。複数枚まとめて添付できます。</span>
      </label>
    `;
  }

  return `
    <label class="question-block">
      ${label}
      <input type="text" name="${name}" ${requiredAttr} />
    </label>
  `;
}

async function collectAnswers(form, survey) {
  const formData = new FormData(form);
  const answers = [];

  for (const question of survey.questions) {
    if (question.id === TICKET_END_COUNT_QUESTION_ID) {
      const ticketCount = getTicketEndCountSelection();
      if (question.required !== false && !ticketCount) {
        throw new Error("回数券の種類を選択してください。");
      }
      answers.push({
        questionId: question.id,
        value: ticketCount,
      });
      continue;
    }

    if (question.id === TICKET_END_SHEET_QUESTION_ID) {
      const ticketSheet = getTicketEndSheetSelection();
      if (question.required !== false && !ticketSheet) {
        throw new Error("回数券の何枚目かを選択してください。");
      }
      answers.push({
        questionId: question.id,
        value: ticketSheet,
      });
      continue;
    }

    if (question.id === TICKET_END_ROUND_QUESTION_ID) {
      const ticketRound = getTicketEndRoundSelection();
      if (question.required !== false && !ticketRound) {
        throw new Error("回数券の何回目かを選択してください。");
      }
      answers.push({
        questionId: question.id,
        value: ticketRound,
      });
      continue;
    }

    const name = `question-${question.id}`;
    if (question.type === "checkbox") {
      const values = formData.getAll(name).map((value) => String(value || ""));
      if (question.required !== false && !values.length) {
        throw new Error("未回答の質問があります。");
      }
      answers.push({ questionId: question.id, value: values });
      continue;
    }

    if (question.type === "photo") {
      const input = form.elements.namedItem(name);
      const selectedFiles = input?.files ? Array.from(input.files).slice(0, PHOTO_FILE_LIMIT) : [];
      if (question.required !== false && !selectedFiles.length) {
        throw new Error("未回答の質問があります。");
      }
      const files = [];
      for (const file of selectedFiles) {
        files.push(await preparePhotoFile(file));
      }
      answers.push({ questionId: question.id, files });
      continue;
    }

    answers.push({
      questionId: question.id,
      value: String(formData.get(name) || ""),
    });
  }

  return answers;
}

async function submitAnswer(event) {
  event.preventDefault();
  if (!appState.customer.name) {
    showToast("先にお客様情報を保存してください。");
    return;
  }

  const survey = appState.surveys.find((item) => item.id === appState.selectedSurveyId);
  if (!survey) return;

  const submitButton = event.currentTarget.querySelector('button[type="submit"]');
  submitButton.disabled = true;
  submitButton.textContent = "送信準備中です";

  try {
    const answers = await collectAnswers(event.currentTarget, survey);
    submitButton.textContent = "送信中です";
    await api.request("/api/public/responses", {
      method: "POST",
      body: {
        surveyId: survey.id,
        customer: appState.customer,
        answers,
      },
    });
    event.currentTarget.reset();
    if (isTicketEndSurvey(survey)) {
      clearTicketEndSelections();
      renderAnswerPanel();
    }
    showToast("回答を送信しました。");
    await loadHistory();
  } catch (error) {
    showToast(error.message || "回答を送信できませんでした。");
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "回答を送信する";
  }
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
                ${preview ? `<img src="${escapeHtml(preview)}" alt="${escapeHtml(file.name)}" />` : ""}
                <span>${escapeHtml(file.name)}</span>
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

  historyList.innerHTML = appState.history
    .slice()
    .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))
    .map(
      (response) => `
        <article class="history-card">
          <strong>${escapeHtml(response.surveyTitle)}</strong>
          <span class="meta">${formatDate(response.submittedAt)}</span>
          ${response.answers
            .map((answer) => `<div class="history-answer"><strong>${escapeHtml(answer.label)}</strong><br />${renderAnswerValue(answer)}</div>`)
            .join("")}
        </article>
      `,
    )
    .join("");
}

function setupInstall() {
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("./sw.js").catch(() => {});
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
    name: String(formData.get("name") || "").trim(),
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
renderCustomerForm();
void loadSurveys();
void loadHistory();
