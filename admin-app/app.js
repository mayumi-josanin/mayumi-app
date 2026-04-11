const TOKEN_KEY = "mayumi_survey_admin_token";
const STATUS_LABELS = {
  new: "未対応",
  checked: "確認済み",
  done: "対応済み",
};

const api = window.MayumiSurveyApi;
const state = {
  token: localStorage.getItem(TOKEN_KEY) || "",
  surveys: [],
  responses: [],
  adminInfo: null,
  selectedSurveyEditorId: "",
  selectedCustomerName: "",
  selectedResponseId: "",
  installPrompt: null,
};

const loginView = document.querySelector("#loginView");
const adminView = document.querySelector("#adminView");
const toast = document.querySelector("#toast");
const loginForm = document.querySelector("#loginForm");
const credentialForm = document.querySelector("#credentialForm");
const installButton = document.querySelector("#installButton");

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

function normalizeStatus(status) {
  return STATUS_LABELS[status] ? status : "new";
}

function normalizeSurveyStatus(status) {
  return status === "draft" ? "draft" : "published";
}

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("show");
  setTimeout(() => toast.classList.remove("show"), 2600);
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
}

function getFallbackSurveys() {
  const makeSurveys = window.MayumiDefaultSurveys;
  return typeof makeSurveys === "function" ? makeSurveys(new Date().toISOString()) : [];
}

async function loadAdminData() {
  const [adminInfoResult, surveysResult, responsesResult] = await Promise.all([
    api.request("/api/admin/info", { token: state.token }),
    api.request("/api/admin/surveys", { token: state.token }),
    api.request("/api/admin/responses", { token: state.token }),
  ]);
  state.adminInfo = adminInfoResult || null;
  state.surveys = (surveysResult.surveys || []).length
    ? surveysResult.surveys || []
    : getFallbackSurveys();
  state.responses = (responsesResult.responses || []).map((response) => ({
    status: "new",
    adminMemo: "",
    ...response,
  }));
  renderAll();
}

function renderAll() {
  renderDashboard();
  renderSurveyManager();
  renderFilters();
  renderResponses();
  renderSettings();
  document.querySelector("#adminUrlBox").textContent = window.location.href;
}

function renderDashboard() {
  const customers = groupByCustomer();
  const unread = state.responses.filter((response) => normalizeStatus(response.status) === "new").length;
  document.querySelector("#statsGrid").innerHTML = `
    <div class="stat-card"><span>回答数</span><strong>${state.responses.length}</strong></div>
    <div class="stat-card"><span>未対応</span><strong>${unread}</strong></div>
    <div class="stat-card"><span>回答者数</span><strong>${customers.length}</strong></div>
    <div class="stat-card"><span>アンケート数</span><strong>${state.surveys.length}</strong></div>
  `;

  const latest = state.responses
    .slice()
    .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))
    .slice(0, 5);
  document.querySelector("#latestResponses").innerHTML = latest.length
    ? latest.map(renderCompactResponse).join("")
    : `<div class="empty">まだ回答はありません。</div>`;

  const summary = state.surveys.map((survey) => ({
    title: survey.title,
    count: state.responses.filter((response) => response.surveyId === survey.id).length,
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
}

function renderCompactResponse(response) {
  const status = normalizeStatus(response.status);
  return `
    <article class="card">
      <strong>${escapeHtml(response.customerName)}</strong>
      <div class="meta">${escapeHtml(response.surveyTitle)} / ${formatDate(response.submittedAt)}</div>
      <span class="badge ${status}">${STATUS_LABELS[status]}</span>
    </article>
  `;
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
  if (Array.isArray(answer.files) && answer.files.length) {
    return `${answer.label}: ${answer.files
      .map((file) => file.url || file.previewUrl || file.name)
      .join(", ")}`;
  }
  return `${answer.label}: ${answer.value || ""}`;
}

function renderFilters() {
  const surveyFilter = document.querySelector("#surveyFilter");
  const current = surveyFilter.value;
  surveyFilter.innerHTML = `
    <option value="">すべて</option>
    ${state.surveys.map((survey) => `<option value="${survey.id}">${escapeHtml(survey.title)}</option>`).join("")}
  `;
  surveyFilter.value = current;
}

function getFilteredResponses() {
  const surveyId = document.querySelector("#surveyFilter").value;
  const status = document.querySelector("#statusFilter").value;
  const keyword = document.querySelector("#keywordFilter").value.trim().toLowerCase();

  return state.responses
    .filter((response) => !surveyId || response.surveyId === surveyId)
    .filter((response) => !status || normalizeStatus(response.status) === status)
    .filter((response) => {
      if (!keyword) return true;
      return (
        String(response.customerName || "").toLowerCase().includes(keyword) ||
        String(response.surveyTitle || "").toLowerCase().includes(keyword)
      );
    })
    .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt));
}

function renderResponses() {
  const responses = getFilteredResponses();
  const stage = document.querySelector("#responseStage");
  const customers = groupByCustomer(responses);

  if (state.selectedCustomerName && !customers.some((customer) => customer.name === state.selectedCustomerName)) {
    state.selectedCustomerName = "";
    state.selectedResponseId = "";
  }

  const customerResponses = state.selectedCustomerName
    ? responses.filter((response) => response.customerName === state.selectedCustomerName)
    : [];

  if (
    state.selectedResponseId &&
    !customerResponses.some((response) => response.id === state.selectedResponseId)
  ) {
    state.selectedResponseId = "";
  }

  const selectedResponse = customerResponses.find(
    (response) => response.id === state.selectedResponseId,
  );

  if (!state.selectedCustomerName) {
    stage.innerHTML = `
      <div class="stage-head">
        <div>
          <div class="card-title">お客様一覧</div>
          <div class="meta">最新の回答順で表示しています。</div>
        </div>
      </div>
      <div class="customer-list">
        ${
          customers.length
            ? customers
                .map(
                  (customer) => `
                    <button
                      class="customer-card selectable-card"
                      type="button"
                      data-customer-name="${escapeHtml(customer.name)}"
                    >
                      <strong>${escapeHtml(customer.name)}</strong>
                      <div>回答数: ${customer.count}件</div>
                      <div class="meta">最終回答: ${formatDate(customer.latestAt)}</div>
                    </button>
                  `,
                )
                .join("")
            : `<div class="empty">条件に一致するお客様はいません。</div>`
        }
      </div>
    `;
  } else if (!selectedResponse) {
    stage.innerHTML = `
      <div class="stage-head">
        <div>
          <div class="card-title">${escapeHtml(state.selectedCustomerName)} さんの回答履歴</div>
          <div class="meta">アンケートタイトルを押すと詳細を表示します。</div>
        </div>
        <button class="secondary-button" type="button" data-back-stage="customers">戻る</button>
      </div>
      <div class="response-list">
        ${
          customerResponses.length
            ? customerResponses
                .map(
                  (response) => `
                    <button
                      class="response-history-card selectable-card"
                      type="button"
                      data-history-response-id="${response.id}"
                    >
                      <strong>${escapeHtml(response.surveyTitle)}</strong>
                      <div class="meta">${formatDate(response.submittedAt)}</div>
                      <span class="badge ${normalizeStatus(response.status)}">
                        ${STATUS_LABELS[normalizeStatus(response.status)]}
                      </span>
                    </button>
                  `,
                )
                .join("")
            : `<div class="empty">このお客様の回答履歴はありません。</div>`
        }
      </div>
    `;
  } else {
    stage.innerHTML = `
      <div class="stage-head">
        <div>
          <div class="card-title">${escapeHtml(selectedResponse.surveyTitle)}</div>
          <div class="meta">${escapeHtml(selectedResponse.customerName)} / ${formatDate(selectedResponse.submittedAt)}</div>
        </div>
        <button class="secondary-button" type="button" data-back-stage="history">戻る</button>
      </div>
      ${renderResponseCard(selectedResponse)}
    `;
  }

  stage.querySelectorAll("[data-customer-name]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedCustomerName = button.dataset.customerName;
      state.selectedResponseId = "";
      renderResponses();
    });
  });

  stage.querySelectorAll("[data-history-response-id]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedResponseId = button.dataset.historyResponseId;
      renderResponses();
    });
  });

  stage.querySelectorAll("[data-back-stage]").forEach((button) => {
    button.addEventListener("click", () => {
      if (button.dataset.backStage === "customers") {
        state.selectedCustomerName = "";
        state.selectedResponseId = "";
      } else {
        state.selectedResponseId = "";
      }
      renderResponses();
    });
  });

  stage.querySelectorAll("[data-save-response]").forEach((button) => {
    button.addEventListener("click", () => {
      void saveResponseManagement(button.dataset.saveResponse);
    });
  });

  stage.querySelectorAll("[data-delete-response]").forEach((button) => {
    button.addEventListener("click", () => {
      void deleteResponse(button.dataset.deleteResponse);
    });
  });
}

function renderResponseCard(response) {
  const status = normalizeStatus(response.status);
  const survey = findSurveyById(response.surveyId);
  const displayAnswers = getDisplayAnswers(response, survey);
  return `
    <article class="response-card" data-response-card="${response.id}">
      <div class="response-head">
        <div>
          <strong>${escapeHtml(response.customerName)}</strong>
          <div class="meta">${formatDate(response.submittedAt)}</div>
          <div class="meta">${escapeHtml(response.surveyTitle)}</div>
        </div>
        <span class="badge ${status}">${STATUS_LABELS[status]}</span>
      </div>
      <div class="answer-list">
        ${displayAnswers
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
          </select>
        </label>
        <label>
          管理メモ
          <textarea data-response-memo>${escapeHtml(response.adminMemo || "")}</textarea>
        </label>
        <button class="secondary-button" type="button" data-save-response="${response.id}">保存</button>
        <button class="secondary-button danger-button" type="button" data-delete-response="${response.id}">削除</button>
      </div>
    </article>
  `;
}

async function saveResponseManagement(responseId) {
  const card = document.querySelector(`[data-response-card="${responseId}"]`);
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
  if (!confirm("この回答を削除しますか？")) return;
  try {
    await api.request(`/api/admin/responses/${encodeURIComponent(responseId)}`, {
      method: "DELETE",
      token: state.token,
    });
    state.responses = state.responses.filter((response) => response.id !== responseId);
    showToast("回答を削除しました。");
    renderAll();
  } catch (error) {
    showToast(error.message || "削除できませんでした。");
  }
}

function groupByCustomer() {
  return groupByCustomerFrom(state.responses);
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

function renderSettings() {
  const credentialInfo = document.querySelector("#credentialInfo");
  const storageInfo = document.querySelector("#storageInfo");
  if (!state.adminInfo) {
    credentialInfo.textContent = "認証情報を読み込み中です。";
    storageInfo.innerHTML = `<div class="empty">保存先情報を読み込み中です。</div>`;
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
      <strong>実行アカウント</strong><br />
      ${escapeHtml(state.adminInfo.ownerEmail || "未取得")}
    </article>
  `;
}

function formatSurveyStatusLabel(status) {
  return normalizeSurveyStatus(status) === "draft" ? "下書き" : "公開中";
}

function questionSupportsOptions(type) {
  return type === "choice" || type === "checkbox";
}

function makeBlankQuestion() {
  return {
    id: "",
    label: "",
    type: "text",
    required: true,
    options: [],
  };
}

function makeBlankSurveyDraft() {
  return {
    id: "",
    title: "",
    description: "",
    status: "published",
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
      ? selected.questions
      : [makeBlankQuestion()],
  };
}

function renderSurveyQuestionEditor(question, index) {
  const type = question?.type || "text";
  const supportsOptions = questionSupportsOptions(type);
  const options = supportsOptions
    ? Array.isArray(question?.options) && question.options.length
      ? question.options
      : ["", ""]
    : [];

  return `
    <article class="survey-question-card" data-question-block data-question-id="${escapeHtml(question?.id || "")}">
      <div class="survey-question-head">
        <strong data-question-number>質問 ${index + 1}</strong>
        <button class="secondary-button" type="button" data-remove-question>削除</button>
      </div>
      <div class="survey-question-grid">
        <label>
          質問文
          <input type="text" data-question-label value="${escapeHtml(question?.label || "")}" required />
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
        <input type="checkbox" data-question-required ${question?.required === false ? "" : "checked"} />
        必須回答
      </label>
      <div data-options-wrapper ${supportsOptions ? "" : "hidden"}>
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
  container.insertAdjacentHTML(
    "beforeend",
    renderSurveyQuestionEditor(question || makeBlankQuestion(), container.children.length),
  );
  syncSurveyQuestionNumbers(container);
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

function buildSurveyPayload(form) {
  const title = String(form.elements.title.value || "").trim();
  const description = String(form.elements.description.value || "").trim();
  const status = normalizeSurveyStatus(form.elements.status.value);
  const questionBlocks = Array.from(form.querySelectorAll("[data-question-block]"));

  const questions = questionBlocks.map((block) => {
    const type = String(block.querySelector("[data-question-type]")?.value || "text");
    return {
      id: block.dataset.questionId || "",
      label: String(block.querySelector("[data-question-label]")?.value || "").trim(),
      type,
      required: Boolean(block.querySelector("[data-question-required]")?.checked),
      options: questionSupportsOptions(type)
        ? Array.from(block.querySelectorAll("[data-option-input]"))
            .map((input) => String(input.value || "").trim())
            .filter(Boolean)
        : [],
    };
  });

  return { title, description, status, questions };
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
        .map(
          (survey) => `
            <button
              class="survey-manager-card selectable-card ${survey.id === state.selectedSurveyEditorId ? "active" : ""}"
              type="button"
              data-edit-survey-id="${escapeHtml(survey.id)}"
            >
              <strong>${escapeHtml(survey.title)}</strong>
              <span class="badge ${normalizeSurveyStatus(survey.status)}">${formatSurveyStatusLabel(survey.status)}</span>
              <div>${survey.questions.length}問</div>
              <div class="meta">更新: ${survey.updatedAt ? formatDate(survey.updatedAt) : "-"}</div>
            </button>
          `,
        )
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
          ? `<button class="secondary-button danger-button" type="button" data-delete-survey="${escapeHtml(selectedSurvey.id)}">削除</button>`
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
        公開状態
        <select name="status">
          <option value="published" ${normalizeSurveyStatus(draft.status) === "published" ? "selected" : ""}>公開</option>
          <option value="draft" ${normalizeSurveyStatus(draft.status) === "draft" ? "selected" : ""}>下書き</option>
        </select>
      </label>
      <div class="survey-editor-head">
        <div class="card-title">質問項目</div>
        <button class="secondary-button" id="addSurveyQuestionButton" type="button">質問を追加</button>
      </div>
      <div id="surveyQuestions" class="survey-question-list">
        ${draft.questions.map((question, index) => renderSurveyQuestionEditor(question, index)).join("")}
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
    }
  });

  questionsContainer?.addEventListener("change", (event) => {
    const typeSelect = event.target.closest("[data-question-type]");
    if (!typeSelect) return;
    const block = typeSelect.closest("[data-question-block]");
    const wrapper = block?.querySelector("[data-options-wrapper]");
    if (!block || !wrapper) return;
    if (questionSupportsOptions(typeSelect.value)) {
      wrapper.hidden = false;
      ensureQuestionOptions(block);
      return;
    }
    wrapper.hidden = true;
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
  const rows = [
    ["日時", "お客様", "アンケート", "対応状況", "管理メモ", "回答"],
    ...state.responses.map((response) => [
      formatDate(response.submittedAt),
      response.customerName,
      response.surveyTitle,
      STATUS_LABELS[normalizeStatus(response.status)],
      response.adminMemo || "",
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

function setupInstall() {
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker
        .register("./sw.js?v=20260411-4", { updateViaCache: "none" })
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
  localStorage.removeItem(TOKEN_KEY);
  void api.logout?.();
  setLoggedIn(false);
});

document.querySelector("#refreshButton").addEventListener("click", () => {
  void loadAdminData();
});

document.querySelector("#exportCsvButton").addEventListener("click", exportCsv);

["surveyFilter", "statusFilter", "keywordFilter"].forEach((id) => {
  document.querySelector(`#${id}`).addEventListener("input", renderResponses);
  document.querySelector(`#${id}`).addEventListener("change", renderResponses);
});

document.querySelectorAll("[data-page]").forEach((button) => {
  button.addEventListener("click", () => setPage(button.dataset.page));
});

setupInstall();
setLoggedIn(Boolean(state.token));
if (state.token) {
  loadAdminData().catch((error) => {
    localStorage.removeItem(TOKEN_KEY);
    state.token = "";
    setLoggedIn(false);
    showToast(error.message || "再ログインしてください。");
  });
}
