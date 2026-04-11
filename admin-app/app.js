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
  return typeof makeSurveys === "function"
    ? makeSurveys(new Date().toISOString()).filter((survey) => survey.status === "published")
    : [];
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
  const customerList = document.querySelector("#responseCustomerList");
  const historyList = document.querySelector("#responseHistoryList");
  const detail = document.querySelector("#responseDetail");
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

  customerList.innerHTML = customers.length
    ? customers
        .map(
          (customer) => `
            <button
              class="customer-card selectable-card ${customer.name === state.selectedCustomerName ? "active" : ""}"
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
    : `<div class="empty">条件に一致するお客様はいません。</div>`;

  historyList.innerHTML = state.selectedCustomerName
    ? customerResponses.length
      ? customerResponses
          .map(
            (response) => `
              <button
                class="response-history-card selectable-card ${response.id === state.selectedResponseId ? "active" : ""}"
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
    : `<div class="empty">お客様名を押すと回答履歴を表示します。</div>`;

  const selectedResponse = customerResponses.find(
    (response) => response.id === state.selectedResponseId,
  );
  detail.innerHTML = selectedResponse
    ? renderResponseCard(selectedResponse)
    : `<div class="empty">回答履歴のアンケートタイトルを押すと詳細を表示します。</div>`;

  customerList.querySelectorAll("[data-customer-name]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedCustomerName = button.dataset.customerName;
      state.selectedResponseId = "";
      renderResponses();
    });
  });

  historyList.querySelectorAll("[data-history-response-id]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedResponseId = button.dataset.historyResponseId;
      renderResponses();
    });
  });

  detail.querySelectorAll("[data-save-response]").forEach((button) => {
    button.addEventListener("click", () => {
      void saveResponseManagement(button.dataset.saveResponse);
    });
  });

  detail.querySelectorAll("[data-delete-response]").forEach((button) => {
    button.addEventListener("click", () => {
      void deleteResponse(button.dataset.deleteResponse);
    });
  });
}

function renderResponseCard(response) {
  const status = normalizeStatus(response.status);
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
        ${response.answers
          .map(
            (answer) => `
              <div class="answer-item">
                <strong>${escapeHtml(answer.label)}</strong><br />
                ${renderAnswerValue(answer)}
              </div>
            `,
          )
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
  const status = card.querySelector("[data-response-status]").value;
  const adminMemo = card.querySelector("[data-response-memo]").value;

  try {
    const result = await api.request(`/api/admin/responses/${encodeURIComponent(responseId)}`, {
      method: "PUT",
      token: state.token,
      body: { status, adminMemo },
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
      navigator.serviceWorker.register("./sw.js").catch(() => {});
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
