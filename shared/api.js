window.MayumiSurveyApi = (() => {
  const API_BASE_STORAGE_KEY = "mayumi_bijiris_api_base";
  const GAS_URL_STORAGE_KEY = "mayumi_bijiris_gas_web_app_url";
  const CLIENT_ID_STORAGE_KEY = "mayumi_survey_client_id";
  const DEFAULT_PUBLIC_API_BASE = "https://mayumi-bijiris.onrender.com";

  function safeGetLocal(key) {
    try {
      return localStorage.getItem(key) || "";
    } catch {
      return "";
    }
  }

  function safeSetLocal(key, value) {
    try {
      localStorage.setItem(key, value);
    } catch {
      // Ignore storage errors.
    }
  }

  function cleanBaseUrl(value) {
    return String(value || "").trim().replace(/\/$/, "");
  }

  function now() {
    return new Date().toISOString();
  }

  function normalizeText(value) {
    return String(value ?? "").trim();
  }

  function normalizeEmail(value) {
    return normalizeText(value).toLowerCase();
  }

  function normalizeStatus(status) {
    return ["new", "checked", "done", "trash"].includes(status) ? status : "new";
  }

  function normalizeSurveyStatus(status) {
    return ["published", "draft", "archived"].includes(status) ? status : "published";
  }

  function sleep(ms) {
    return new Promise((resolve) => {
      setTimeout(resolve, ms);
    });
  }

  function makeId(prefix) {
    if (window.crypto?.randomUUID) return `${prefix}_${window.crypto.randomUUID()}`;
    return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2)}`;
  }

  function getClientId() {
    const existing = safeGetLocal(CLIENT_ID_STORAGE_KEY);
    if (existing) return existing;
    const clientId = makeId("client");
    safeSetLocal(CLIENT_ID_STORAGE_KEY, clientId);
    return clientId;
  }

  function isSameOriginApiHost() {
    const host = window.location.hostname;
    return (
      host === "localhost" ||
      host === "127.0.0.1" ||
      host === "0.0.0.0" ||
      host.endsWith(".local") ||
      host.endsWith(".onrender.com") ||
      /^10\./.test(host) ||
      /^172\.(1[6-9]|2\d|3[0-1])\./.test(host) ||
      /^192\.168\./.test(host)
    );
  }

  function resolveExplicitApiBase() {
    const params = new URLSearchParams(window.location.search);
    const queryBase = cleanBaseUrl(params.get("apiBase"));
    if (queryBase) {
      safeSetLocal(API_BASE_STORAGE_KEY, queryBase);
      return queryBase;
    }

    const configuredBase = cleanBaseUrl(window.MAYUMI_BIJIRIS_API_BASE);
    if (configuredBase) return configuredBase;

    return cleanBaseUrl(safeGetLocal(API_BASE_STORAGE_KEY));
  }

  function resolveGasUrl() {
    const params = new URLSearchParams(window.location.search);
    const queryUrl = cleanBaseUrl(params.get("gasUrl"));
    if (queryUrl) {
      safeSetLocal(GAS_URL_STORAGE_KEY, queryUrl);
      return queryUrl;
    }

    const configuredUrl = cleanBaseUrl(window.MAYUMI_GAS_WEB_APP_URL);
    if (configuredUrl) return configuredUrl;

    return cleanBaseUrl(safeGetLocal(GAS_URL_STORAGE_KEY));
  }

  function getDefaultSurveys() {
    const makeSurveys = window.MayumiDefaultSurveys;
    return typeof makeSurveys === "function" ? makeSurveys(now()) : [];
  }

  function getRouteId(path, prefix) {
    if (!path.startsWith(prefix)) return "";
    return decodeURIComponent(path.slice(prefix.length));
  }

  function toAnswerSignatureValue(answer) {
    if (Array.isArray(answer?.files)) {
      return answer.files
        .map((file) => normalizeText(file?.fileId || file?.name || file?.url || file?.dataUrl))
        .join("|");
    }
    if (Array.isArray(answer?.value)) {
      return answer.value.map(normalizeText).filter(Boolean).join(", ");
    }
    return normalizeText(answer?.value);
  }

  function makeAnswerSignature(answers) {
    return JSON.stringify(
      (Array.isArray(answers) ? answers : []).map((answer) => ({
        questionId: normalizeText(answer?.questionId),
        value: toAnswerSignatureValue(answer),
      })),
    );
  }

  function normalizeVisibilityCondition(condition) {
    const questionId = normalizeText(condition?.questionId);
    const value = normalizeText(condition?.value);
    if (!questionId || !value) return null;
    return { questionId, value };
  }

  function normalizeVisibilityConditions(question) {
    const conditions = Array.isArray(question?.visibilityConditions)
      ? question.visibilityConditions
          .map(normalizeVisibilityCondition)
          .filter(Boolean)
      : [];
    if (conditions.length) return conditions;
    const fallback = normalizeVisibilityCondition(question?.visibleWhen);
    return fallback ? [fallback] : [];
  }

  function makeSurveySignature(survey, options = {}) {
    const includeId = options.includeId !== false;
    const includeSortOrder = options.includeSortOrder !== false;
    const signature = {
      title: normalizeText(survey?.title),
      description: normalizeText(survey?.description),
      introMessage: normalizeText(survey?.introMessage),
      completionMessage: normalizeText(survey?.completionMessage),
      status: normalizeSurveyStatus(survey?.status),
      acceptingResponses: survey?.acceptingResponses === false ? false : true,
      startAt: normalizeText(survey?.startAt),
      endAt: normalizeText(survey?.endAt),
      questions: (Array.isArray(survey?.questions) ? survey.questions : []).map((question) => ({
        id: normalizeText(question?.id),
        label: normalizeText(question?.label),
        type: normalizeText(question?.type) || "text",
        required: question?.required === false ? false : true,
        options: Array.isArray(question?.options)
          ? question.options.map(normalizeText).filter(Boolean)
          : [],
        visibilityConditions: normalizeVisibilityConditions(question),
      })),
    };
    if (includeId) {
      signature.id = normalizeText(survey?.id);
    }
    if (includeSortOrder) {
      signature.sortOrder = Number.isFinite(Number(survey?.sortOrder)) ? Number(survey.sortOrder) : 0;
    }
    return JSON.stringify(signature);
  }

  function makePreferencesSignature(preferences) {
    return JSON.stringify({
      notificationEnabled: preferences?.notificationEnabled === false ? false : true,
      notificationEmail: normalizeEmail(preferences?.notificationEmail),
      notificationSubject: normalizeText(preferences?.notificationSubject),
      notificationBody: normalizeText(preferences?.notificationBody),
      dataPolicyText: normalizeText(preferences?.dataPolicyText),
      requireConsent: preferences?.requireConsent === false ? false : true,
      consentText: normalizeText(preferences?.consentText),
      autoBackupEnabled: preferences?.autoBackupEnabled === false ? false : true,
      backupHour: Number(preferences?.backupHour || 0),
      retentionDays: Number(preferences?.retentionDays || 0),
      recoveryMemo: normalizeText(preferences?.recoveryMemo),
      twoFactorEnabled: preferences?.twoFactorEnabled === false ? false : true,
    });
  }

  function buildLocalResponse(payload, responseId) {
    const surveys = getDefaultSurveys();
    const survey = surveys.find((item) => item.id === payload.surveyId && item.status === "published");
    if (!survey) throw new Error("回答できるアンケートが見つかりません。");

    const customerName = normalizeText(payload.customer?.name);
    if (!customerName) throw new Error("お名前を入力してください。");

    const answerMap = new Map(
      (Array.isArray(payload.answers) ? payload.answers : []).map((answer) => [answer.questionId, answer]),
    );

    const answers = survey.questions.map((question) => {
      const answer = answerMap.get(question.id) || {};
      if (question.type === "photo") {
        const files = Array.isArray(answer.files) ? answer.files : [];
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

      return {
        questionId: question.id,
        label: question.label,
        type: question.type,
        value: values.join(", "),
      };
    });

    return {
      id: responseId,
      surveyId: survey.id,
      surveyTitle: survey.title,
      customerClientId: getClientId(),
      customerName,
      customerEmail: "",
      answers,
      status: "new",
      adminMemo: "",
      submittedAt: now(),
    };
  }

  function jsonp(gasUrl, action, params = {}) {
    return new Promise((resolve, reject) => {
      const callbackName = `mayumiGasCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
      const timeout = setTimeout(() => {
        cleanup();
        reject(new Error("Google Apps Script から応答がありません。"));
      }, 20000);

      function cleanup() {
        clearTimeout(timeout);
        delete window[callbackName];
        script.remove();
      }

      window[callbackName] = (data) => {
        cleanup();
        if (data && data.error) {
          reject(new Error(data.error));
          return;
        }
        resolve(data);
      };

      const url = new URL(gasUrl);
      url.searchParams.set("action", action);
      url.searchParams.set("callback", callbackName);
      url.searchParams.set("_", String(Date.now()));
      Object.entries(params).forEach(([key, value]) => {
        if (value !== undefined && value !== null) {
          url.searchParams.set(key, String(value));
        }
      });

      const script = document.createElement("script");
      script.src = url.toString();
      script.onerror = () => {
        cleanup();
        reject(new Error("Google Apps Script に接続できません。"));
      };
      document.head.append(script);
    });
  }

  async function postToGas(gasUrl, action, payload) {
    await fetch(gasUrl, {
      method: "POST",
      mode: "no-cors",
      headers: {
        "Content-Type": "text/plain;charset=utf-8",
      },
      body: JSON.stringify({ action, ...payload }),
    });
  }

  async function waitForSavedResponse(gasUrl, responseId, customerName = "") {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const data = await jsonp(gasUrl, "history", {
        clientId: getClientId(),
        name: customerName || undefined,
      });
      const saved = (data.responses || []).find((response) => response.id === responseId);
      if (saved) return saved;
    }
    return null;
  }

  async function waitForPublicResponse(gasUrl, customerName, matcher) {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const data = await jsonp(gasUrl, "history", {
        clientId: getClientId(),
        name: normalizeText(customerName),
      });
      const responses = Array.isArray(data.responses) ? data.responses : [];
      const matched = matcher(responses);
      if (matched !== undefined) return matched;
    }
    return null;
  }

  async function waitForAdminUpdatedResponse(gasUrl, token, responseId, expected) {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const data = await jsonp(gasUrl, "adminResponses", { token });
      const saved = (data.responses || []).find((response) => response.id === responseId);
      if (!saved) continue;
      if (
        saved.status === expected.status &&
        normalizeText(saved.adminMemo) === normalizeText(expected.adminMemo) &&
        (!expected.answerSignature ||
          makeAnswerSignature(saved.answers) === expected.answerSignature)
      ) {
        return saved;
      }
    }
    return null;
  }

  async function waitForAdminSurvey(gasUrl, token, matcher) {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const data = await jsonp(gasUrl, "adminSurveys", { token });
      const surveys = Array.isArray(data.surveys) ? data.surveys : [];
      const matched = matcher(surveys);
      if (matched !== undefined) return matched;
    }
    return null;
  }

  async function waitForAdminPreferences(gasUrl, token, expectedSignature) {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const data = await jsonp(gasUrl, "adminPreferences", { token });
      const preferences = data.preferences || {};
      if (makePreferencesSignature(preferences) === expectedSignature) {
        return preferences;
      }
    }
    return null;
  }

  async function waitForAdminMemos(gasUrl, token, customerName, expectedMemo) {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const data = await jsonp(gasUrl, "adminCustomerMemos", { token });
      const memos = data.memos || {};
      if (normalizeText(memos[customerName]) === normalizeText(expectedMemo)) {
        return memos;
      }
    }
    return null;
  }

  async function waitForAdminCustomerUpdate(gasUrl, token, currentName, nextName) {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const data = await jsonp(gasUrl, "adminResponses", { token });
      const responses = Array.isArray(data.responses) ? data.responses : [];
      const matched = responses.filter(
        (response) => normalizeText(response.customerName) === normalizeText(nextName || currentName),
      );
      if (matched.length) return matched;
    }
    return null;
  }

  async function waitForAdminCustomerDeletion(gasUrl, token, customerName) {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const data = await jsonp(gasUrl, "adminResponses", { token });
      const responses = Array.isArray(data.responses) ? data.responses : [];
      const remains = responses.some(
        (response) => normalizeText(response.customerName) === normalizeText(customerName),
      );
      if (!remains) return true;
    }
    return false;
  }

  function initGasApi(gasUrl) {
    async function request(path, options = {}) {
      const method = options.method || "GET";

      if (path === "/api/health") {
        return jsonp(gasUrl, "health");
      }

      if (path === "/api/public/surveys") {
        try {
          const remoteData = await jsonp(gasUrl, "surveys");
          const remoteSurveys = remoteData.surveys || [];
          if (remoteSurveys.length) {
            return {
              surveys: remoteSurveys.filter((survey) => survey.status === "published"),
              dataPolicyText: remoteData.dataPolicyText || "",
              requireConsent: remoteData.requireConsent === false ? false : true,
              consentText: remoteData.consentText || "",
              version: remoteData.version || "",
            };
          }
        } catch {
          // Fall back to bundled survey definitions when GAS is temporarily unavailable.
        }
        return {
          surveys: getDefaultSurveys().filter((survey) => survey.status === "published"),
          dataPolicyText: "",
          requireConsent: true,
          consentText: "",
          version: "",
        };
      }

      if (path.startsWith("/api/public/responses") && method === "GET") {
        const url = new URL(path, window.location.origin);
        return jsonp(gasUrl, "history", {
          clientId: getClientId(),
          name: normalizeText(url.searchParams.get("name")),
          recoverByName: url.searchParams.get("recoverByName") === "1" ? "1" : "",
        });
      }

      if (path === "/api/public/responses" && method === "POST") {
        const responseId = makeId("response");
        const response = buildLocalResponse(options.body || {}, responseId);
        await postToGas(gasUrl, "submitResponse", {
          responseId,
          clientId: getClientId(),
          payload: options.body || {},
        });
        const savedResponse = await waitForSavedResponse(
          gasUrl,
          responseId,
          normalizeText(options.body?.customer?.name),
        );
        if (!savedResponse) {
          throw new Error(
            "回答保存を確認できませんでした。通信状態を確認して、履歴に表示されない場合は再送信してください。",
          );
        }
        return { response: savedResponse || response };
      }

      if (path.startsWith("/api/public/responses/") && method === "PUT") {
        const responseId = getRouteId(path, "/api/public/responses/");
        const payload = options.body || {};
        await postToGas(gasUrl, "updatePublicResponse", {
          responseId,
          clientId: getClientId(),
          payload,
        });
        const updated = await waitForPublicResponse(
          gasUrl,
          payload.customer?.name,
          (responses) => responses.find((response) => response.id === responseId),
        );
        if (!updated) {
          throw new Error("回答更新を確認できませんでした。通信状態を確認して、もう一度お試しください。");
        }
        return { response: updated };
      }

      if (path === "/api/public/errors" && method === "POST") {
        await postToGas(gasUrl, "logClientError", {
          payload: options.body || {},
        });
        return { ok: true };
      }

      if (path === "/api/admin/login" && method === "POST") {
        return jsonp(gasUrl, "adminLogin", {
          loginId: options.body?.loginId,
          password: options.body?.password,
        });
      }

      if (path === "/api/admin/login/verify" && method === "POST") {
        return jsonp(gasUrl, "adminVerifyOtp", {
          sessionId: options.body?.sessionId,
          code: options.body?.code,
        });
      }

      if (path === "/api/admin/info" && method === "GET") {
        return jsonp(gasUrl, "adminInfo", { token: options.token });
      }

      if (path === "/api/admin/users" && method === "GET") {
        return jsonp(gasUrl, "adminUsers", { token: options.token });
      }

      if (path === "/api/admin/surveys" && method === "GET") {
        return jsonp(gasUrl, "adminSurveys", { token: options.token });
      }

      if (path === "/api/admin/surveys" && method === "POST") {
        const payload = options.body || {};
        const expectedSignature = makeSurveySignature(payload, {
          includeId: false,
          includeSortOrder: false,
        });
        await postToGas(gasUrl, "adminCreateSurvey", {
          token: options.token,
          payload,
        });
        const survey = await waitForAdminSurvey(gasUrl, options.token, (surveys) =>
          surveys.find(
            (item) =>
              makeSurveySignature(item, {
                includeId: false,
                includeSortOrder: false,
              }) === expectedSignature,
          ),
        );
        if (!survey) throw new Error("アンケートの作成を確認できませんでした。");
        return { survey };
      }

      if (path === "/api/admin/surveys/replace" && method === "POST") {
        const surveys = Array.isArray(options.body?.surveys) ? options.body.surveys : [];
        const expectedSignature = JSON.stringify(surveys.map(makeSurveySignature));
        await postToGas(gasUrl, "adminReplaceSurveys", {
          token: options.token,
          payload: { surveys },
        });
        const restored = await waitForAdminSurvey(gasUrl, options.token, (items) => {
          if (JSON.stringify(items.map(makeSurveySignature)) === expectedSignature) {
            return items;
          }
          return undefined;
        });
        if (!restored) throw new Error("アンケート復元を確認できませんでした。");
        return { surveys: restored };
      }

      if (path.startsWith("/api/admin/surveys/") && method === "PUT") {
        const surveyId = getRouteId(path, "/api/admin/surveys/");
        const payload = options.body || {};
        const expectedSignature = makeSurveySignature(
          { ...payload, id: surveyId },
          { includeSortOrder: false },
        );
        await postToGas(gasUrl, "adminUpdateSurvey", {
          token: options.token,
          surveyId,
          payload,
        });
        const survey = await waitForAdminSurvey(gasUrl, options.token, (surveys) => {
          const updated = surveys.find((item) => item.id === surveyId);
          if (!updated) return undefined;
          return makeSurveySignature(updated, { includeSortOrder: false }) === expectedSignature
            ? updated
            : undefined;
        });
        if (!survey) throw new Error("アンケートの更新を確認できませんでした。");
        return { survey };
      }

      if (path.startsWith("/api/admin/surveys/") && method === "DELETE") {
        const surveyId = getRouteId(path, "/api/admin/surveys/");
        await postToGas(gasUrl, "adminDeleteSurvey", {
          token: options.token,
          surveyId,
        });
        const deleted = await waitForAdminSurvey(gasUrl, options.token, (surveys) =>
          surveys.some((item) => item.id === surveyId) ? undefined : true,
        );
        if (!deleted) throw new Error("アンケートの削除を確認できませんでした。");
        return { ok: true };
      }

      if (path === "/api/admin/preferences" && method === "GET") {
        return jsonp(gasUrl, "adminPreferences", { token: options.token });
      }

      if (path === "/api/admin/preferences" && method === "PUT") {
        const payload = options.body || {};
        const expectedSignature = makePreferencesSignature(payload);
        await postToGas(gasUrl, "adminUpdatePreferences", {
          token: options.token,
          payload,
        });
        const preferences = await waitForAdminPreferences(gasUrl, options.token, expectedSignature);
        if (!preferences) throw new Error("設定更新を確認できませんでした。");
        return { preferences };
      }

      if (path === "/api/admin/users" && method === "PUT") {
        const users = Array.isArray(options.body?.adminUsers) ? options.body.adminUsers : [];
        await postToGas(gasUrl, "adminUpdateUsers", {
          token: options.token,
          payload: { adminUsers: users },
        });
        const info = await jsonp(gasUrl, "adminInfo", { token: options.token });
        return { adminUsers: info.adminUsers || [] };
      }

      if (path === "/api/admin/maintenance/run" && method === "POST") {
        await postToGas(gasUrl, "adminRunMaintenance", {
          token: options.token,
        });
        return { ok: true };
      }

      if (path === "/api/admin/logs" && method === "GET") {
        return jsonp(gasUrl, "adminLogs", { token: options.token });
      }

      if (path === "/api/admin/customer-memos" && method === "GET") {
        return jsonp(gasUrl, "adminCustomerMemos", { token: options.token });
      }

      if (path.startsWith("/api/admin/customers/") && method === "PUT") {
        const currentName = getRouteId(path, "/api/admin/customers/");
        const payload = {
          name: normalizeText(options.body?.name),
          ticketPlan: normalizeText(options.body?.ticketPlan),
          ticketSheet: normalizeText(options.body?.ticketSheet),
          ticketRound: normalizeText(options.body?.ticketRound),
        };
        await postToGas(gasUrl, "adminUpdateCustomer", {
          token: options.token,
          customerName: currentName,
          payload,
        });
        const updatedName = payload.name || currentName;
        const updated = await waitForAdminCustomerUpdate(gasUrl, options.token, currentName, updatedName);
        if (!updated) throw new Error("顧客情報の更新を確認できませんでした。");
        return { customerName: updatedName };
      }

      if (path.startsWith("/api/admin/customers/") && method === "DELETE") {
        const customerName = getRouteId(path, "/api/admin/customers/");
        await postToGas(gasUrl, "adminDeleteCustomer", {
          token: options.token,
          customerName,
        });
        const deleted = await waitForAdminCustomerDeletion(gasUrl, options.token, customerName);
        if (!deleted) throw new Error("顧客削除を確認できませんでした。");
        return { ok: true };
      }

      if (path.startsWith("/api/admin/customer-memos/") && method === "PUT") {
        const customerName = getRouteId(path, "/api/admin/customer-memos/");
        const memo = normalizeText(options.body?.memo);
        await postToGas(gasUrl, "adminUpdateCustomerMemo", {
          token: options.token,
          customerName,
          memo,
        });
        const memos = await waitForAdminMemos(gasUrl, options.token, customerName, memo);
        if (!memos) throw new Error("お客様メモの更新を確認できませんでした。");
        return { memos };
      }

      if (path === "/api/admin/responses" && method === "GET") {
        return jsonp(gasUrl, "adminResponses", { token: options.token });
      }

      if (path.startsWith("/api/admin/responses/") && method === "PUT") {
        const responseId = getRouteId(path, "/api/admin/responses/");
        const payload = {
          status: normalizeStatus(options.body?.status),
          adminMemo: normalizeText(options.body?.adminMemo),
          answers: Array.isArray(options.body?.answers) ? options.body.answers : [],
        };
        await postToGas(gasUrl, "adminUpdateResponse", {
          token: options.token,
          responseId,
          payload,
        });
        const updated = await waitForAdminUpdatedResponse(gasUrl, options.token, responseId, {
          status: payload.status,
          adminMemo: payload.adminMemo,
          answerSignature: payload.answers.length ? makeAnswerSignature(payload.answers) : "",
        });
        if (!updated) throw new Error("回答内容の更新を確認できませんでした。");
        return { response: updated };
      }

      if (path.startsWith("/api/admin/responses/") && method === "DELETE") {
        const responseId = getRouteId(path, "/api/admin/responses/");
        return jsonp(gasUrl, "adminDelete", {
          token: options.token,
          responseId,
        });
      }

      if (path === "/api/admin/export" && method === "GET") {
        return jsonp(gasUrl, "adminExport", { token: options.token });
      }

      if (path === "/api/admin/credentials" && method === "PUT") {
        return jsonp(gasUrl, "adminUpdateCredentials", {
          token: options.token,
          loginId: normalizeText(options.body?.loginId),
          password: normalizeText(options.body?.password),
        });
      }

      throw new Error("API が見つかりません。");
    }

    return { request, mode: "gas" };
  }

  const gasUrl = resolveGasUrl();
  const explicitApiBase = resolveExplicitApiBase();
  const gasApi = gasUrl ? initGasApi(gasUrl) : null;
  const usesSameOriginRest = isSameOriginApiHost();
  const isUnconfiguredStatic = !gasApi && !explicitApiBase && !usesSameOriginRest;
  const restApiBase = explicitApiBase || (usesSameOriginRest || gasApi ? "" : DEFAULT_PUBLIC_API_BASE);

  async function request(path, options = {}) {
    if (gasApi) {
      return gasApi.request(path, options);
    }

    if (isUnconfiguredStatic) {
      if (path === "/api/health") {
        return { ok: false, backend: "unconfigured" };
      }
      if (path === "/api/public/surveys") {
        return {
          surveys: getDefaultSurveys().filter((survey) => survey.status === "published"),
        };
      }
      throw new Error("保存先が未設定です。Google Apps Script のURLを設定してください。");
    }

    const headers = {
      Accept: "application/json",
      ...(options.headers || {}),
    };

    if (options.body !== undefined) {
      headers["Content-Type"] = "application/json";
    }

    if (options.token) {
      headers.Authorization = `Bearer ${options.token}`;
    }

    const response = await fetch(`${restApiBase}${path}`, {
      method: options.method || "GET",
      headers,
      body: options.body === undefined ? undefined : JSON.stringify(options.body),
    });
    const contentType = response.headers.get("content-type") || "";
    const data = contentType.includes("application/json") ? await response.json() : {};

    if (!response.ok) {
      throw new Error(data.error || `API error: ${response.status}`);
    }

    return data;
  }

  async function logout() {
    return Promise.resolve();
  }

  async function logError(source, error, detail = {}) {
    const message = typeof error === "string" ? error : error?.message || "不明なエラー";
    try {
      await request("/api/public/errors", {
        method: "POST",
        body: {
          source,
          message,
          detail,
        },
      });
    } catch {
      // Ignore logging failures on the client.
    }
  }

  return {
    request,
    logout,
    logError,
    mode: gasApi ? "gas" : "rest",
    getClientId,
  };
})();
