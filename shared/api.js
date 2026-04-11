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
      // Ignore storage errors in private browsing modes.
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
    return ["new", "checked", "done"].includes(status) ? status : "new";
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

  function buildLocalResponse(payload, responseId) {
    const surveys = getDefaultSurveys();
    const survey = surveys.find(
      (item) => item.id === payload.surveyId && item.status === "published",
    );
    if (!survey) throw new Error("回答できるアンケートが見つかりません。");

    const customerName = normalizeText(payload.customer?.name);
    const customerEmail = normalizeEmail(payload.customer?.email);
    if (!customerName) throw new Error("お名前を入力してください。");
    if (!customerEmail || !customerEmail.includes("@")) {
      throw new Error("メールアドレスを入力してください。");
    }

    const answerMap = new Map(
      (Array.isArray(payload.answers) ? payload.answers : []).map((answer) => [
        answer.questionId,
        answer,
      ]),
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
      customerEmail,
      answers,
      status: "new",
      adminMemo: "",
      submittedAt: now(),
    };
  }

  function jsonp(gasUrl, action, params = {}) {
    return new Promise((resolve, reject) => {
      const callbackName = `mayumiGasCallback_${Date.now()}_${Math.random()
        .toString(36)
        .slice(2)}`;
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

  async function waitForSavedResponse(gasUrl, responseId) {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const data = await jsonp(gasUrl, "history", { clientId: getClientId() });
      const saved = (data.responses || []).find((response) => response.id === responseId);
      if (saved) return saved;
    }
    return null;
  }

  function initGasApi(gasUrl) {
    async function request(path, options = {}) {
      const method = options.method || "GET";

      if (path === "/api/health") {
        return jsonp(gasUrl, "health");
      }

      if (path === "/api/public/surveys") {
        const localSurveys = getDefaultSurveys().filter((survey) => survey.status === "published");
        if (localSurveys.length) {
          return { surveys: localSurveys };
        }
        return {
          surveys: (await jsonp(gasUrl, "surveys")).surveys || [],
        };
      }

      if (path.startsWith("/api/public/responses") && method === "GET") {
        return jsonp(gasUrl, "history", { clientId: getClientId() });
      }

      if (path === "/api/public/responses" && method === "POST") {
        const responseId = makeId("response");
        const response = buildLocalResponse(options.body || {}, responseId);
        await postToGas(gasUrl, "submitResponse", {
          responseId,
          clientId: getClientId(),
          payload: options.body || {},
        });
        const savedResponse = await waitForSavedResponse(gasUrl, responseId);
        if (!savedResponse) {
          throw new Error(
            "回答保存を確認できませんでした。通信状態を確認して、回答履歴に表示されない場合はもう一度送信してください。",
          );
        }
        return { response: savedResponse || response };
      }

      if (path === "/api/admin/login" && method === "POST") {
        return jsonp(gasUrl, "adminLogin", {
          loginId: options.body?.loginId,
          password: options.body?.password,
        });
      }

      if (path === "/api/admin/surveys" && method === "GET") {
        return { surveys: getDefaultSurveys() };
      }

      if (path === "/api/admin/responses" && method === "GET") {
        return jsonp(gasUrl, "adminResponses", { token: options.token });
      }

      if (path.startsWith("/api/admin/responses/") && method === "PUT") {
        const responseId = getRouteId(path, "/api/admin/responses/");
        return jsonp(gasUrl, "adminUpdate", {
          token: options.token,
          responseId,
          status: normalizeStatus(options.body?.status),
          adminMemo: normalizeText(options.body?.adminMemo),
        });
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

      throw new Error("API が見つかりません。");
    }

    return { request, mode: "gas" };
  }

  const gasUrl = resolveGasUrl();
  const explicitApiBase = resolveExplicitApiBase();
  const gasApi = gasUrl ? initGasApi(gasUrl) : null;
  const usesSameOriginRest = isSameOriginApiHost();
  const isUnconfiguredStatic = !gasApi && !explicitApiBase && !usesSameOriginRest;
  const restApiBase =
    explicitApiBase || (usesSameOriginRest || gasApi ? "" : DEFAULT_PUBLIC_API_BASE);

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

  return {
    request,
    logout,
    mode: gasApi ? "gas" : "rest",
  };
})();
