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

  function normalizeMeasurementValue(value) {
    if (value === null || value === undefined || value === "") return "";
    const normalized = String(value).replace(/[^\d.-]/g, "");
    if (!normalized) return "";
    const parsed = Number(normalized);
    if (!Number.isFinite(parsed)) return "";
    return Math.round(parsed * 10) / 10;
  }

  function normalizeMeasurementTargets(targets) {
    const normalized = {
      waist: normalizeMeasurementValue(targets?.waist),
      hip: normalizeMeasurementValue(targets?.hip),
      thighRight: normalizeMeasurementValue(targets?.thighRight),
      thighLeft: normalizeMeasurementValue(targets?.thighLeft),
    };
    return Object.values(normalized).some((value) => value !== "") ? normalized : null;
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

  function makeMeasurementSignature(measurement) {
    return JSON.stringify({
      id: normalizeText(measurement?.id),
      customerName: normalizeText(measurement?.customerName),
      measuredAt: normalizeText(measurement?.measuredAt),
      waist: normalizeMeasurementValue(measurement?.waist),
      hip: normalizeMeasurementValue(measurement?.hip),
      thighRight: normalizeMeasurementValue(measurement?.thighRight),
      thighLeft: normalizeMeasurementValue(measurement?.thighLeft),
      memo: normalizeText(measurement?.memo),
    });
  }

  function makeBijirisFileSignature(file, kind = "") {
    return JSON.stringify({
      kind: normalizeText(file?.kind || kind),
      name: normalizeText(file?.name),
      type: normalizeText(file?.type || file?.mimeType),
    });
  }

  function makeBijirisPostSignature(post) {
    return JSON.stringify({
      id: normalizeText(post?.id),
      title: normalizeText(post?.title),
      category: normalizeText(post?.category),
      summary: normalizeText(post?.summary),
      body: normalizeText(post?.body),
      status: normalizeText(post?.status) || "draft",
      pinned: post?.pinned === true,
      publishedAt: normalizeText(post?.publishedAt),
      photos: (Array.isArray(post?.photos) ? post.photos : []).map((file) => makeBijirisFileSignature(file, "photo")),
      documents: (Array.isArray(post?.documents) ? post.documents : []).map((file) => makeBijirisFileSignature(file, "pdf")),
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

  async function waitForCustomerProfile(gasUrl, customer, matcher) {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const data = await jsonp(gasUrl, "history", {
        clientId: getClientId(),
        name: normalizeText(customer?.name),
        nameKana: normalizeText(customer?.nameKana),
        recoverByName: customer?.historyMatchMode === "name" ? "1" : "",
      });
      if (!matcher || matcher(data.customerProfile || null, data.responses || [], data)) {
        return data;
      }
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

  async function waitForAdminMeasurement(gasUrl, token, matcher) {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const data = await jsonp(gasUrl, "adminMeasurements", { token });
      const measurements = Array.isArray(data.measurements) ? data.measurements : [];
      const matched = matcher(measurements);
      if (matched !== undefined) return matched;
    }
    return null;
  }

  async function waitForPublicBijirisPosts(gasUrl, matcher) {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const data = await jsonp(gasUrl, "bijirisPosts", {});
      const posts = Array.isArray(data.posts) ? data.posts : [];
      const matched = matcher(posts);
      if (matched !== undefined) return matched;
    }
    return null;
  }

  async function waitForAdminBijirisPosts(gasUrl, token, matcher) {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const data = await jsonp(gasUrl, "adminBijirisPosts", { token });
      const posts = Array.isArray(data.posts) ? data.posts : [];
      const matched = matcher(posts);
      if (matched !== undefined) return matched;
    }
    return null;
  }

  async function waitForAdminCustomerUpdate(gasUrl, token, currentName, nextName, expected = {}) {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const info = await jsonp(gasUrl, "adminInfo", { token });
      const profiles = Array.isArray(info.customerProfiles) ? info.customerProfiles : [];
      const matchedProfile = profiles.find(
        (profile) => normalizeText(profile.name) === normalizeText(nextName || currentName),
      );
      if (matchedProfile) {
        const activeTicketCard = matchedProfile.activeTicketCard || {};
        const memberNumber = normalizeText(expected.memberNumber).toUpperCase();
        const ticketPlan = normalizeText(expected.ticketPlan);
        const ticketSheet = normalizeText(expected.ticketSheet);
        const ticketRound = normalizeText(expected.ticketRound);
        const shouldMatchTargets = Object.prototype.hasOwnProperty.call(expected, "measurementTargets");
        const expectedTargets = normalizeMeasurementTargets(expected.measurementTargets);
        const actualTargets = normalizeMeasurementTargets(matchedProfile.measurementTargets);
        const memberMatches =
          !memberNumber || normalizeText(matchedProfile.memberNumber).toUpperCase() === memberNumber;
        const planMatches = !ticketPlan || normalizeText(activeTicketCard.plan) === ticketPlan;
        const sheetMatches =
          !ticketSheet || `${Math.floor(Number(activeTicketCard.sheetNumber) || 0)}枚目` === ticketSheet;
        const roundMatches =
          !ticketRound || `${Math.floor(Number(activeTicketCard.round) || 0)}回目` === ticketRound;
        const targetMatches =
          !shouldMatchTargets || JSON.stringify(actualTargets || null) === JSON.stringify(expectedTargets || null);
        if (memberMatches && planMatches && sheetMatches && roundMatches && targetMatches) {
          return matchedProfile;
        }
      }

      const data = await jsonp(gasUrl, "adminResponses", { token });
      const responses = Array.isArray(data.responses) ? data.responses : [];
      const matchedResponse = responses.find(
        (response) => normalizeText(response.customerName) === normalizeText(nextName || currentName),
      );
      if (matchedResponse) return matchedResponse;
    }
    return null;
  }

  async function waitForAdminCustomerDeletion(gasUrl, token, customerName) {
    for (let attempt = 0; attempt < 6; attempt += 1) {
      if (attempt) await sleep(1000);
      const info = await jsonp(gasUrl, "adminInfo", { token });
      const profiles = Array.isArray(info.customerProfiles) ? info.customerProfiles : [];
      const profileRemains = profiles.some(
        (profile) => normalizeText(profile.name) === normalizeText(customerName),
      );
      if (!profileRemains) {
        const data = await jsonp(gasUrl, "adminResponses", { token });
        const responses = Array.isArray(data.responses) ? data.responses : [];
        const remains = responses.some(
          (response) => normalizeText(response.customerName) === normalizeText(customerName),
        );
        if (!remains) return true;
      }

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
              pushAppId: remoteData.pushAppId || "",
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
          pushAppId: window.MAYUMI_ONESIGNAL_APP_ID || "",
          version: "",
        };
      }

      if (path === "/api/public/bijiris-posts" && method === "GET") {
        return jsonp(gasUrl, "bijirisPosts");
      }

      if (path.startsWith("/api/public/responses") && method === "GET") {
        const url = new URL(path, window.location.origin);
        return jsonp(gasUrl, "history", {
          clientId: getClientId(),
          name: normalizeText(url.searchParams.get("name")),
          nameKana: normalizeText(url.searchParams.get("nameKana")),
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

      if (path === "/api/public/customer-profile/ticket-card" && method === "POST") {
        const customer = options.body?.customer || {};
        const ticketCard = options.body?.ticketCard || {};
        const expectedPlan = normalizeText(ticketCard.plan);
        const expectedSheetNumber = Math.floor(Number(ticketCard.sheetNumber) || 0);
        const expectedRound = Math.max(0, Math.floor(Number(ticketCard.round) || 0));
        await postToGas(gasUrl, "updatePublicTicketCard", {
          clientId: getClientId(),
          customer,
          ticketCard,
        });
        const updated = await waitForCustomerProfile(gasUrl, customer, (profile) => {
          const activeTicketCard = profile?.activeTicketCard || {};
          return (
            normalizeText(activeTicketCard.plan) === expectedPlan &&
            Math.floor(Number(activeTicketCard.sheetNumber) || 0) === expectedSheetNumber &&
            Math.max(0, Math.floor(Number(activeTicketCard.round) || 0)) === expectedRound
          );
        });
        if (!updated) {
          throw new Error("スタンプカード情報の保存を確認できませんでした。");
        }
        return { customerProfile: updated.customerProfile || null };
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

      if (path === "/api/admin/measurements" && method === "GET") {
        return jsonp(gasUrl, "adminMeasurements", { token: options.token });
      }

      if (path === "/api/admin/bijiris-posts" && method === "GET") {
        return jsonp(gasUrl, "adminBijirisPosts", { token: options.token });
      }

      if (path === "/api/admin/bijiris-posts" && method === "POST") {
        const payload = {
          ...options.body,
          id: normalizeText(options.body?.id) || makeId("bijiris"),
        };
        await postToGas(gasUrl, "adminCreateBijirisPost", {
          token: options.token,
          payload,
        });
        const post = await waitForAdminBijirisPosts(gasUrl, options.token, (posts) => {
          const matched = posts.find((item) => normalizeText(item.id) === payload.id);
          if (!matched) return undefined;
          return makeBijirisPostSignature(matched) === makeBijirisPostSignature(payload) ? matched : undefined;
        });
        if (!post) throw new Error("豆知識の保存を確認できませんでした。");
        return { post };
      }

      if (path === "/api/admin/bijiris-posts/replace" && method === "POST") {
        const posts = Array.isArray(options.body?.posts) ? options.body.posts : [];
        const expectedSignature = JSON.stringify(posts.map(makeBijirisPostSignature).sort());
        await postToGas(gasUrl, "adminReplaceBijirisPosts", {
          token: options.token,
          payload: { posts },
        });
        const restored = await waitForAdminBijirisPosts(gasUrl, options.token, (items) => {
          const actualSignature = JSON.stringify(items.map(makeBijirisPostSignature).sort());
          return actualSignature === expectedSignature ? items : undefined;
        });
        if (!restored) throw new Error("豆知識の復元を確認できませんでした。");
        return { posts: restored };
      }

      if (path.startsWith("/api/admin/customers/") && method === "PUT") {
        const currentName = getRouteId(path, "/api/admin/customers/");
        const payload = {
          name: normalizeText(options.body?.name),
          memberNumber: normalizeText(options.body?.memberNumber),
          ticketPlan: normalizeText(options.body?.ticketPlan),
          ticketSheet: normalizeText(options.body?.ticketSheet),
          ticketRound: normalizeText(options.body?.ticketRound),
        };
        if (Object.prototype.hasOwnProperty.call(options.body || {}, "measurementTargets")) {
          payload.measurementTargets = normalizeMeasurementTargets(options.body?.measurementTargets);
        }
        await postToGas(gasUrl, "adminUpdateCustomer", {
          token: options.token,
          customerName: currentName,
          payload,
        });
        const updatedName = payload.name || currentName;
        const updated = await waitForAdminCustomerUpdate(
          gasUrl,
          options.token,
          currentName,
          updatedName,
          payload,
        );
        if (!updated) throw new Error("顧客情報の更新を確認できませんでした。");
        return { customerName: updatedName };
      }

      if (path.startsWith("/api/admin/customers/") && path.endsWith("/measurements") && method === "POST") {
        const customerName = getRouteId(path, "/api/admin/customers/").replace(/\/measurements$/, "");
        const payload = {
          id: normalizeText(options.body?.id) || makeId("measure"),
          measuredAt: normalizeText(options.body?.measuredAt),
          waist: normalizeMeasurementValue(options.body?.waist),
          hip: normalizeMeasurementValue(options.body?.hip),
          thighRight: normalizeMeasurementValue(options.body?.thighRight),
          thighLeft: normalizeMeasurementValue(options.body?.thighLeft),
          memo: normalizeText(options.body?.memo),
        };
        await postToGas(gasUrl, "adminCreateMeasurement", {
          token: options.token,
          customerName,
          payload,
        });
        const measurement = await waitForAdminMeasurement(gasUrl, options.token, (measurements) =>
          measurements.find((item) => normalizeText(item.id) === payload.id),
        );
        if (!measurement) throw new Error("測定データの保存を確認できませんでした。");
        return { measurement };
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

      if (path === "/api/admin/measurements/replace" && method === "POST") {
        const measurements = Array.isArray(options.body?.measurements) ? options.body.measurements : [];
        const expectedSignature = JSON.stringify(
          measurements
            .map((measurement) =>
              makeMeasurementSignature({
                id: measurement?.id,
                customerName: measurement?.customerName,
                measuredAt: measurement?.measuredAt,
                waist: measurement?.waist,
                hip: measurement?.hip,
                thighRight: measurement?.thighRight,
                thighLeft: measurement?.thighLeft,
                memo: measurement?.memo,
              }),
            )
            .sort(),
        );
        await postToGas(gasUrl, "adminReplaceMeasurements", {
          token: options.token,
          payload: { measurements },
        });
        const restored = await waitForAdminMeasurement(gasUrl, options.token, (items) => {
          const actualSignature = JSON.stringify(items.map(makeMeasurementSignature).sort());
          return actualSignature === expectedSignature ? items : undefined;
        });
        if (!restored) throw new Error("測定データの復元を確認できませんでした。");
        return { measurements: restored };
      }

      if (path === "/api/admin/responses" && method === "GET") {
        return jsonp(gasUrl, "adminResponses", { token: options.token });
      }

      if (path.startsWith("/api/admin/measurements/") && method === "PUT") {
        const measurementId = getRouteId(path, "/api/admin/measurements/");
        const payload = {
          measuredAt: normalizeText(options.body?.measuredAt),
          waist: normalizeMeasurementValue(options.body?.waist),
          hip: normalizeMeasurementValue(options.body?.hip),
          thighRight: normalizeMeasurementValue(options.body?.thighRight),
          thighLeft: normalizeMeasurementValue(options.body?.thighLeft),
          memo: normalizeText(options.body?.memo),
        };
        await postToGas(gasUrl, "adminUpdateMeasurement", {
          token: options.token,
          measurementId,
          payload,
        });
        const measurement = await waitForAdminMeasurement(gasUrl, options.token, (measurements) => {
          const matched = measurements.find((item) => normalizeText(item.id) === measurementId);
          if (!matched) return undefined;
          const sameValues =
            normalizeText(matched.measuredAt) === payload.measuredAt &&
            normalizeMeasurementValue(matched.waist) === payload.waist &&
            normalizeMeasurementValue(matched.hip) === payload.hip &&
            normalizeMeasurementValue(matched.thighRight) === payload.thighRight &&
            normalizeMeasurementValue(matched.thighLeft) === payload.thighLeft &&
            normalizeText(matched.memo) === payload.memo;
          return sameValues ? matched : undefined;
        });
        if (!measurement) throw new Error("測定データの更新を確認できませんでした。");
        return { measurement };
      }

      if (path.startsWith("/api/admin/bijiris-posts/") && method === "PUT") {
        const postId = getRouteId(path, "/api/admin/bijiris-posts/");
        const payload = {
          ...options.body,
          id: postId,
        };
        await postToGas(gasUrl, "adminUpdateBijirisPost", {
          token: options.token,
          postId,
          payload,
        });
        const post = await waitForAdminBijirisPosts(gasUrl, options.token, (posts) => {
          const matched = posts.find((item) => normalizeText(item.id) === postId);
          if (!matched) return undefined;
          return makeBijirisPostSignature(matched) === makeBijirisPostSignature(payload) ? matched : undefined;
        });
        if (!post) throw new Error("豆知識の更新を確認できませんでした。");
        return { post };
      }

      if (path.startsWith("/api/admin/measurements/") && method === "DELETE") {
        const measurementId = getRouteId(path, "/api/admin/measurements/");
        await postToGas(gasUrl, "adminDeleteMeasurement", {
          token: options.token,
          measurementId,
        });
        const deleted = await waitForAdminMeasurement(gasUrl, options.token, (measurements) =>
          measurements.some((item) => normalizeText(item.id) === measurementId) ? undefined : true,
        );
        if (!deleted) throw new Error("測定データの削除を確認できませんでした。");
        return { ok: true };
      }

      if (path.startsWith("/api/admin/bijiris-posts/") && method === "DELETE") {
        const postId = getRouteId(path, "/api/admin/bijiris-posts/");
        await postToGas(gasUrl, "adminDeleteBijirisPost", {
          token: options.token,
          postId,
        });
        const deleted = await waitForAdminBijirisPosts(gasUrl, options.token, (posts) =>
          posts.some((item) => normalizeText(item.id) === postId) ? undefined : true,
        );
        if (!deleted) throw new Error("豆知識の削除を確認できませんでした。");
        return { ok: true };
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
      if (path === "/api/public/bijiris-posts") {
        return {
          posts: [],
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
