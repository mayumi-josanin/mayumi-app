const TOKEN_KEY = "mayumi_survey_admin_token";
const CACHE_PREFIX = "mayumi-admin-survey-";
const ACTIVE_CACHE_NAME = "mayumi-admin-survey-v74";
const AUTO_CACHE_MAINTENANCE_INTERVAL_MS = 6 * 60 * 60 * 1000;
const AUTO_CACHE_MAINTENANCE_KEY = "mayumi_admin_cache_maintenance_at";
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
const SESSION_TYPE_QUESTION_ID = "q_bijiris_session_type";
const MEASUREMENT_MONITOR_PHOTO_QUESTION_IDS = [
  "q_bijiris_session_monitor_photos_6",
  "q_bijiris_session_monitor_photos_10",
  "q_bijiris_session_monitor_photos",
];
const MEASUREMENT_CURRENT_PHOTO_QUESTION_IDS = [
  "q_bijiris_session_ticket_end_photos_6",
  "q_bijiris_session_ticket_end_photos_10",
  "q_bijiris_session_ticket_end_photos",
  "q_ticket_end_photo_last",
];
const MEASUREMENT_PHOTO_QUESTION_IDS = [
  ...MEASUREMENT_MONITOR_PHOTO_QUESTION_IDS,
  ...MEASUREMENT_CURRENT_PHOTO_QUESTION_IDS,
  "q_bijiris_session_ticket_photos",
];
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
const BIJIRIS_CATEGORY_MODE_OPTIONS = [
  { value: "general", label: "通常カテゴリ" },
  { value: "concern", label: "お悩みカテゴリ" },
  { value: "custom", label: "自由入力" },
];
const BIJIRIS_CONCERN_CATEGORY_ROOT_LABEL = "お悩み";
const BIJIRIS_CONCERN_CATEGORY_TREE = [
  {
    id: "female",
    label: "女性",
    groups: [
      {
        id: "female-pelvic-floor",
        label: "産後女性を含む骨盤底筋まわりのお悩み",
        concerns: [
          { id: "female-01", label: "産後の骨盤底筋のゆるみ感が気になる" },
          { id: "female-02", label: "くしゃみ、咳、大笑いでヒヤッとする" },
          { id: "female-03", label: "骨盤まわりを土台からケアしたい" },
        ],
      },
      {
        id: "female-toilet",
        label: "トイレまわりのお悩み",
        concerns: [
          { id: "female-04", label: "尿漏れが気になる" },
          { id: "female-05", label: "頻尿が気になる" },
          { id: "female-06", label: "急な尿意が気になる" },
          { id: "female-07", label: "便秘がち" },
          { id: "female-08", label: "お通じのリズムが気になる" },
        ],
      },
      {
        id: "female-body-shape",
        label: "体型・見た目のお悩み",
        concerns: [
          { id: "female-09", label: "産後のぽっこりお腹が気になる" },
          { id: "female-10", label: "年齢とともに体型の変化が気になる" },
          { id: "female-11", label: "下半身太りが気になる" },
          { id: "female-12", label: "ヒップの下垂が気になる" },
        ],
      },
      {
        id: "female-posture",
        label: "姿勢・日常動作のお悩み",
        concerns: [
          { id: "female-13", label: "産後に姿勢が崩れやすくなった" },
          { id: "female-14", label: "抱っこや家事で下腹や骨盤まわりが気になる" },
          { id: "female-15", label: "姿勢を整えたい" },
        ],
      },
      {
        id: "female-delicate",
        label: "デリケートゾーンまわりのお悩み",
        concerns: [
          { id: "female-16", label: "デリケートゾーンのケアを意識したい" },
          { id: "female-17", label: "膣トレを始めてみたい" },
        ],
      },
      {
        id: "female-circulation",
        label: "冷え・巡りのお悩み",
        concerns: [
          { id: "female-18", label: "冷えやすさが気になる" },
        ],
      },
    ],
  },
  {
    id: "male",
    label: "男性",
    groups: [
      {
        id: "male-toilet",
        label: "トイレまわりのお悩み",
        concerns: [
          { id: "male-01", label: "頻尿が気になる" },
          { id: "male-02", label: "ちょい漏れが気になる" },
          { id: "male-03", label: "急な尿意で不安がある" },
          { id: "male-04", label: "トイレ悩みをケアしたい" },
        ],
      },
      {
        id: "male-delicate",
        label: "デリケートなお悩み",
        concerns: [
          { id: "male-05", label: "EDケアを意識したい" },
          { id: "male-06", label: "デリケートなお悩みを人知れずケアしたい" },
        ],
      },
      {
        id: "male-posture",
        label: "姿勢・骨盤まわりのお悩み",
        concerns: [
          { id: "male-07", label: "長時間の座り仕事で骨盤まわりが気になる" },
          { id: "male-08", label: "猫背や前かがみ姿勢が気になる" },
          { id: "male-09", label: "腰まわりの違和感が気になる" },
        ],
      },
      {
        id: "male-lower-body",
        label: "下半身・体型のお悩み",
        concerns: [
          { id: "male-10", label: "下半身の筋力低下が気になる" },
          { id: "male-11", label: "ヒップラインの崩れが気になる" },
          { id: "male-12", label: "むくみや冷えが気になる" },
          { id: "male-13", label: "運動不足が気になる" },
          { id: "male-14", label: "筋トレが続かない" },
        ],
      },
    ],
  },
];

function buildBijirisConcernPathsFromTree(tree) {
  return (Array.isArray(tree) ? tree : []).flatMap((audience) =>
    (Array.isArray(audience?.groups) ? audience.groups : []).flatMap((group) =>
      (Array.isArray(group?.concerns) ? group.concerns : []).map((concern) =>
        [audience?.label, group?.label, concern?.label]
          .map((value) => normalizeLabel(value))
          .filter(Boolean)
          .join(" > "),
      ),
    ),
  );
}

function getDefaultBijirisCategoryConfig() {
  return {
    generalCategories: BIJIRIS_POST_CATEGORY_OPTIONS.slice(),
    concernRootLabel: BIJIRIS_CONCERN_CATEGORY_ROOT_LABEL,
    concernPaths: buildBijirisConcernPathsFromTree(BIJIRIS_CONCERN_CATEGORY_TREE),
  };
}

function buildBijirisConcernTreeFromPaths(paths, rootLabel) {
  const audienceMap = new Map();
  const tree = [];
  (Array.isArray(paths) ? paths : []).forEach((line) => {
    let parts = String(line || "")
      .split(/\s*>\s*/)
      .map((value) => normalizeLabel(value))
      .filter(Boolean);
    if (parts[0] === rootLabel) parts = parts.slice(1);
    if (parts.length < 3) return;
    const audienceLabel = parts[0];
    const groupLabel = parts[1];
    const concernLabel = parts.slice(2).join(" > ");
    if (!concernLabel) return;

    let audience = audienceMap.get(audienceLabel);
    if (!audience) {
      audience = {
        id: `audience_${tree.length + 1}`,
        label: audienceLabel,
        groups: [],
        _groupMap: new Map(),
      };
      audienceMap.set(audienceLabel, audience);
      tree.push(audience);
    }

    let group = audience._groupMap.get(groupLabel);
    if (!group) {
      group = {
        id: `group_${audience.groups.length + 1}`,
        label: groupLabel,
        concerns: [],
        _concernSet: new Set(),
      };
      audience._groupMap.set(groupLabel, group);
      audience.groups.push(group);
    }

    if (group._concernSet.has(concernLabel)) return;
    group._concernSet.add(concernLabel);
    group.concerns.push({
      id: `concern_${group.concerns.length + 1}`,
      label: concernLabel,
    });
  });

  return tree.map((audience) => ({
    id: audience.id,
    label: audience.label,
    groups: audience.groups.map((group) => ({
      id: group.id,
      label: group.label,
      concerns: group.concerns.map((concern) => ({
        id: concern.id,
        label: concern.label,
      })),
    })),
  }));
}

function normalizeBijirisCategoryConfig(rawConfig) {
  const defaults = getDefaultBijirisCategoryConfig();
  const generalCategories = Array.isArray(rawConfig?.generalCategories)
    ? rawConfig.generalCategories.map((value) => normalizeLabel(value)).filter(Boolean)
    : defaults.generalCategories.slice();
  const concernRootLabel = normalizeLabel(rawConfig?.concernRootLabel) || defaults.concernRootLabel;
  const concernPaths = Array.isArray(rawConfig?.concernPaths)
    ? rawConfig.concernPaths.map((value) => normalizeLabel(value)).filter(Boolean)
    : defaults.concernPaths.slice();
  const normalized = {
    generalCategories: generalCategories.length ? generalCategories : defaults.generalCategories.slice(),
    concernRootLabel,
    concernPaths: concernPaths.length ? concernPaths : defaults.concernPaths.slice(),
  };
  return {
    ...normalized,
    concernTree: buildBijirisConcernTreeFromPaths(normalized.concernPaths, normalized.concernRootLabel),
  };
}

function getBijirisCategoryConfig(preferences = state.preferences) {
  return normalizeBijirisCategoryConfig(preferences?.bijirisCategoryConfig);
}

function renderBijirisGeneralCategorySettingRow(value = "") {
  return `
    <div class="bijiris-settings-input-row" data-bijiris-general-item>
      <input
        type="text"
        value="${escapeHtml(normalizeLabel(value))}"
        placeholder="通常カテゴリ名"
        data-bijiris-general-value
      />
      <button class="secondary-button danger-button" type="button" data-remove-bijiris-general-category>削除</button>
    </div>
  `;
}

function renderBijirisConcernSettingRow(value = "") {
  return `
    <div class="bijiris-settings-input-row" data-bijiris-concern-item>
      <input
        type="text"
        value="${escapeHtml(normalizeLabel(value))}"
        placeholder="個別お悩み名"
        data-bijiris-concern-value
      />
      <button class="secondary-button danger-button" type="button" data-remove-bijiris-concern>削除</button>
    </div>
  `;
}

function renderBijirisConcernGroupSettingCard(group = {}) {
  const concerns = Array.isArray(group?.concerns) && group.concerns.length
    ? group.concerns
    : [{ label: "" }];
  return `
    <article class="history-card bijiris-settings-group-card" data-bijiris-group-item>
      <div class="survey-manager-card-head">
        <strong>大分類</strong>
        <button class="secondary-button danger-button" type="button" data-remove-bijiris-group>大分類を削除</button>
      </div>
      <label>
        大分類名
        <input type="text" value="${escapeHtml(normalizeLabel(group?.label))}" placeholder="大分類名" data-bijiris-group-label />
      </label>
      <div class="stack" data-bijiris-concern-list>
        ${concerns.map((concern) => renderBijirisConcernSettingRow(concern?.label)).join("")}
      </div>
      <button class="secondary-button" type="button" data-add-bijiris-concern>個別お悩みを追加</button>
    </article>
  `;
}

function renderBijirisConcernAudienceSettingCard(audience = {}) {
  const groups = Array.isArray(audience?.groups) && audience.groups.length
    ? audience.groups
    : [{ label: "", concerns: [{ label: "" }] }];
  return `
    <article class="history-card bijiris-settings-audience-card" data-bijiris-audience-item>
      <div class="survey-manager-card-head">
        <strong>対象</strong>
        <button class="secondary-button danger-button" type="button" data-remove-bijiris-audience>対象を削除</button>
      </div>
      <label>
        対象名
        <input type="text" value="${escapeHtml(normalizeLabel(audience?.label))}" placeholder="対象名" data-bijiris-audience-label />
      </label>
      <div class="stack" data-bijiris-group-list>
        ${groups.map((group) => renderBijirisConcernGroupSettingCard(group)).join("")}
      </div>
      <button class="secondary-button" type="button" data-add-bijiris-group>大分類を追加</button>
    </article>
  `;
}

function renderBijirisCategorySettingsEditor(config) {
  const normalizedConfig = normalizeBijirisCategoryConfig(config);
  const audiences = Array.isArray(normalizedConfig.concernTree) && normalizedConfig.concernTree.length
    ? normalizedConfig.concernTree
    : [{ label: "", groups: [{ label: "", concerns: [{ label: "" }] }] }];
  return `
    <article class="history-card bijiris-settings-category-card" data-bijiris-category-settings>
      <div class="survey-editor-head">
        <div>
          <strong>豆知識カテゴリ設定</strong>
          <div class="meta">管理者がカテゴリを追加・削除できます。保存すると豆知識作成時のカテゴリ選択へ反映されます。</div>
        </div>
      </div>
      <section class="stack">
        <div class="card-title">通常カテゴリ</div>
        <div class="stack" data-bijiris-general-list>
          ${normalizedConfig.generalCategories.map((category) => renderBijirisGeneralCategorySettingRow(category)).join("")}
        </div>
        <button class="secondary-button" type="button" data-add-bijiris-general-category>通常カテゴリを追加</button>
      </section>
      <section class="stack">
        <label>
          お悩みカテゴリのルート名
          <input name="bijirisConcernRootLabel" type="text" value="${escapeHtml(normalizedConfig.concernRootLabel)}" />
        </label>
        <div class="card-title">お悩みカテゴリ階層</div>
        <div class="stack" data-bijiris-audience-list>
          ${audiences.map((audience) => renderBijirisConcernAudienceSettingCard(audience)).join("")}
        </div>
        <button class="secondary-button" type="button" data-add-bijiris-audience>対象を追加</button>
      </section>
    </article>
  `;
}

function setupBijirisCategorySettingsEditor(form) {
  if (!form) return;
  form.addEventListener("click", (event) => {
    const generalList = form.querySelector("[data-bijiris-general-list]");
    const audienceList = form.querySelector("[data-bijiris-audience-list]");
    const target = event.target.closest("button");
    if (!target) return;

    if (target.matches("[data-add-bijiris-general-category]")) {
      generalList?.insertAdjacentHTML("beforeend", renderBijirisGeneralCategorySettingRow(""));
      return;
    }

    if (target.matches("[data-remove-bijiris-general-category]")) {
      const items = form.querySelectorAll("[data-bijiris-general-item]");
      if (items.length <= 1) {
        showToast("通常カテゴリは1件以上必要です。");
        return;
      }
      target.closest("[data-bijiris-general-item]")?.remove();
      return;
    }

    if (target.matches("[data-add-bijiris-audience]")) {
      audienceList?.insertAdjacentHTML("beforeend", renderBijirisConcernAudienceSettingCard({ label: "", groups: [{ label: "", concerns: [{ label: "" }] }] }));
      return;
    }

    if (target.matches("[data-remove-bijiris-audience]")) {
      const audienceItem = target.closest("[data-bijiris-audience-item]");
      const totalConcerns = form.querySelectorAll("[data-bijiris-concern-item]").length;
      const audienceConcerns = audienceItem?.querySelectorAll("[data-bijiris-concern-item]").length || 0;
      if (totalConcerns <= audienceConcerns) {
        showToast("お悩みカテゴリは1件以上必要です。");
        return;
      }
      audienceItem?.remove();
      return;
    }

    if (target.matches("[data-add-bijiris-group]")) {
      target.closest("[data-bijiris-audience-item]")?.querySelector("[data-bijiris-group-list]")?.insertAdjacentHTML(
        "beforeend",
        renderBijirisConcernGroupSettingCard({ label: "", concerns: [{ label: "" }] }),
      );
      return;
    }

    if (target.matches("[data-remove-bijiris-group]")) {
      const groupItem = target.closest("[data-bijiris-group-item]");
      const totalConcerns = form.querySelectorAll("[data-bijiris-concern-item]").length;
      const groupConcerns = groupItem?.querySelectorAll("[data-bijiris-concern-item]").length || 0;
      if (totalConcerns <= groupConcerns) {
        showToast("お悩みカテゴリは1件以上必要です。");
        return;
      }
      groupItem?.remove();
      return;
    }

    if (target.matches("[data-add-bijiris-concern]")) {
      target.closest("[data-bijiris-group-item]")?.querySelector("[data-bijiris-concern-list]")?.insertAdjacentHTML(
        "beforeend",
        renderBijirisConcernSettingRow(""),
      );
      return;
    }

    if (target.matches("[data-remove-bijiris-concern]")) {
      const items = form.querySelectorAll("[data-bijiris-concern-item]");
      if (items.length <= 1) {
        showToast("お悩みカテゴリは1件以上必要です。");
        return;
      }
      target.closest("[data-bijiris-concern-item]")?.remove();
    }
  });
}

function collectBijirisCategoryConfigFromForm(form) {
  const generalCategories = Array.from(form.querySelectorAll("[data-bijiris-general-value]"))
    .map((input) => normalizeLabel(input.value))
    .filter(Boolean);
  const concernRootLabel = normalizeLabel(form.elements.bijirisConcernRootLabel?.value);
  const concernPaths = Array.from(form.querySelectorAll("[data-bijiris-audience-item]"))
    .flatMap((audienceItem) => {
      const audienceLabel = normalizeLabel(audienceItem.querySelector("[data-bijiris-audience-label]")?.value);
      return Array.from(audienceItem.querySelectorAll("[data-bijiris-group-item]")).flatMap((groupItem) => {
        const groupLabel = normalizeLabel(groupItem.querySelector("[data-bijiris-group-label]")?.value);
        return Array.from(groupItem.querySelectorAll("[data-bijiris-concern-value]"))
          .map((input) => normalizeLabel(input.value))
          .filter(Boolean)
          .map((concernLabel) => [audienceLabel, groupLabel, concernLabel].filter(Boolean).join(" > "));
      });
    })
    .filter((value) => value.split(/\s*>\s*/).filter(Boolean).length >= 3);
  if (!generalCategories.length) {
    throw new Error("通常カテゴリを1件以上入力してください。");
  }
  if (!concernRootLabel) {
    throw new Error("お悩みカテゴリのルート名を入力してください。");
  }
  if (!concernPaths.length) {
    throw new Error("お悩みカテゴリを1件以上入力してください。");
  }
  return {
    generalCategories,
    concernRootLabel,
    concernPaths,
  };
}
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
  bijirisPendingTaskCount: 0,
  selectedMeasurementPeriod: "6m",
  measurementMetricVisibility: { ...DEFAULT_MEASUREMENT_VISIBILITY },
  installPrompt: null,
};

const loginView = document.querySelector("#loginView");
const adminView = document.querySelector("#adminView");
const toast = document.querySelector("#toast");
const confirmModal = document.querySelector("#confirmModal");
const confirmDialogTitle = document.querySelector("#confirmDialogTitle");
const confirmDialogMessage = document.querySelector("#confirmDialogMessage");
const confirmDialogCancel = document.querySelector("#confirmDialogCancel");
const confirmDialogConfirm = document.querySelector("#confirmDialogConfirm");
const measurementPhotoModal = document.querySelector("#measurementPhotoModal");
const measurementPhotoDialogTitle = document.querySelector("#measurementPhotoDialogTitle");
const measurementPhotoDialogMeta = document.querySelector("#measurementPhotoDialogMeta");
const measurementPhotoDialogBody = document.querySelector("#measurementPhotoDialogBody");
const loginForm = document.querySelector("#loginForm");
const credentialForm = document.querySelector("#credentialForm");
const appUpdateButton = document.querySelector("#appUpdateButton");
const installButton = document.querySelector("#installButton");
const loginSubmitButton = document.querySelector("#loginSubmitButton");
let toastTimer = 0;
let activeConfirmResolver = null;

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

function normalizeLabel(value) {
  return String(value ?? "").trim();
}

function getDefaultBijirisConcernCategoryState(config = getBijirisCategoryConfig()) {
  const audience = config.concernTree[0];
  const group = audience?.groups?.[0];
  const concern = group?.concerns?.[0];
  return {
    concernAudienceId: normalizeLabel(audience?.id),
    concernGroupId: normalizeLabel(group?.id),
    concernId: normalizeLabel(concern?.id),
  };
}

function getBijirisConcernAudienceById(audienceId, config = getBijirisCategoryConfig()) {
  return config.concernTree.find((audience) => audience.id === normalizeLabel(audienceId)) || null;
}

function getBijirisConcernGroupById(audienceId, groupId, config = getBijirisCategoryConfig()) {
  return getBijirisConcernAudienceById(audienceId, config)?.groups.find((group) => group.id === normalizeLabel(groupId)) || null;
}

function getBijirisConcernById(audienceId, groupId, concernId, config = getBijirisCategoryConfig()) {
  return getBijirisConcernGroupById(audienceId, groupId, config)?.concerns.find((concern) => concern.id === normalizeLabel(concernId)) || null;
}

function resolveBijirisConcernCategoryState(rawState = {}, config = getBijirisCategoryConfig()) {
  const fallback = getDefaultBijirisConcernCategoryState(config);
  const audience =
    getBijirisConcernAudienceById(rawState.concernAudienceId, config) || getBijirisConcernAudienceById(fallback.concernAudienceId, config);
  const group =
    audience?.groups.find((item) => item.id === normalizeLabel(rawState.concernGroupId)) ||
    audience?.groups?.[0] ||
    null;
  const concern =
    group?.concerns.find((item) => item.id === normalizeLabel(rawState.concernId)) ||
    group?.concerns?.[0] ||
    null;
  return {
    concernAudienceId: normalizeLabel(audience?.id),
    concernGroupId: normalizeLabel(group?.id),
    concernId: normalizeLabel(concern?.id),
  };
}

function buildBijirisConcernCategoryLabel(rawState = {}, config = getBijirisCategoryConfig()) {
  const resolved = resolveBijirisConcernCategoryState(rawState, config);
  const audience = getBijirisConcernAudienceById(resolved.concernAudienceId, config);
  const group = getBijirisConcernGroupById(resolved.concernAudienceId, resolved.concernGroupId, config);
  const concern = getBijirisConcernById(resolved.concernAudienceId, resolved.concernGroupId, resolved.concernId, config);
  if (!audience || !group || !concern) return "";
  return [
    config.concernRootLabel,
    audience.label,
    group.label,
    concern.label,
  ].join(" > ");
}

function findBijirisConcernCategoryState(category, config = getBijirisCategoryConfig()) {
  const normalizedCategory = normalizeLabel(category);
  if (!normalizedCategory) return null;
  const pathParts = normalizedCategory.split(/\s*>\s*/).map((part) => normalizeLabel(part)).filter(Boolean);
  if (pathParts.length >= 4 && pathParts[0] === config.concernRootLabel) {
    const audience = config.concernTree.find((item) => item.label === pathParts[1]);
    const group = audience?.groups.find((item) => item.label === pathParts[2]);
    const concern = group?.concerns.find((item) => item.label === pathParts[3]);
    if (audience && group && concern) {
      return {
        concernAudienceId: audience.id,
        concernGroupId: group.id,
        concernId: concern.id,
      };
    }
  }
  for (const audience of config.concernTree) {
    for (const group of audience.groups) {
      for (const concern of group.concerns) {
        const fullLabel = buildBijirisConcernCategoryLabel({
          concernAudienceId: audience.id,
          concernGroupId: group.id,
          concernId: concern.id,
        }, config);
        if (normalizedCategory === fullLabel || normalizedCategory === concern.label) {
          return {
            concernAudienceId: audience.id,
            concernGroupId: group.id,
            concernId: concern.id,
          };
        }
      }
    }
  }
  return null;
}

function getBijirisCategoryEditorState(category) {
  const config = getBijirisCategoryConfig();
  const normalizedCategory = normalizeLabel(category);
  const fallbackConcern = getDefaultBijirisConcernCategoryState(config);
  if (!normalizedCategory) {
    return {
      mode: "general",
      generalCategory: config.generalCategories[0],
      customCategory: "",
      ...fallbackConcern,
    };
  }
  if (config.generalCategories.includes(normalizedCategory)) {
    return {
      mode: "general",
      generalCategory: normalizedCategory,
      customCategory: "",
      ...fallbackConcern,
    };
  }
  const concernState = findBijirisConcernCategoryState(normalizedCategory, config);
  if (concernState) {
    return {
      mode: "concern",
      generalCategory: config.generalCategories[0],
      customCategory: "",
      ...resolveBijirisConcernCategoryState(concernState, config),
    };
  }
  return {
    mode: "custom",
    generalCategory: config.generalCategories[0],
    customCategory: normalizedCategory,
    ...fallbackConcern,
  };
}

function resolveBijirisCategoryFromForm(form) {
  const config = getBijirisCategoryConfig();
  const mode = normalizeLabel(form?.elements.categoryMode?.value) || "general";
  if (mode === "concern") {
    return buildBijirisConcernCategoryLabel({
      concernAudienceId: form?.elements.categoryConcernAudience?.value,
      concernGroupId: form?.elements.categoryConcernGroup?.value,
      concernId: form?.elements.categoryConcern?.value,
    }, config);
  }
  if (mode === "custom") {
    return normalizeLabel(form?.elements.categoryCustom?.value);
  }
  const generalCategory = normalizeLabel(form?.elements.categoryGeneral?.value);
  return generalCategory || config.generalCategories[0];
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
    <div class="meta customer-summary-item">
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
  const thumbnailFileId = String(file.thumbnailFileId || "").trim();
  const dataUrl = typeof file.dataUrl === "string" ? file.dataUrl : "";
  const base64Data = typeof file.base64Data === "string" ? file.base64Data : "";
  const thumbnailDataUrl = typeof file.thumbnailDataUrl === "string" ? file.thumbnailDataUrl : "";
  const thumbnailRemoved = file?.thumbnailRemoved === true;
  const title = String(file.title || "").trim();
  if (!name && !fileId && !url && !previewUrl && !downloadUrl && !dataUrl && !base64Data) return null;
  return {
    kind: normalizedKind,
    name: name || (normalizedKind === "pdf" ? "資料" : "写真"),
    title: title || (normalizedKind === "pdf" ? String((name || "資料").replace(/\.pdf$/i, "")).trim() : ""),
    type: String(file.type || file.mimeType || "").trim(),
    fileId,
    url,
    previewUrl,
    downloadUrl,
    thumbnailUrl,
    thumbnailFileId,
    dataUrl,
    base64Data,
    thumbnailDataUrl,
    thumbnailRemoved,
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

function showToast(message, options = {}) {
  if (toastTimer) {
    window.clearTimeout(toastTimer);
    toastTimer = 0;
  }
  toast.textContent = message;
  toast.classList.toggle("busy", options.busy === true);
  toast.classList.add("show");
  if (options.persist === true) return;
  toastTimer = window.setTimeout(() => {
    toast.classList.remove("show", "busy");
    toastTimer = 0;
  }, 2600);
}

function showBusyToast(message) {
  showToast(message, { busy: true, persist: true });
}

function hideToast() {
  if (toastTimer) {
    window.clearTimeout(toastTimer);
    toastTimer = 0;
  }
  toast.classList.remove("show", "busy");
}

function hasPendingBijirisTasks() {
  return state.bijirisPendingTaskCount > 0;
}

function updateBijirisSubmitButtonState() {
  const button = document.querySelector("#bijirisSubmitButton");
  if (!button) return;
  const pending = hasPendingBijirisTasks();
  const idleLabel = button.dataset.idleLabel || button.textContent;
  button.dataset.idleLabel = idleLabel;
  button.disabled = pending;
  button.textContent = pending ? "画像を準備中…" : idleLabel;
}

async function runBijirisPendingTask(task, fallbackMessage) {
  state.bijirisPendingTaskCount += 1;
  updateBijirisSubmitButtonState();
  try {
    return await task;
  } catch (error) {
    showToast(error.message || fallbackMessage);
    return null;
  } finally {
    state.bijirisPendingTaskCount = Math.max(0, state.bijirisPendingTaskCount - 1);
    updateBijirisSubmitButtonState();
  }
}

async function waitForBijirisPendingTasks() {
  if (!hasPendingBijirisTasks()) return;
  showBusyToast("表紙画像を準備中です。完了までお待ちください。");
  while (hasPendingBijirisTasks()) {
    await new Promise((resolve) => window.setTimeout(resolve, 120));
  }
}

function closeConfirmDialog(result = false) {
  if (!confirmModal || confirmModal.hidden) return;
  confirmModal.hidden = true;
  const resolver = activeConfirmResolver;
  activeConfirmResolver = null;
  if (resolver) resolver(result);
}

function openConfirmDialog({
  title = "確認",
  message = "",
  confirmLabel = "実行する",
  cancelLabel = "キャンセル",
  danger = false,
} = {}) {
  if (!confirmModal || !confirmDialogTitle || !confirmDialogMessage || !confirmDialogCancel || !confirmDialogConfirm) {
    return Promise.resolve(window.confirm(message || title));
  }
  if (activeConfirmResolver) {
    activeConfirmResolver(false);
    activeConfirmResolver = null;
  }
  confirmDialogTitle.textContent = title;
  confirmDialogMessage.textContent = message || title;
  confirmDialogCancel.textContent = cancelLabel;
  confirmDialogConfirm.textContent = confirmLabel;
  confirmDialogConfirm.classList.toggle("danger-button", danger);
  confirmModal.hidden = false;
  return new Promise((resolve) => {
    activeConfirmResolver = resolve;
    window.setTimeout(() => confirmDialogConfirm.focus(), 0);
  });
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

function getQuestionRequirementMeta(question) {
  const required = question?.required !== false;
  return {
    label: required ? "必須" : "任意",
    badgeClass: required ? "required" : "optional",
  };
}

function renderQuestionHeading(question, fallbackLabel = "") {
  const requirement = getQuestionRequirementMeta(question);
  return `
    <div class="question-title-row">
      <strong>${escapeHtml(question?.label || fallbackLabel || "質問")}</strong>
      <span class="badge ${requirement.badgeClass}">${requirement.label}</span>
    </div>
  `;
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

function stripPdfExtension(fileName) {
  return String(fileName || "").trim().replace(/\.pdf$/i, "");
}

function getBijirisDocumentDisplayTitle(file, index = 0) {
  return String(file?.title || stripPdfExtension(file?.name) || `資料${index + 1}`).trim();
}

function createPdfThumbnailDataUrl(fileName) {
  const label = getBijirisDocumentDisplayTitle({ title: fileName, name: fileName }).slice(0, 24) || "PDF";
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
  const thumbnailFileId = String(file?.thumbnailFileId || "").trim();
  const fallbackFileId = thumbnailFileId || extractDriveFileId(file?.thumbnailUrl || file?.previewUrl || file?.url || "");
  return buildDriveThumbnailUrl(fallbackFileId) || String(file?.thumbnailUrl || "").trim() || createPdfThumbnailDataUrl(file?.name);
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
    <article class="answer-item customer-summary-card">
      <strong>${escapeHtml(getCustomerNameWithMember(customerName))}</strong>
      <div class="customer-summary-grid">
        <div class="meta customer-summary-item">フリガナ: ${escapeHtml(profile?.nameKana || "-")}</div>
        ${renderCustomerPushStatus(profile?.pushStatus)}
        <div class="meta customer-summary-item">回答数: ${responses.length}件 / アンケート種類: ${surveyCount}件</div>
        <div class="meta customer-summary-item">最新回答: ${latestResponse ? `${escapeHtml(latestResponse.surveyTitle)} / ${formatDate(latestResponse.submittedAt)}` : "-"}</div>
        <div class="meta customer-summary-item customer-summary-item-wide">最新測定: ${latestMeasurement ? `${escapeHtml(formatDateOnly(latestMeasurement.measuredAt))} / WHR ${escapeHtml(formatWhr(latestMeasurement.whr))}` : "-"}</div>
      </div>
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

function formatMeasurementDeltaText(value, unit = "cm") {
  if (value === "" || value === null || value === undefined || !Number.isFinite(Number(value))) {
    return "-";
  }
  const normalized = roundMeasurementDelta(value);
  if (normalized === 0) return `0.0${unit}`;
  return `${normalized > 0 ? "+" : ""}${normalized.toFixed(1)}${unit}`;
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

function buildMeasurementHistoryRows(measurements) {
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
      metrics[metric.key] = {
        value,
        previousDelta:
          value === "" || previousValue === "" ? "" : roundMeasurementDelta(value - previousValue),
        firstDelta: value === "" || firstValue === "" ? "" : roundMeasurementDelta(value - firstValue),
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

function getMeasurementPeriodLabel(period) {
  return MEASUREMENT_PERIOD_OPTIONS.find((option) => option.value === period)?.label || "全期間";
}

function getMeasurementPhotoDateKey(value) {
  if (!value) return "";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return String(value || "").trim().slice(0, 10);
  }
  return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}`;
}

function isMeasurementPhotoAnswer(answer) {
  const questionId = String(answer?.questionId || "").trim();
  const label = String(answer?.label || "").trim();
  const files = getPhotoFilesFromAnswer(answer);
  if (!files.length) return false;
  if (MEASUREMENT_PHOTO_QUESTION_IDS.includes(questionId)) return true;
  return /(モニター時の写真|回数券終了時の写真|計測写真)/.test(label);
}

function isMonitorMeasurementPhotoAnswer(answer) {
  const questionId = String(answer?.questionId || "").trim();
  const label = String(answer?.label || "").trim();
  return MEASUREMENT_MONITOR_PHOTO_QUESTION_IDS.includes(questionId) || label.includes("モニター");
}

function getResponseSessionType(response) {
  const answerMap = new Map((response?.answers || []).map((answer) => [answer.questionId, answer]));
  return getAnswerValueFromQuestionIds(answerMap, [SESSION_TYPE_QUESTION_ID]);
}

function getResponseTicketPlan(response) {
  return getResponseTicketInfo(response).find((item) => item.label === "回数券")?.value || "";
}

function buildMeasurementPhotoStagesFromResponse(response) {
  const survey = findSurveyById(response?.surveyId);
  const sessionType = getResponseSessionType(response);
  return getDisplayAnswers(response, survey)
    .map((answer) => ({
      ...answer,
      files: getPhotoFilesFromAnswer(answer),
    }))
    .filter(isMeasurementPhotoAnswer)
    .map((answer) => {
      const questionId = String(answer.questionId || "").trim();
      const label = String(answer.label || "").trim();
      let kind = "other";
      let buttonLabel = "計測写真";
      if (isMonitorMeasurementPhotoAnswer(answer)) {
        kind = "monitor";
        buttonLabel = "モニター時写真";
      } else if (
        MEASUREMENT_CURRENT_PHOTO_QUESTION_IDS.includes(questionId) ||
        questionId.includes("ticket_end_photos") ||
        label.includes("回数券終了時") ||
        label.includes("計測写真")
      ) {
        kind = "ticket_end";
        buttonLabel = "回数券終了";
      } else if (sessionType === "トライアル") {
        kind = "trial";
        buttonLabel = "トライアル";
      }
      if (kind === "ticket_end" && sessionType === "トライアル") {
        kind = "trial";
        buttonLabel = "トライアル";
      }
      return {
        id: `${response.id}:${questionId || label || buttonLabel}`,
        response,
        answer,
        files: answer.files,
        kind,
        buttonLabel,
      };
    });
}

function getMeasurementPhotoPrimaryResponseCandidate(measurement, responseCandidates) {
  const measuredAtKey = getMeasurementPhotoDateKey(measurement?.measuredAt);
  const measurementTime = new Date(measurement?.measuredAt).getTime();
  return responseCandidates
    .slice()
    .sort((a, b) => {
      const aExact = getMeasurementPhotoDateKey(a.response.submittedAt) === measuredAtKey ? 0 : 1;
      const bExact = getMeasurementPhotoDateKey(b.response.submittedAt) === measuredAtKey ? 0 : 1;
      if (aExact !== bExact) return aExact - bExact;
      const aHasCurrent = a.stages.some((stage) => stage.kind !== "monitor") ? 0 : 1;
      const bHasCurrent = b.stages.some((stage) => stage.kind !== "monitor") ? 0 : 1;
      if (aHasCurrent !== bHasCurrent) return aHasCurrent - bHasCurrent;
      const aDistance = Number.isFinite(measurementTime)
        ? Math.abs(new Date(a.response.submittedAt).getTime() - measurementTime)
        : Number.MAX_SAFE_INTEGER;
      const bDistance = Number.isFinite(measurementTime)
        ? Math.abs(new Date(b.response.submittedAt).getTime() - measurementTime)
        : Number.MAX_SAFE_INTEGER;
      if (aDistance !== bDistance) return aDistance - bDistance;
      return new Date(b.response.submittedAt) - new Date(a.response.submittedAt);
    })[0] || null;
}

function getMeasurementPhotoMonitorStage(responseCandidates, primaryCandidate) {
  const directMonitor = primaryCandidate?.stages.find((stage) => stage.kind === "monitor");
  if (directMonitor) return directMonitor;

  const primaryTime = new Date(primaryCandidate?.response?.submittedAt || 0).getTime();
  const primaryPlan = getResponseTicketPlan(primaryCandidate?.response);
  return responseCandidates
    .flatMap((candidate) => candidate.stages.filter((stage) => stage.kind === "monitor"))
    .sort((a, b) => {
      const aPlan = primaryPlan && getResponseTicketPlan(a.response) === primaryPlan ? 0 : 1;
      const bPlan = primaryPlan && getResponseTicketPlan(b.response) === primaryPlan ? 0 : 1;
      if (aPlan !== bPlan) return aPlan - bPlan;
      const aTime = new Date(a.response.submittedAt).getTime();
      const bTime = new Date(b.response.submittedAt).getTime();
      const aAfter = aTime > primaryTime ? 1 : 0;
      const bAfter = bTime > primaryTime ? 1 : 0;
      if (aAfter !== bAfter) return aAfter - bAfter;
      const aDistance = Math.abs(primaryTime - aTime);
      const bDistance = Math.abs(primaryTime - bTime);
      if (aDistance !== bDistance) return aDistance - bDistance;
      return bTime - aTime;
    })[0] || null;
}

function getMeasurementPhotoCurrentStage(primaryCandidate) {
  if (!primaryCandidate) return null;
  const nonMonitorStages = primaryCandidate.stages.filter((stage) => stage.kind !== "monitor");
  return nonMonitorStages
    .slice()
    .sort((a, b) => {
      const order = { ticket_end: 0, trial: 1, other: 2 };
      const aRank = Object.prototype.hasOwnProperty.call(order, a.kind) ? order[a.kind] : 99;
      const bRank = Object.prototype.hasOwnProperty.call(order, b.kind) ? order[b.kind] : 99;
      return aRank - bRank;
    })[0] || null;
}

function getMeasurementPhotoBundleForMeasurement(measurement) {
  const customerName = String(measurement?.customerName || "").trim();
  if (!customerName) return null;

  const responseCandidates = getActiveCustomerResponses(customerName)
    .map((response) => ({
      response,
      stages: buildMeasurementPhotoStagesFromResponse(response),
    }))
    .filter((candidate) => candidate.stages.length);

  if (!responseCandidates.length) return null;

  const primaryCandidate = getMeasurementPhotoPrimaryResponseCandidate(measurement, responseCandidates);
  const monitorStage = getMeasurementPhotoMonitorStage(responseCandidates, primaryCandidate);
  const currentStage = getMeasurementPhotoCurrentStage(primaryCandidate);
  const stages = [monitorStage, currentStage]
    .filter(Boolean)
    .filter((stage, index, list) => list.findIndex((item) => item.id === stage.id) === index);

  if (!stages.length) return null;

  return {
    measurement,
    stages,
  };
}

function getMeasurementPhotoFileCount(measurement) {
  const bundle = getMeasurementPhotoBundleForMeasurement(measurement);
  if (!bundle?.stages?.length) return 0;
  return bundle.stages.reduce((count, stage) => count + (Array.isArray(stage.files) ? stage.files.length : 0), 0);
}

function renderMeasurementPhotoCell(measurement) {
  const bundle = getMeasurementPhotoBundleForMeasurement(measurement);
  const hasPhotos = bundle?.stages?.length > 0;
  return `
    <div class="measurement-photo-cell">
      <button
        class="secondary-button"
        type="button"
        data-open-measurement-photo="${escapeHtml(measurement.id)}"
        ${hasPhotos ? "" : "disabled"}
      >
        計測写真
      </button>
      <div class="measurement-photo-label">${hasPhotos ? `${bundle.stages.length}種類` : "写真なし"}</div>
    </div>
  `;
}

function renderMeasurementPhotoStageGallery(stage) {
  return `
    <div class="measurement-photo-modal-grid">
      ${stage.files
        .map((file, index) => {
          const preview = getPhotoPreviewSrc(file);
          const title = file.name || `計測写真${index + 1}`;
          return `
            <button
              class="measurement-photo-modal-card lightbox-trigger"
              type="button"
              data-lightbox-src="${escapeHtml(getPhotoLightboxSrc(file))}"
              data-lightbox-title="${escapeHtml(title)}"
            >
              ${preview ? `<img src="${escapeHtml(preview)}" alt="${escapeHtml(title)}" />` : `<div class="empty">画像を表示できません。</div>`}
              <span>${escapeHtml(title)}</span>
            </button>
          `;
        })
        .join("")}
    </div>
  `;
}

function renderMeasurementPhotoModalView(bundle, selectedStageId = "") {
  if (!measurementPhotoDialogMeta || !measurementPhotoDialogBody) return;
  const selectedStage = bundle.stages.find((stage) => stage.id === selectedStageId) || null;
  measurementPhotoDialogMeta.textContent = selectedStage
    ? `${formatDateOnly(bundle.measurement.measuredAt)} / ${selectedStage.buttonLabel} / ${selectedStage.files.length}枚`
    : `${formatDateOnly(bundle.measurement.measuredAt)} / 表示する写真を選択してください。`;
  measurementPhotoDialogBody.innerHTML = `
    <div class="measurement-photo-stage-list">
      ${bundle.stages
        .map(
          (stage) => `
            <button
              class="secondary-button measurement-photo-stage-button ${selectedStage?.id === stage.id ? "is-active" : ""}"
              type="button"
              data-measurement-photo-stage="${escapeHtml(stage.id)}"
            >
              ${escapeHtml(stage.buttonLabel)}
            </button>
          `,
        )
        .join("")}
    </div>
    ${selectedStage ? renderMeasurementPhotoStageGallery(selectedStage) : '<div class="empty">写真ボタンを選択してください。</div>'}
  `;
  measurementPhotoDialogBody.querySelectorAll("[data-measurement-photo-stage]").forEach((button) => {
    button.addEventListener("click", () => {
      renderMeasurementPhotoModalView(bundle, button.dataset.measurementPhotoStage || "");
    });
  });
  attachLightboxHandlers(measurementPhotoDialogBody);
}

function openMeasurementPhotoModal(measurementId) {
  const measurement = state.measurements.find((item) => item.id === measurementId);
  const bundle = measurement ? getMeasurementPhotoBundleForMeasurement(measurement) : null;
  if (!measurement || !measurementPhotoModal || !measurementPhotoDialogTitle || !measurementPhotoDialogMeta || !measurementPhotoDialogBody) {
    return;
  }

  if (!bundle?.stages?.length) {
    showToast("表示できる計測写真はありません。");
    return;
  }

  measurementPhotoDialogTitle.textContent = `${measurement.customerName || "顧客"} / 計測写真`;
  measurementPhotoModal.hidden = false;
  renderMeasurementPhotoModalView(bundle);
}

function closeMeasurementPhotoModal() {
  if (!measurementPhotoModal || !measurementPhotoDialogBody) return;
  measurementPhotoModal.hidden = true;
  measurementPhotoDialogBody.innerHTML = "";
}

function renderMeasurementSummaryCards(measurements) {
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
          </article>
        `;
      }).join("")}
    </div>
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
  const values = [];
  chronological.forEach((measurement) => {
    visibleMetrics.forEach((metric) => {
      const value = normalizeMeasurementValue(measurement[metric.key]);
      if (value !== "") values.push(Number(value));
    });
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
          ${series}
          ${xLabels}
        </svg>
      </div>
    </article>
  `;
}

function renderMeasurementHistoryTable(measurements) {
  const rows = buildMeasurementHistoryRows(measurements);
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
      <div class="stage-head">
        <div>
          <strong>測定履歴</strong>
          <div class="meta">新しい測定が上に表示されます。前回比と初回比を自動計算しています。</div>
        </div>
        <div class="action-row">
          <button class="secondary-button" type="button" data-export-measurements-csv>CSV出力</button>
          <button class="secondary-button" type="button" data-export-measurements-pdf>PDF出力</button>
        </div>
      </div>
      <div class="measurement-table-wrap">
        <table class="measurement-table">
          <thead>
            <tr>
              <th>測定日</th>
              ${MEASUREMENT_METRICS.map((metric) => `<th>${escapeHtml(metric.label)}</th>`).join("")}
              <th>WHR</th>
              <th>左右差</th>
              <th>メモ</th>
              <th>計測写真</th>
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
                      </td>
                    `;
                  }).join("")}
                  <td>${escapeHtml(formatWhr(row.whr))}</td>
                  <td>${row.leftRightGap === "" ? "-" : `${escapeHtml(row.leftRightGap.toFixed(1))}cm`}</td>
                  <td class="measurement-memo-cell">${escapeHtml(row.memo || "-")}</td>
                  <td class="measurement-photo-table-cell">${renderMeasurementPhotoCell(row)}</td>
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
  return `
    <section class="measurement-section">
      ${renderMeasurementSummaryCards(allMeasurements)}
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
      ${renderMeasurementHistoryTable(filteredMeasurements)}
    </section>
  `;
}

function getFilteredCustomerMeasurements(customerName) {
  return filterMeasurementsByPeriod(getCustomerMeasurements(customerName), state.selectedMeasurementPeriod);
}

function exportCustomerMeasurementsCsv(customerName) {
  const rows = buildMeasurementHistoryRows(getFilteredCustomerMeasurements(customerName));
  if (!rows.length) {
    showToast("出力対象の測定履歴がありません。");
    return;
  }
  const header = ["測定日"];
  MEASUREMENT_METRICS.forEach((metric) => {
    header.push(metric.label, `${metric.label} 前回比`, `${metric.label} 初回比`);
  });
  header.push("WHR", "左右差", "メモ", "計測写真枚数");
  const csvRows = [
    header,
    ...rows.map((row) => {
      const line = [formatDateOnly(row.measuredAt)];
      MEASUREMENT_METRICS.forEach((metric) => {
        const metricRow = row.metrics[metric.key];
        line.push(
          formatMeasurementValue(metricRow.value),
          formatMeasurementDeltaText(metricRow.previousDelta),
          formatMeasurementDeltaText(metricRow.firstDelta),
        );
      });
      line.push(
        formatWhr(row.whr),
        row.leftRightGap === "" ? "-" : `${row.leftRightGap.toFixed(1)}cm`,
        row.memo || "",
        String(getMeasurementPhotoFileCount(row)),
      );
      return line;
    }),
  ];
  const csv = csvRows
    .map((row) => row.map((cell) => `"${String(cell).replaceAll('"', '""')}"`).join(","))
    .join("\n");
  const blob = new Blob([`\uFEFF${csv}`], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${sanitizeDownloadNamePart(customerName, "customer")}-measurements-${Date.now()}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function exportCustomerMeasurementsPdf(customerName) {
  const rows = buildMeasurementHistoryRows(getFilteredCustomerMeasurements(customerName));
  if (!rows.length) {
    showToast("出力対象の測定履歴がありません。");
    return;
  }
  const printWindow = window.open("", "_blank", "noopener,noreferrer");
  if (!printWindow) {
    showToast("PDF出力画面を開けませんでした。");
    return;
  }
  const headerCells = [
    "<th>測定日</th>",
    ...MEASUREMENT_METRICS.map((metric) => `<th>${escapeHtml(metric.label)}</th>`),
    "<th>WHR</th>",
    "<th>左右差</th>",
    "<th>メモ</th>",
    "<th>計測写真</th>",
  ].join("");
  const bodyRows = rows
    .map((row) => {
      const metricCells = MEASUREMENT_METRICS.map((metric) => {
        const metricRow = row.metrics[metric.key];
        return `
          <td>
            <div>${escapeHtml(formatMeasurementValue(metricRow.value))}</div>
            <div class="sub">前回 ${escapeHtml(formatMeasurementDeltaText(metricRow.previousDelta))}</div>
            <div class="sub">初回 ${escapeHtml(formatMeasurementDeltaText(metricRow.firstDelta))}</div>
          </td>
        `;
      }).join("");
      return `
        <tr>
          <td>${escapeHtml(formatDateOnly(row.measuredAt))}</td>
          ${metricCells}
          <td>${escapeHtml(formatWhr(row.whr))}</td>
          <td>${row.leftRightGap === "" ? "-" : `${escapeHtml(row.leftRightGap.toFixed(1))}cm`}</td>
          <td>${escapeHtml(row.memo || "-")}</td>
          <td>${escapeHtml(String(getMeasurementPhotoFileCount(row)))}枚</td>
        </tr>
      `;
    })
    .join("");
  printWindow.document.write(`
    <!doctype html>
    <html lang="ja">
      <head>
        <meta charset="UTF-8" />
        <title>測定履歴 PDF</title>
        <style>
          body { font-family: sans-serif; padding: 24px; color: #222; }
          h1 { font-size: 22px; margin-bottom: 8px; }
          .meta { color: #666; margin-bottom: 8px; }
          table { width: 100%; border-collapse: collapse; table-layout: fixed; }
          th, td { border: 1px solid #ddd; padding: 8px; vertical-align: top; font-size: 12px; text-align: left; word-break: break-word; }
          th { background: #f8f5ef; }
          .sub { color: #666; font-size: 11px; margin-top: 2px; }
        </style>
      </head>
      <body>
        <h1>測定履歴</h1>
        <div class="meta">顧客: ${escapeHtml(getCustomerNameWithMember(customerName))}</div>
        <div class="meta">表示期間: ${escapeHtml(getMeasurementPeriodLabel(state.selectedMeasurementPeriod))} / 件数: ${rows.length}件</div>
        <table>
          <thead>
            <tr>${headerCells}</tr>
          </thead>
          <tbody>
            ${bodyRows}
          </tbody>
        </table>
      </body>
    </html>
  `);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
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
  if (!(await openConfirmDialog({
    title: "測定データ削除",
    message: "この測定データを削除しますか？",
    confirmLabel: "削除する",
    danger: true,
  }))) return;
  try {
    showBusyToast("測定データを削除中です。");
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
  if (!(await openConfirmDialog({
    title: "顧客情報削除",
    message: "この顧客情報と関連する回答履歴をすべて削除しますか？",
    confirmLabel: "削除する",
    danger: true,
  }))) return;
  try {
    showBusyToast("顧客情報を削除中です。");
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
          .map((answer) => {
            const question = survey?.questions.find((item) => item.id === answer.questionId) || {
              id: answer.questionId,
              label: answer.label,
              type: answer.type,
              options: [],
            };
            return `
              <div class="answer-item">
                ${renderQuestionHeading(question, answer.label)}
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
            `;
          })
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
                    .map((answer) => {
                      const question = survey?.questions.find((item) => item.id === answer.questionId) || {
                        id: answer.questionId,
                        label: answer.label,
                        type: answer.type,
                        options: [],
                      };
                      return `
                        <div class="answer-item">
                          ${renderQuestionHeading(question, answer.label)}
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
                      `;
                    })
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

  stage.querySelector("[data-export-measurements-csv]")?.addEventListener("click", () => {
    if (selectedCustomer?.name) {
      exportCustomerMeasurementsCsv(selectedCustomer.name);
    }
  });

  stage.querySelector("[data-export-measurements-pdf]")?.addEventListener("click", () => {
    if (selectedCustomer?.name) {
      exportCustomerMeasurementsPdf(selectedCustomer.name);
    }
  });

  stage.querySelectorAll("[data-open-measurement-photo]").forEach((button) => {
    button.addEventListener("click", () => {
      openMeasurementPhotoModal(button.dataset.openMeasurementPhoto || "");
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
          ${renderQuestionHeading(question)}
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
                ${renderQuestionHeading(question, answer.label)}
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
  const hardDelete = normalizeStatus(target.status) === "trash";
  if (!(await openConfirmDialog({
    title: hardDelete ? "回答完全削除" : "回答削除",
    message: hardDelete ? "この回答を完全削除しますか？" : "この回答をゴミ箱へ移動しますか？",
    confirmLabel: hardDelete ? "完全削除する" : "ゴミ箱へ移動",
    danger: true,
  }))) return;
  try {
    showBusyToast(hardDelete ? "回答を完全削除中です。" : "回答をゴミ箱へ移動中です。");
    if (hardDelete) {
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
      bijirisCategoryConfig: getDefaultBijirisCategoryConfig(),
    };
    const bijirisCategoryConfig = normalizeBijirisCategoryConfig(preferences.bijirisCategoryConfig);
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
        ${renderBijirisCategorySettingsEditor(bijirisCategoryConfig)}
        <button class="primary-button" type="submit">設定を保存</button>
      </form>
    `;
    setupBijirisCategorySettingsEditor(preferencesCard.querySelector("form"));
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
  const requirement = getQuestionRequirementMeta(normalizedQuestion);

  return `
    <article class="survey-question-card" data-question-block data-question-id="${escapeHtml(normalizedQuestion.id || "")}">
      <div class="survey-question-head">
        <div class="survey-question-title">
          <strong data-question-number>質問 ${index + 1}</strong>
          <span class="badge ${requirement.badgeClass}" data-question-required-badge>${requirement.label}</span>
        </div>
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

function syncQuestionRequiredBadge(block) {
  if (!block) return;
  const badge = block.querySelector("[data-question-required-badge]");
  const checkbox = block.querySelector("[data-question-required]");
  if (!badge || !checkbox) return;
  const requirement = getQuestionRequirementMeta({ required: checkbox.checked });
  badge.className = `badge ${requirement.badgeClass}`;
  badge.textContent = requirement.label;
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

function getBijirisSubmitButtonLabel(draft, selectedPost, isEditingPublishedPost) {
  if (draft.status === "published") {
    if (selectedPost) {
      return isEditingPublishedPost ? "保存して顧客アプリへ反映" : "公開内容を保存";
    }
    return "作成して公開";
  }
  if (draft.status === "draft") {
    return selectedPost ? "下書きを保存" : "下書きを作成";
  }
  return selectedPost ? "保存" : "作成";
}

function createEmptyBijirisPostDraft() {
  const config = getBijirisCategoryConfig();
  return normalizeBijirisPost({
    id: "",
    title: "",
    category: config.generalCategories[0] || BIJIRIS_POST_CATEGORY_OPTIONS[0],
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

function renderBijirisCategoryEditor(category) {
  const config = getBijirisCategoryConfig();
  const editorState = getBijirisCategoryEditorState(category);
  const concernState = resolveBijirisConcernCategoryState(editorState, config);
  const audience = getBijirisConcernAudienceById(concernState.concernAudienceId, config);
  const groups = Array.isArray(audience?.groups) ? audience.groups : [];
  const concerns = Array.isArray(getBijirisConcernGroupById(concernState.concernAudienceId, concernState.concernGroupId, config)?.concerns)
    ? getBijirisConcernGroupById(concernState.concernAudienceId, concernState.concernGroupId, config).concerns
    : [];
  const resolvedCategory = editorState.mode === "concern"
    ? buildBijirisConcernCategoryLabel(concernState, config)
    : editorState.mode === "custom"
      ? editorState.customCategory
      : editorState.generalCategory;
  return `
    <article class="history-card bijiris-category-card">
      <div class="survey-editor-head">
        <div>
          <div class="card-title">カテゴリ</div>
          <div class="meta">通常カテゴリ / ${escapeHtml(config.concernRootLabel)} / 自由入力から選べます。</div>
        </div>
      </div>
      <div class="bijiris-category-grid">
        <label>
          分類方法
          <select name="categoryMode" data-bijiris-category-control>
            ${BIJIRIS_CATEGORY_MODE_OPTIONS.map((option) => `
              <option value="${escapeHtml(option.value)}" ${editorState.mode === option.value ? "selected" : ""}>${escapeHtml(option.label)}</option>
            `).join("")}
          </select>
        </label>
        ${
          editorState.mode === "general"
            ? `
              <label>
                通常カテゴリ
                <select name="categoryGeneral" data-bijiris-category-control>
                  ${config.generalCategories.map((option) => `
                    <option value="${escapeHtml(option)}" ${editorState.generalCategory === option ? "selected" : ""}>${escapeHtml(option)}</option>
                  `).join("")}
                </select>
              </label>
            `
            : ""
        }
        ${
          editorState.mode === "concern"
            ? `
              <label>
                対象
                <select name="categoryConcernAudience" data-bijiris-category-control>
                  ${config.concernTree.map((option) => `
                    <option value="${escapeHtml(option.id)}" ${concernState.concernAudienceId === option.id ? "selected" : ""}>${escapeHtml(option.label)}</option>
                  `).join("")}
                </select>
              </label>
              <label>
                大分類
                <select name="categoryConcernGroup" data-bijiris-category-control>
                  ${groups.map((option) => `
                    <option value="${escapeHtml(option.id)}" ${concernState.concernGroupId === option.id ? "selected" : ""}>${escapeHtml(option.label)}</option>
                  `).join("")}
                </select>
              </label>
              <label class="bijiris-category-grid-span">
                個別お悩み
                <select name="categoryConcern" data-bijiris-category-control>
                  ${concerns.map((option) => `
                    <option value="${escapeHtml(option.id)}" ${concernState.concernId === option.id ? "selected" : ""}>${escapeHtml(option.label)}</option>
                  `).join("")}
                </select>
              </label>
            `
            : ""
        }
        ${
          editorState.mode === "custom"
            ? `
              <label class="bijiris-category-grid-span">
                カスタムカテゴリ
                <input
                  name="categoryCustom"
                  type="text"
                  value="${escapeHtml(editorState.customCategory)}"
                  placeholder="例: キャンペーン / 期間限定"
                  data-bijiris-category-custom
                />
              </label>
            `
            : ""
        }
      </div>
      <div class="bijiris-category-preview">
        <strong>保存されるカテゴリ</strong>
        <div class="meta">${escapeHtml(resolvedCategory || "未設定")}</div>
      </div>
    </article>
  `;
}

function renderAdminBijirisDocumentPreview(file, index, compact = false) {
  const href = String(file?.downloadUrl || file?.previewUrl || file?.url || "#").trim();
  const title = getBijirisDocumentDisplayTitle(file, index);
  return `
    <a class="history-card bijiris-preview-document ${compact ? "compact" : ""}" href="${escapeHtml(href)}" target="_blank" rel="noreferrer">
      <img class="bijiris-preview-document-thumb" src="${escapeHtml(getBijirisDocumentThumbnailSrc(file))}" alt="${escapeHtml(title)}" />
      <strong>${escapeHtml(title)}</strong>
      <div class="meta">${escapeHtml(file?.name || `document_${index + 1}.pdf`)}</div>
    </a>
  `;
}

function renderAdminBijirisStaticDocumentPreview(file, index, compact = false) {
  const title = getBijirisDocumentDisplayTitle(file, index);
  return `
    <article class="history-card bijiris-preview-document ${compact ? "compact" : ""}">
      <img class="bijiris-preview-document-thumb" src="${escapeHtml(getBijirisDocumentThumbnailSrc(file))}" alt="${escapeHtml(title)}" />
      <div class="bijiris-preview-media-copy">
        <span class="badge">PDF</span>
        <strong>${escapeHtml(title)}</strong>
        <div class="meta">${escapeHtml(file?.name || `document_${index + 1}.pdf`)}</div>
      </div>
    </article>
  `;
}

function renderAdminBijirisStaticPhotoPreview(file, index, compact = false) {
  const preview = getPhotoPreviewSrc(file);
  const title = escapeHtml(file?.name || `写真${index + 1}`);
  return `
    <article class="history-card bijiris-preview-photo ${compact ? "compact" : ""}">
      ${preview ? `<img src="${escapeHtml(preview)}" alt="${title}" />` : ""}
      <div class="bijiris-preview-media-copy">
        <span class="badge draft">写真</span>
        <strong>${title}</strong>
      </div>
    </article>
  `;
}

function renderAdminBijirisListMediaStrip(post, compact = false) {
  const items = [];
  if (post.photos.length) {
    items.push(renderAdminBijirisStaticPhotoPreview(post.photos[0], 0, compact));
  }
  if (post.documents.length) {
    items.push(renderAdminBijirisStaticDocumentPreview(post.documents[0], 0, compact));
  }
  if (items.length < 2 && post.photos.length > 1) {
    items.push(renderAdminBijirisStaticPhotoPreview(post.photos[1], 1, compact));
  }
  if (items.length < 2 && post.documents.length > 1) {
    items.push(renderAdminBijirisStaticDocumentPreview(post.documents[1], 1, compact));
  }
  if (!items.length) return "";
  return `<div class="bijiris-preview-document-grid ${compact ? "compact" : ""}">${items.join("")}</div>`;
}

function renderAdminBijirisPreviewCard(post) {
  const publishedAt = post.publishedAt || post.updatedAt || post.createdAt || new Date().toISOString();
  const preview = post.summary || post.body.slice(0, 90);
  const mediaPreview = renderAdminBijirisListMediaStrip(post, true);
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
      ${mediaPreview || (preview ? `<div class="meta">${escapeHtml(preview)}</div>` : `<div class="meta">写真またはPDFを追加すると一覧に表示されます。</div>`)}
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

function syncBijirisDocumentMetaFromForm(form, draft) {
  if (!form || !draft) return;
  const titleInputs = Array.from(form.querySelectorAll("[data-bijiris-document-title]"));
  titleInputs.forEach((input) => {
    const index = Number(input.dataset.bijirisDocumentTitle);
    if (!Number.isInteger(index) || !draft.documents[index]) return;
    draft.documents[index].title = String(input.value || "").trim() || stripPdfExtension(draft.documents[index].name);
  });
}

function syncBijirisDraftFromForm(form) {
  if (!form) return getBijirisEditorDraft();
  const draft = getBijirisEditorDraft();
  draft.title = String(form.elements.title?.value || "").trim();
  draft.category = resolveBijirisCategoryFromForm(form);
  draft.summary = String(form.elements.summary?.value || "").trim();
  draft.body = String(form.elements.body?.value || "").trim();
  draft.status = ["published", "draft", "archived"].includes(String(form.elements.status?.value || "").trim())
    ? String(form.elements.status.value).trim()
    : "published";
  draft.pinned = Boolean(form.elements.pinned?.checked);
  draft.notifyCustomers = Boolean(form.elements.notifyCustomers?.checked);
  draft.notificationTitle = String(form.elements.notificationTitle?.value || "").trim();
  draft.notificationBody = String(form.elements.notificationBody?.value || "").trim();
  syncBijirisDocumentMetaFromForm(form, draft);
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

async function prepareBijirisDocumentCoverFile(file) {
  if (!String(file?.type || "").startsWith("image/")) {
    throw new Error("表紙画像ファイルを選択してください。");
  }
  const source = await readFileAsDataUrl(file, "表紙画像を読み込めませんでした。");
  const image = await loadImage(source);
  const maxSize = 1200;
  const quality = 0.78;
  const scale = Math.min(1, maxSize / Math.max(image.naturalWidth, image.naturalHeight));
  const width = Math.max(1, Math.round(image.naturalWidth * scale));
  const height = Math.max(1, Math.round(image.naturalHeight * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  context.drawImage(image, 0, 0, width, height);
  return canvas.toDataURL("image/jpeg", quality);
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
    title: stripPdfExtension(file.name || "document"),
    mimeType: "application/pdf",
    type: "application/pdf",
    base64Data: match[1],
    thumbnailDataUrl: "",
  };
}

function renderBijirisAttachmentPreview(files, kind) {
  if (!files.length) {
    return `<div class="empty">${kind === "pdf" ? "PDF はまだありません。" : "写真はまだありません。"}</div>`;
  }
  if (kind === "pdf") {
    return `
      <div class="bijiris-document-editor-list">
        ${files
          .map((file, index) => `
            <article class="history-card bijiris-document-editor-card">
              <div class="bijiris-document-editor-thumb-wrap">
                <img
                  class="bijiris-preview-document-thumb"
                  src="${escapeHtml(getBijirisDocumentThumbnailSrc(file))}"
                  alt="${escapeHtml(getBijirisDocumentDisplayTitle(file, index))}"
                />
              </div>
              <div class="bijiris-document-editor-body">
                <label>
                  表示タイトル
                  <input
                    type="text"
                    value="${escapeHtml(getBijirisDocumentDisplayTitle(file, index))}"
                    data-bijiris-document-title="${index}"
                    placeholder="PDFタイトル"
                  />
                </label>
                <div class="meta">${escapeHtml(file.name || `document_${index + 1}.pdf`)}</div>
                <div class="action-row">
                  <label class="secondary-button upload-button">
                    表紙画像を追加
                    <input type="file" accept="image/*" data-bijiris-document-cover-input="${index}" hidden />
                  </label>
                  ${
                    file.thumbnailUrl || file.thumbnailDataUrl
                      ? `<button class="ghost-button" type="button" data-remove-bijiris-document-cover="${index}">表紙画像を削除</button>`
                      : ""
                  }
                  <button class="ghost-button" type="button" data-remove-bijiris-document="${index}">削除</button>
                </div>
              </div>
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

async function handleBijirisDocumentCoverInput(input, index) {
  const form = input?.form;
  syncBijirisDraftFromForm(form);
  const targetIndex = Number(index);
  if (!Number.isInteger(targetIndex) || !state.bijirisEditorDraft.documents[targetIndex]) return;
  const file = Array.from(input?.files || [])[0];
  if (!file) return;
  state.bijirisEditorDraft.documents[targetIndex].thumbnailDataUrl = await prepareBijirisDocumentCoverFile(file);
  state.bijirisEditorDraft.documents[targetIndex].thumbnailUrl = state.bijirisEditorDraft.documents[targetIndex].thumbnailDataUrl;
  state.bijirisEditorDraft.documents[targetIndex].thumbnailRemoved = false;
  renderBijirisManager();
  setPage("bijiris");
}

function removeBijirisDocumentCover(index, form) {
  syncBijirisDraftFromForm(form);
  const targetIndex = Number(index);
  if (!Number.isInteger(targetIndex) || !state.bijirisEditorDraft.documents[targetIndex]) return;
  state.bijirisEditorDraft.documents[targetIndex].thumbnailDataUrl = "";
  state.bijirisEditorDraft.documents[targetIndex].thumbnailUrl = "";
  state.bijirisEditorDraft.documents[targetIndex].thumbnailRemoved = true;
  renderBijirisManager();
  setPage("bijiris");
}

async function saveBijirisPost() {
  const draft = normalizeBijirisPost(getBijirisEditorDraft());
  if (!draft.title) throw new Error("タイトルを入力してください。");
  if (!draft.category) throw new Error("カテゴリを選択してください。");
  if (!draft.summary && !draft.body && !draft.photos.length && !draft.documents.length) {
    throw new Error("本文、要約、写真、PDF のいずれかを入力してください。");
  }
  const existingPost = draft.id ? getBijirisPostById(draft.id) : null;
  const effectiveStatus =
    existingPost?.status === "published" && draft.status !== "archived"
      ? "published"
      : draft.status;
  const effectivePublishedAt =
    effectiveStatus === "published"
      ? (draft.publishedAt || existingPost?.publishedAt || new Date().toISOString())
      : draft.publishedAt;
  const isPublished = effectiveStatus === "published";
  const reflectsToCustomerApp = Boolean(existingPost?.id) && existingPost.status === "published" && effectiveStatus === "published";
  showBusyToast(
    draft.id
      ? (reflectsToCustomerApp
          ? "豆知識を保存して顧客アプリへ反映中です。"
          : isPublished
            ? "豆知識の公開内容を保存中です。"
            : "豆知識を保存中です。")
      : (isPublished ? "豆知識を公開中です。" : "豆知識を作成中です。"),
  );
  const payload = {
    id: draft.id,
    title: draft.title,
    category: draft.category,
    summary: draft.summary,
    body: draft.body,
    status: effectiveStatus,
    pinned: draft.pinned,
    updatedAt: draft.updatedAt,
    publishedAt: effectivePublishedAt,
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
  showToast(
    draft.id
      ? (reflectsToCustomerApp
          ? "豆知識を保存し、顧客アプリへ反映しました。"
          : isPublished
            ? "豆知識の公開内容を保存しました。"
            : "豆知識を更新しました。")
      : (isPublished ? "豆知識を公開しました。" : "豆知識を作成しました。"),
  );
}

async function deleteBijirisPost(postId) {
  const targetId = String(postId || state.selectedBijirisPostId || "").trim();
  if (!targetId) return;
  const targetPost = getBijirisPostById(targetId);
  if (!(await openConfirmDialog({
    title: "豆知識削除",
    message: `「${targetPost?.title || "豆知識"}」を削除しますか？`,
    confirmLabel: "削除する",
    danger: true,
  }))) return;
  showBusyToast("豆知識を削除中です。");
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
  const categorySettingsButton = document.querySelector("#openBijirisCategorySettingsButton");
  const listSection = document.querySelector("#bijirisListSection");
  const editorSection = document.querySelector("#bijirisEditorSection");
  const stageTitle = document.querySelector("#bijirisStageTitle");
  if (!list || !editorCard || !createButton || !categorySettingsButton || !listSection || !editorSection || !stageTitle) return;

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

  stageTitle.textContent = view === "history" ? "投稿詳細" : "投稿一覧";
  listSection.hidden = view === "editor";
  editorSection.hidden = view !== "editor";
  createButton.hidden = view !== "list";
  categorySettingsButton.hidden = view !== "list";

  if (view === "history" && selectedPost) {
    list.innerHTML = `
      <div class="survey-editor-head">
        <div>
          <div class="card-title">投稿詳細</div>
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
            <article class="survey-manager-card selectable-card bijiris-list-card ${post.id === state.selectedBijirisPostId ? "active" : ""}">
              <button
                class="bijiris-list-open-button"
                type="button"
                data-open-bijiris-history="${escapeHtml(post.id)}"
                aria-label="${escapeHtml(post.title || "タイトル未入力")}"
              >
                <strong>${escapeHtml(post.title || "タイトル未入力")}</strong>
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
    categorySettingsButton.onclick = () => {
      setPage("settings");
      window.setTimeout(() => {
        document.querySelector("#preferencesCard")?.scrollIntoView({ behavior: "smooth", block: "start" });
      }, 80);
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
  const isEditingPublishedPost = Boolean(selectedPost?.id) && selectedPost.status === "published";
  if (isEditingPublishedPost && draft.status === "draft") {
    draft.status = "published";
  }
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
      ${renderBijirisCategoryEditor(draft.category)}
      <div class="survey-question-grid">
        <label>
          公開状態
          <select name="status">
            <option value="published" ${draft.status === "published" ? "selected" : ""}>公開</option>
            ${isEditingPublishedPost ? "" : `<option value="draft" ${draft.status === "draft" ? "selected" : ""}>下書き</option>`}
            <option value="archived" ${draft.status === "archived" ? "selected" : ""}>アーカイブ</option>
          </select>
          <div class="meta">
            ${
              isEditingPublishedPost
                ? "この投稿は保存すると同じ投稿を更新し、そのまま顧客アプリへ反映されます。顧客アプリから外したい場合のみ「アーカイブ」を選択してください。"
                : "顧客アプリに表示されるのは「公開」の投稿だけです。"
            }
          </div>
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
        <button
          id="bijirisSubmitButton"
          class="primary-button"
          type="submit"
          data-idle-label="${escapeHtml(getBijirisSubmitButtonLabel(draft, selectedPost, isEditingPublishedPost))}"
          ${hasPendingBijirisTasks() ? "disabled" : ""}
        >
          ${escapeHtml(hasPendingBijirisTasks() ? "画像を準備中…" : getBijirisSubmitButtonLabel(draft, selectedPost, isEditingPublishedPost))}
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
  document.querySelectorAll("[data-bijiris-category-control]").forEach((control) => {
    control.addEventListener("change", () => {
      syncBijirisDraftFromForm(form);
      renderBijirisManager();
      setPage("bijiris");
    });
  });
  form?.addEventListener("submit", (event) => {
    event.preventDefault();
    waitForBijirisPendingTasks()
      .then(() => {
        syncBijirisDraftFromForm(form);
        return saveBijirisPost();
      })
      .catch((error) => {
        showToast(error.message || "豆知識を保存できませんでした。");
      });
  });
  document.querySelector("#bijirisPhotoInput")?.addEventListener("change", (event) => {
    void runBijirisPendingTask(
      handleBijirisImageInput(event.currentTarget),
      "写真を追加できませんでした。",
    );
  });
  document.querySelector("#bijirisDocumentInput")?.addEventListener("change", (event) => {
    void runBijirisPendingTask(
      handleBijirisDocumentInput(event.currentTarget),
      "PDF を追加できませんでした。",
    );
  });
  document.querySelectorAll("[data-bijiris-document-cover-input]").forEach((input) => {
    input.addEventListener("change", (event) => {
      void runBijirisPendingTask(
        handleBijirisDocumentCoverInput(event.currentTarget, input.dataset.bijirisDocumentCoverInput),
        "表紙画像を追加できませんでした。",
      );
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
  document.querySelectorAll("[data-remove-bijiris-document-cover]").forEach((button) => {
    button.addEventListener("click", () => {
      removeBijirisDocumentCover(button.dataset.removeBijirisDocumentCover, form);
    });
  });
  document.querySelector("[data-delete-bijiris-post]")?.addEventListener("click", () => {
    deleteBijirisPost(state.selectedBijirisPostId).catch((error) => {
      showToast(error.message || "豆知識を削除できませんでした。");
    });
  });
  updateBijirisSubmitButtonState();
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

    const requiredToggle = event.target.closest("[data-question-required]");
    if (requiredToggle) {
      syncQuestionRequiredBadge(requiredToggle.closest("[data-question-block]"));
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
      const isPublished = normalizeSurveyStatus(payload.status) === "published";
      showBusyToast(
        surveyId
          ? (isPublished ? "アンケートの公開内容を保存中です。" : "アンケートを保存中です。")
          : (isPublished ? "アンケートを公開中です。" : "アンケートを作成中です。"),
      );
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
      showToast(
        surveyId
          ? (isPublished ? "アンケートの公開内容を保存しました。" : "アンケートを保存しました。")
          : (isPublished ? "アンケートを公開しました。" : "アンケートを作成しました。"),
      );
    } catch (error) {
      showToast(error.message || "アンケートを保存できませんでした。");
    } finally {
      submitButton.disabled = false;
      submitButton.textContent = surveyId ? "保存" : "作成";
    }
  });

  surveyEditorCard.querySelector("[data-delete-survey]")?.addEventListener("click", async () => {
    if (!selectedSurvey) return;
    if (!(await openConfirmDialog({
      title: "アンケート削除",
      message: `「${selectedSurvey.title}」を削除しますか？`,
      confirmLabel: "削除する",
      danger: true,
    }))) return;
    try {
      showBusyToast("アンケートを削除中です。");
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
  try {
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
      bijirisCategoryConfig: collectBijirisCategoryConfigFromForm(form),
    };
    const result = await api.request("/api/admin/preferences", {
      method: "PUT",
      token: state.token,
      body: payload,
    });
    state.preferences = result.preferences || payload;
    showToast("設定を保存しました。");
    renderAll();
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
      await Promise.all(cacheKeys.filter((key) => key.startsWith(CACHE_PREFIX)).map((key) => caches.delete(key)));
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

async function pruneAppCaches() {
  if (!("caches" in window)) return;
  const cacheKeys = await caches.keys();
  const staleKeys = cacheKeys.filter(
    (key) => key.startsWith(CACHE_PREFIX) && key !== ACTIVE_CACHE_NAME,
  );
  await Promise.all(staleKeys.map((key) => caches.delete(key)));
}

async function runAutoCacheMaintenance({ force = false } = {}) {
  try {
    const lastRun = Number(localStorage.getItem(AUTO_CACHE_MAINTENANCE_KEY) || "0");
    const now = Date.now();
    if (!force && now - lastRun < AUTO_CACHE_MAINTENANCE_INTERVAL_MS) return;
    localStorage.setItem(AUTO_CACHE_MAINTENANCE_KEY, String(now));

    if ("serviceWorker" in navigator) {
      const registration = await navigator.serviceWorker.getRegistration();
      await registration?.update().catch(() => {});
      registration?.waiting?.postMessage({ type: "SKIP_WAITING" });
    }
    await pruneAppCaches();
  } catch (error) {
    reportClientError("admin.autoCacheMaintenance", error);
  }
}

function scheduleAutoCacheMaintenance(force = false) {
  const task = () => {
    void runAutoCacheMaintenance({ force });
  };
  if ("requestIdleCallback" in window) {
    window.requestIdleCallback(task, { timeout: 10000 });
    return;
  }
  window.setTimeout(task, 1200);
}

function setupAutoCacheMaintenance() {
  window.addEventListener("load", () => {
    scheduleAutoCacheMaintenance();
  });
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      scheduleAutoCacheMaintenance();
    }
  });
  window.setInterval(() => {
    scheduleAutoCacheMaintenance(true);
  }, AUTO_CACHE_MAINTENANCE_INTERVAL_MS);
}

function setupInstall() {
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker
        .register("./sw.js?v=20260417-16", { updateViaCache: "none" })
        .then((registration) => {
          const activateWaiting = () => {
            registration.waiting?.postMessage({ type: "SKIP_WAITING" });
          };
          activateWaiting();
          registration.addEventListener("updatefound", () => {
            registration.installing?.addEventListener("statechange", () => {
              if (registration.waiting) activateWaiting();
            });
          });
          return registration.update().catch(() => {});
        })
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
setupAutoCacheMaintenance();
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
measurementPhotoModal?.addEventListener("click", (event) => {
  if (event.target === measurementPhotoModal || event.target.closest("[data-close-measurement-photo]")) {
    closeMeasurementPhotoModal();
  }
});
confirmDialogCancel?.addEventListener("click", () => {
  closeConfirmDialog(false);
});
confirmDialogConfirm?.addEventListener("click", () => {
  closeConfirmDialog(true);
});
confirmModal?.addEventListener("click", (event) => {
  if (event.target === confirmModal) {
    closeConfirmDialog(false);
  }
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && measurementPhotoModal && !measurementPhotoModal.hidden) {
    closeMeasurementPhotoModal();
    return;
  }
  if (event.key === "Escape" && confirmModal && !confirmModal.hidden) {
    closeConfirmDialog(false);
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
