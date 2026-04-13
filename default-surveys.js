const BIJIRIS_SESSION_CONCERN_CATEGORIES = [
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

const BIJIRIS_SESSION_TICKET_SHEET_OPTIONS = [
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

const BIJIRIS_SESSION_TICKET_ROUND_OPTIONS = [
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

const BIJIRIS_SESSION_LIFE_CHANGE_OPTIONS = [
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

function getBijirisSessionConcernOptions() {
  return BIJIRIS_SESSION_CONCERN_CATEGORIES.flatMap((category) => category.options);
}

function makeDefaultSurveys(timestamp = new Date().toISOString()) {
  return [
    {
      id: "survey_bijiris_session",
      title: "ビジリス施術アンケート",
      description: "ビジリス施術後の体感やお悩みをお聞かせください。",
      introMessage: "施術内容を選択後、本日の体感や気になることをご回答ください。",
      completionMessage: "ビジリス施術アンケートのご回答ありがとうございました。",
      status: "published",
      createdAt: timestamp,
      updatedAt: timestamp,
      questions: [
        {
          id: "q_bijiris_session_type",
          label: "施術内容",
          type: "choice",
          required: true,
          options: ["初回お試し", "回数券", "単発", "トライアル"],
        },
        {
          id: "q_bijiris_session_ticket_plan",
          label: "回数券の種類",
          type: "choice",
          required: true,
          options: ["6回券", "10回券"],
          visibleWhen: { questionId: "q_bijiris_session_type", value: "回数券" },
        },
        {
          id: "q_bijiris_session_ticket_sheet",
          label: "回数券の何枚目ですか？",
          type: "choice",
          required: true,
          options: BIJIRIS_SESSION_TICKET_SHEET_OPTIONS,
          visibleWhen: { questionId: "q_bijiris_session_type", value: "回数券" },
        },
        {
          id: "q_bijiris_session_ticket_round",
          label: "回数券の何回目ですか？",
          type: "choice",
          required: true,
          options: BIJIRIS_SESSION_TICKET_ROUND_OPTIONS,
          visibleWhen: { questionId: "q_bijiris_session_type", value: "回数券" },
        },
        {
          id: "q_bijiris_session_feeling",
          label: "本日のビジリスの体感はいかがでしたか？　以前と比べて変化したことなどがあればご記載ください",
          type: "textarea",
          required: true,
          options: [],
        },
        {
          id: "q_bijiris_session_concern",
          label:
            "普段のお身体のお悩みや、ビジリス（骨盤底筋ケア）について気になること・知りたいことはありますか？（複数選択可）",
          type: "checkbox",
          required: false,
          options: getBijirisSessionConcernOptions(),
        },
        {
          id: "q_bijiris_session_life_changes",
          label: "日常生活にどのような変化がありましたか？（複数選択可）",
          type: "checkbox",
          required: false,
          options: BIJIRIS_SESSION_LIFE_CHANGE_OPTIONS,
        },
        {
          id: "q_bijiris_session_life_changes_other",
          label: "日常生活の変化（その他）",
          type: "textarea",
          required: false,
          options: [],
          visibleWhen: {
            questionId: "q_bijiris_session_life_changes",
            value: "その他（自由記述）",
          },
        },
        {
          id: "q_bijiris_session_monitor_photos_6",
          label: "モニター時の写真2枚",
          type: "photo",
          required: true,
          options: [],
          visibilityConditions: [
            { questionId: "q_bijiris_session_type", value: "回数券" },
            { questionId: "q_bijiris_session_ticket_plan", value: "6回券" },
            { questionId: "q_bijiris_session_ticket_round", value: "6回目" },
          ],
        },
        {
          id: "q_bijiris_session_ticket_end_photos_6",
          label: "回数券終了時の写真2枚",
          type: "photo",
          required: true,
          options: [],
          visibilityConditions: [
            { questionId: "q_bijiris_session_type", value: "回数券" },
            { questionId: "q_bijiris_session_ticket_plan", value: "6回券" },
            { questionId: "q_bijiris_session_ticket_round", value: "6回目" },
          ],
        },
        {
          id: "q_bijiris_session_monitor_photos_10",
          label: "モニター時の写真2枚",
          type: "photo",
          required: true,
          options: [],
          visibilityConditions: [
            { questionId: "q_bijiris_session_type", value: "回数券" },
            { questionId: "q_bijiris_session_ticket_plan", value: "10回券" },
            { questionId: "q_bijiris_session_ticket_round", value: "10回目" },
          ],
        },
        {
          id: "q_bijiris_session_ticket_end_photos_10",
          label: "回数券終了時の写真2枚",
          type: "photo",
          required: true,
          options: [],
          visibilityConditions: [
            { questionId: "q_bijiris_session_type", value: "回数券" },
            { questionId: "q_bijiris_session_ticket_plan", value: "10回券" },
            { questionId: "q_bijiris_session_ticket_round", value: "10回目" },
          ],
        },
      ],
    },
  ];
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { makeDefaultSurveys };
}

if (typeof window !== "undefined") {
  window.MayumiDefaultSurveys = makeDefaultSurveys;
}
