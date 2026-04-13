const BIJIRIS_SESSION_CONCERN_CATEGORIES = [
  {
    id: "toilet",
    label: "【トイレ・デリケートゾーンのお悩み】",
    options: [
      "咳、くしゃみ、大笑いした時に尿もれすることがある",
      "ジャンプや運動、重いものを持った時に尿もれすることがある",
      "急にトイレに行きたくなる、我慢できず間に合わないことがある",
      "トイレが近い・夜中に何度もトイレで起きる",
      "デリケートゾーンの違和感や乾燥が気になる",
      "便秘しやすい・お腹が張りやすい",
    ],
  },
  {
    id: "posture",
    label: "【姿勢・体型のお悩み】",
    options: [
      "姿勢が崩れやすい・猫背や反り腰が気になる",
      "下腹ぽっこりが気になる",
      "ヒップラインや骨盤まわりのゆるみが気になる",
      "産後の体型変化が気になる",
    ],
  },
  {
    id: "pain",
    label: "【痛み・めぐりのお悩み】",
    options: [
      "腰痛がある",
      "股関節や骨盤まわりに痛み・違和感がある",
      "冷えやむくみが気になる",
      "疲れやすい・眠りが浅い",
    ],
  },
  {
    id: "care",
    label: "【ビジリス（骨盤底筋ケア）について知りたいこと】",
    options: [
      "骨盤底筋ケアでどんな変化が期待できるか知りたい",
      "自宅でできる骨盤底筋ケアを知りたい",
      "妊活や産後ケアにどう役立つか知りたい",
      "自分に合う通い方や頻度を知りたい",
      "その他",
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
          id: "q_bijiris_session_monitor_photos",
          label: "モニター時の写真2枚",
          type: "photo",
          required: true,
          options: [],
          visibleWhen: { questionId: "q_bijiris_session_type", value: "回数券" },
        },
        {
          id: "q_bijiris_session_ticket_end_photos",
          label: "回数券終了時の写真2枚",
          type: "photo",
          required: true,
          options: [],
          visibleWhen: { questionId: "q_bijiris_session_type", value: "回数券" },
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
