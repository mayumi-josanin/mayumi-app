// 会員データを、データベースへ取り込める形（JSON）で書き出す道具。
//
//   会員データの下見()    … 何件あるか・欠けはないかを見るだけ
//   会員データを書き出す() … JSONにしてドライブへ置く
//
// **どちらも読むだけ。**スプレッドシートは一切変えない。
//
// 使いかた:
//   1. 会員データの下見() で件数を確かめる
//   2. 会員データを書き出す() を実行し、ログのURLからダウンロードする
//   3. サーバー側で
//        python manage.py 会員を取り込む 会員データ.json --下見
//   4. 件数が合っていれば --下見 を外して実行
//
// パスコードは平文のまま書き出す。**このJSONは扱いに注意すること。**
//   ・取り込みが済んだらファイルを消す（ドライブにも手元にも残さない）
//   ・サーバー側はハッシュにして保存するので、DBには平文は入らない
//   ・そもそも平文で持っている今の状態を直すのが、この移行の目的の1つ

var 会員書出_シート名 = '会員データ';

function 会員データを書き出す() {
  var r = 会員書出_集める_();
  var 中身 = JSON.stringify({
    書き出した日時: Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX"),
    件数: r.members.length,
    members: r.members
  }, null, 2);

  var 名前 = '会員データ_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, 中身, MimeType.PLAIN_TEXT);

  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('');
  Logger.log('  会員: ' + r.members.length + '名 / 除いた行: ' + r.除いた.length + '件');
  Logger.log('');
  Logger.log('  ※ このファイルにはパスコードが平文で入っています。');
  Logger.log('     取り込みが済んだら、ドライブからも手元からも消してください。');
}

function 会員データの下見() {
  var r = 会員書出_集める_();
  Logger.log('■ 会員データ: ' + r.全体 + '行');
  Logger.log('');
  Logger.log('  書き出せる: ' + r.members.length + '名');
  Logger.log('  除いた行:   ' + r.除いた.length + '件');
  r.除いた.slice(0, 20).forEach(function (x) { Logger.log('    ' + x); });
  if (r.除いた.length > 20) Logger.log('    …ほか ' + (r.除いた.length - 20) + '件');
  Logger.log('');
  Logger.log('■ 中身の埋まりぐあい');
  ['kana', 'phone', 'birthday', 'address', 'passcode',
   'memo', 'pushSubscription', 'stampAchievedAt'].forEach(function (k) {
    var n = r.members.filter(function (m) { return m[k]; }).length;
    Logger.log('    ' + 会員書出_見出し_(k) + ': ' + n + '名 / ' + r.members.length + '名');
  });
  Logger.log('');
  var 消した = r.members.filter(function (m) { return m.deleted || m.deletedAt; }).length;
  var ビジ = r.members.filter(function (m) { return m.bijirisRegistered; }).length;
  Logger.log('    退会の印がある方: ' + 消した + '名（消さずに印だけ移します）');
  Logger.log('    ビジリス登録:     ' + ビジ + '名');
  Logger.log('');
  Logger.log('  よければ 会員データを書き出す() を実行してください。');
}

function 会員書出_見出し_(鍵) {
  return ({
    kana: 'フリガナ', phone: '電話番号', birthday: '生年月日',
    address: '住所', passcode: 'パスコード',
    memo: 'メモ', pushSubscription: '通知の届け先', stampAchievedAt: 'スタンプ達成日時'
  })[鍵] || 鍵;
}

function 会員書出_空か_(x) { return x === '' || x === null || x === undefined; }

function 会員書出_文字_(値) {
  return 会員書出_空か_(値) ? '' : String(値).trim();
}

function 会員書出_日付_(値) {
  if (会員書出_空か_(値)) return null;
  var d = 値 instanceof Date ? 値 : new Date(値);
  if (isNaN(d.getTime())) return null;
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function 会員書出_日時_(値) {
  if (会員書出_空か_(値)) return null;
  var d = 値 instanceof Date ? 値 : new Date(値);
  if (isNaN(d.getTime())) return null;
  return Utilities.formatDate(d, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function 会員書出_集める_() {
  var sheet = getOrCreateUsersSheet_(getOrCreateSpreadsheet());
  var 最終 = sheet.getLastRow();
  if (最終 < 2) return { 全体: 0, members: [], 除いた: [] };

  var 値 = sheet.getRange(2, 1, 最終 - 1, USER_HEADERS.length).getValues();
  var members = [], 除いた = [];

  値.forEach(function (row, i) {
    var 行番号 = i + 2;
    var 会員ID = 会員書出_文字_(row[USER_COL.MEMBER_ID - 1]);
    var 氏名 = 会員書出_文字_(row[USER_COL.NAME - 1]);

    // 会員IDかお名前が無い行は、データベースに置き場が無い。
    // 黙って落とさず、何行目だったかを残す。
    if (!会員ID || !氏名) {
      除いた.push(行番号 + '行目: ' +
        (会員ID ? '' : '会員IDが空 ') + (氏名 ? '' : 'お名前が空 ') +
        '（' + (会員ID || '—') + ' / ' + (氏名 || '—') + '）');
      return;
    }

    members.push({
      memberId: 会員ID,
      createdAt: 会員書出_日時_(row[USER_COL.TIMESTAMP - 1]),
      name: 氏名,
      kana: 会員書出_文字_(row[USER_COL.KANA - 1]),
      phone: 会員書出_文字_(row[USER_COL.PHONE - 1]),
      birthday: 会員書出_日付_(row[USER_COL.BIRTHDAY - 1]),
      address: 会員書出_文字_(row[USER_COL.ADDRESS - 1]),
      avatarUrl: 会員書出_文字_(row[USER_COL.AVATAR_URL - 1]),
      // **8列目は「オン/オフ」ではなく、通知の届け先そのもの。**
      // GAS の getPushUsers() が、この値を subscription として配信に渡している。
      // 真偽値の false がそのまま入っていることがあるので、文字にすると "false"
      // になる。取り込み側で「オフ」に寄せている。
      //
      // 2026-08-23 まで、取り込み側がこれを 真偽() に通していた。
      // 購読IDは "true" でも "1" でもないので**偽に落ちていた**（1名該当）。
      pushSubscription: 会員書出_文字_(row[USER_COL.PUSH - 1]),
      pushEnabled: 会員書出_文字_(row[USER_COL.PUSH - 1]),
      // 受付の覚え書き。2026-08-23 時点で 0名（お客様アプリが保存のたびに
      // memo:'' を送って消しているため）。それでも列としては移す。
      memo: 会員書出_文字_(row[USER_COL.MEMO - 1]),
      // スタンプが10個そろった日時。**特典の有効期限の基準。**
      stampAchievedAt: 会員書出_日時_(row[USER_COL.STAMP_ACHIEVED_AT - 1]),
      status: 会員書出_文字_(row[USER_COL.STATUS - 1]),
      stampCount: row[USER_COL.STAMP_COUNT - 1],
      stampCardNumber: row[USER_COL.STAMP_CARD_NUM - 1],
      lastStampAt: 会員書出_日時_(row[USER_COL.LAST_STAMP_AT - 1]) ||
                   会員書出_日時_(row[USER_COL.LAST_STAMP_DATE - 1]),
      passcode: 会員書出_文字_(row[USER_COL.PASSCODE - 1]),
      passwordHash: 会員書出_文字_(row[USER_COL.PASSWORD_HASH - 1]),
      passwordSalt: 会員書出_文字_(row[USER_COL.PASSWORD_SALT - 1]),
      role: 会員書出_文字_(row[USER_COL.ROLE - 1]),
      transferCode: 会員書出_文字_(row[USER_COL.TRANSFER_CODE - 1]),
      transferCodeIssuedAt: 会員書出_日時_(row[USER_COL.TRANSFER_CODE_ISSUED_AT - 1]),
      deviceSessions: 会員書出_文字_(row[USER_COL.DEVICE_SESSIONS - 1]),
      stampHistory: 会員書出_文字_(row[USER_COL.STAMP_HISTORY_JSON - 1]),
      rewardHistory: 会員書出_文字_(row[USER_COL.REWARDS - 1]),
      // 印と日時は別々に渡す。印だけあって日時が無い方がいるため、
      // 日時に寄せると実在しない日付を作ることになる。
      deleted: 会員書出_文字_(row[USER_COL.DELETE_STATUS - 1]) ? 'true' : '',
      deletedAt: 会員書出_日時_(row[USER_COL.DELETED_AT - 1]),
      mergedIntoId: 会員書出_文字_(row[USER_COL.MERGED_INTO - 1]),
      registrationSource: 会員書出_文字_(row[USER_COL.REGISTRATION_SOURCE - 1]),
      registrationSourceDetail: 会員書出_文字_(row[USER_COL.REGISTRATION_SOURCE_DETAIL - 1]),
      registrationSourceUpdatedAt: 会員書出_日時_(row[USER_COL.REGISTRATION_SOURCE_UPDATED_AT - 1]),
      lastOnlineAt: 会員書出_日時_(row[USER_COL.LAST_ONLINE_AT - 1]),
      bijirisRegistered: 会員書出_文字_(row[USER_COL.BIJIRIS - 1])
    });
  });

  return { 全体: 値.length, members: members, 除いた: 除いた };
}
