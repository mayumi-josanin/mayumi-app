// 自宅PCから送られてくるデータベースの控えを、Googleドライブに保存する。
//
//   控えの設定を作る()   … 受け取り用の合鍵を作る（値は画面に出さない）
//   控えの状態を見る()   … いま何件・いつのものがあるかを見るだけ
//
// なぜこの向きなのか:
//   「GASがサーバーへ取りに行く」形にすると、サーバー側に
//   **データベース全体を返す窓口**を作ることになる。合鍵ひとつで会員情報が
//   丸ごと持ち出せる口を、インターネットに向けて開けたくない。
//   PCから送る形なら、その口を作らずに済む。
//
// なぜドライブに置くのか:
//   自宅PCが壊れたら、PCの中の控えも一緒に無くなる。会員情報は作り直せない。
//   置き場をPCの外にして初めて「備え」になる。

var 控え受取_フォルダ名 = 'まゆみ助産院 データベース控え';
var 控え受取_合鍵の鍵 = 'BACKUP_UPLOAD_KEY';
var 控え受取_残す件数 = 30;
var 控え受取_上限バイト = 25 * 1024 * 1024;   // 25MB。これを超えるなら送り方を見直す合図

// ---- doPost から呼ばれる本体 ----
//
// 既存の受け口（管理者・お客様.gs の doPost）は、管理者トークンを要求する。
// PCは管理者としてログインしないので、この処理だけは自分の合鍵で判断する。
// **保存先はこのフォルダに固定**で、任意の場所には書けないようにしてある。
function 控えを受け取る_(data) {
  var 合鍵 = PropertiesService.getScriptProperties().getProperty(控え受取_合鍵の鍵);
  if (!合鍵) {
    return { status: 'error', message: '受け取りの設定がまだです。' };
  }
  if (String(data && data.key || '') !== 合鍵) {
    // どこが違うかは知らせない。総当たりの手がかりを与えないため。
    return { status: 'error', message: '受け取れませんでした。' };
  }

  var 名前 = String(data.filename || '').replace(/[^A-Za-z0-9._-]/g, '');
  if (!/^mayumi-\d{8}-\d{4}\.dump$/.test(名前)) {
    return { status: 'error', message: 'ファイル名の形が違います。' };
  }

  var 中身;
  try {
    中身 = Utilities.base64Decode(String(data.content || ''));
  } catch (e) {
    return { status: 'error', message: '中身を読み取れませんでした。' };
  }
  if (!中身 || !中身.length) {
    return { status: 'error', message: '中身が空です。' };
  }
  if (中身.length > 控え受取_上限バイト) {
    return { status: 'error', message: '大きすぎます（' + 中身.length + 'バイト）。' };
  }

  var folder = 控え受取_フォルダを用意する_();

  // 同じ名前が来たら置き換える。同じ日に2回動いても増え続けないように。
  var 既存 = folder.getFilesByName(名前);
  while (既存.hasNext()) 既存.next().setTrashed(true);

  var file = folder.createFile(Utilities.newBlob(中身, 'application/octet-stream', 名前));
  var 消した = 控え受取_古いものを片付ける_(folder);

  return {
    status: 'ok',
    saved: 名前,
    bytes: 中身.length,
    removed: 消した,
    folder: folder.getName()
  };
}

function 控え受取_フォルダを用意する_() {
  var it = DriveApp.getFoldersByName(控え受取_フォルダ名);
  return it.hasNext() ? it.next() : DriveApp.createFolder(控え受取_フォルダ名);
}

// 新しいものから数えて 控え受取_残す件数 まで残し、それより古いものを捨てる。
function 控え受取_古いものを片付ける_(folder) {
  var 一覧 = [];
  var it = folder.getFilesByType('application/octet-stream');
  while (it.hasNext()) {
    var f = it.next();
    if (/^mayumi-\d{8}-\d{4}\.dump$/.test(f.getName())) {
      一覧.push({ file: f, 名: f.getName() });
    }
  }
  // 名前に日時が入っているので、名前で並べれば新しい順になる。
  一覧.sort(function (a, b) { return a.名 < b.名 ? 1 : -1; });
  var 消した = 0;
  一覧.slice(控え受取_残す件数).forEach(function (x) {
    x.file.setTrashed(true);
    消した += 1;
  });
  return 消した;
}

// ---- 受付で使う道具 ----

// いま何件・いつのものがあるかを見るだけ。
function 控えの状態を見る() {
  var it = DriveApp.getFoldersByName(控え受取_フォルダ名);
  if (!it.hasNext()) {
    Logger.log('■ 保存先のフォルダがまだありません。');
    Logger.log('  控えの設定を作る() を先に実行してください。');
    return;
  }
  var folder = it.next();
  var 一覧 = [];
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    一覧.push({ 名: f.getName(), 大きさ: f.getSize(), 日: f.getDateCreated() });
  }
  一覧.sort(function (a, b) { return a.名 < b.名 ? 1 : -1; });

  Logger.log('■ ' + folder.getName() + ': ' + 一覧.length + '件');
  Logger.log('  ' + folder.getUrl());
  Logger.log('');
  一覧.slice(0, 10).forEach(function (x) {
    Logger.log('    ' + x.名 + '  ' + Math.round(x.大きさ / 1024) + 'KB  ' +
      Utilities.formatDate(x.日, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'));
  });
  if (一覧.length > 10) Logger.log('    …ほか ' + (一覧.length - 10) + '件');
  Logger.log('');
  if (!一覧.length) {
    Logger.log('  ※ まだ1件も届いていません。PC側の送信が動いているか確かめてください。');
  } else {
    var 最新 = 一覧[0];
    var 経過 = Math.round((Date.now() - 最新.日.getTime()) / 3600000);
    Logger.log('  いちばん新しい控え: ' + 経過 + '時間前');
    if (経過 > 30) {
      Logger.log('  **30時間以上届いていません。**PC側の毎日の実行が止まっている可能性があります。');
    }
  }
}

// 受け取り用の合鍵を作る。**値は画面に出さない。**
// 値を見るときは、GASの「プロジェクトの設定」→「スクリプト プロパティ」で
// BACKUP_UPLOAD_KEY を開いてください。そこからPCの .env に貼ります。
function 控えの設定を作る() {
  var props = PropertiesService.getScriptProperties();
  var いま = props.getProperty(控え受取_合鍵の鍵);
  if (いま) {
    Logger.log('■ 受け取りの合鍵は、すでに作られています（' + いま.length + '文字）');
    Logger.log('');
    Logger.log('  作り直したいときは、先にスクリプトプロパティから');
    Logger.log('  ' + 控え受取_合鍵の鍵 + ' を消してから、もう一度実行してください。');
    return;
  }

  var 文字 = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  var 鍵 = '';
  for (var i = 0; i < 50; i += 1) {
    鍵 += 文字.charAt(Math.floor(Math.random() * 文字.length));
  }
  props.setProperty(控え受取_合鍵の鍵, 鍵);

  var folder = 控え受取_フォルダを用意する_();

  Logger.log('■ 受け取りの合鍵を作りました（' + 鍵.length + '文字）');
  Logger.log('  **値はここに出しません。**');
  Logger.log('');
  Logger.log('  見かた:');
  Logger.log('    左の歯車（プロジェクトの設定）→ スクリプト プロパティ');
  Logger.log('    → ' + 控え受取_合鍵の鍵 + ' の値をコピー');
  Logger.log('');
  Logger.log('■ 保存先のフォルダ: ' + folder.getName());
  Logger.log('  ' + folder.getUrl());
}
