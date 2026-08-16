// 自宅PCのデータベースが生きているかを見張る道具。
//
//   サーバーの生死を見る()   … いま届くかを1回だけ見る（何も送らない）
//   見張りを始める()         … 15分おきの見張りを仕掛ける
//   見張りを止める()         … 見張りを外す
//
// なぜ要るのか:
//   データベースを自宅のPCに置くため、停電・Windows Update の再起動・
//   回線の不調で止まっても、誰も気づけない。院に置いてあれば目に入るが、
//   自宅だと次に使うまで分からない。だから、こちらから見に行く。
//
// メールは「状態が変わったときだけ」送る。15分おきに届いたら誰も読まなくなり、
// 本当に困ったときの1通が埋もれてしまう。

// 見に行き先。スクリプトプロパティ DB_HEALTH_URL があればそちらを使う。
var 見張り_既定のURL = 'https://desktop-rmsk0vg.tail8efe0d.ts.net:10000/api/health';

// 何回続けて届かなかったら知らせるか。
// 1回で送ると、回線の一瞬の不調でも鳴ってしまう。2回＝約30分止まって初めて知らせる。
var 見張り_知らせるまでの回数 = 2;

var 見張り_状態の鍵 = 'DB_HEALTH_STATE';
var 見張り_扱う関数名 = 'サーバーを見張る';

function サーバーの生死を見る() {
  var r = 見張り_叩く_();
  Logger.log('■ ' + 見張り_URLを取る_());
  Logger.log('');
  if (r.ok) {
    Logger.log('  届きました（' + r.かかった + 'ミリ秒）');
    Logger.log('  応答: ' + r.本文);
  } else {
    Logger.log('  **届きません**');
    Logger.log('  理由: ' + r.理由);
  }
  Logger.log('');
  var 状態 = 見張り_状態を読む_();
  Logger.log('  いまの記録: ' + (状態.最後 === 'ng'
    ? '止まっていると判断中（' + 状態.連続失敗 + '回連続）'
    : '動いていると判断中'));
  Logger.log('');
  Logger.log('  ※ これは見るだけです。メールは送りません。');
}

// 15分おきに呼ばれる。ここだけがメールを送る。
function サーバーを見張る() {
  var r = 見張り_叩く_();
  var 状態 = 見張り_状態を読む_();

  if (r.ok) {
    if (状態.最後 === 'ng') {
      var 止まっていた = 見張り_経過を書く_(状態.落ちた時刻);
      sendGmail(
        '【復旧】まゆみ助産院データベース',
        'データベースにまた届くようになりました。\n\n' +
        '止まっていた時間: ' + 止まっていた + '\n' +
        '見に行き先: ' + 見張り_URLを取る_() + '\n\n' +
        'この間、引っ越し済みの記録は見られない状態でした。'
      );
    }
    見張り_状態を書く_({ 最後: 'ok', 連続失敗: 0, 落ちた時刻: '' });
    return;
  }

  var 連続 = Number(状態.連続失敗 || 0) + 1;
  var 落ちた時刻 = 状態.落ちた時刻 || new Date().toISOString();

  // まだ知らせる回数に届いていない、またはすでに知らせ済み。
  if (連続 < 見張り_知らせるまでの回数 || 状態.最後 === 'ng') {
    見張り_状態を書く_({ 最後: 連続 >= 見張り_知らせるまでの回数 ? 'ng' : 'ok', 連続失敗: 連続, 落ちた時刻: 落ちた時刻 });
    return;
  }

  sendGmail(
    '【停止】まゆみ助産院データベースに届きません',
    'データベースに ' + 連続 + '回続けて届きませんでした（約' + (連続 * 15) + '分）。\n\n' +
    '見に行き先: ' + 見張り_URLを取る_() + '\n' +
    '理由: ' + r.理由 + '\n\n' +
    '確かめるところ:\n' +
    '  1. 自宅のPCの電源が入っているか\n' +
    '  2. Docker Desktop が動いているか\n' +
    '  3. PowerShell で docker compose ps（server フォルダで）\n' +
    '  4. tailscale funnel status に 10000 が出ているか\n\n' +
    '復旧したら、また自動でお知らせします。'
  );
  見張り_状態を書く_({ 最後: 'ng', 連続失敗: 連続, 落ちた時刻: 落ちた時刻 });
}

function 見張りを始める() {
  見張りを止める();
  ScriptApp.newTrigger(見張り_扱う関数名).timeBased().everyMinutes(15).create();
  見張り_状態を書く_({ 最後: 'ok', 連続失敗: 0, 落ちた時刻: '' });
  Logger.log('■ 15分おきの見張りを始めました');
  Logger.log('  見に行き先: ' + 見張り_URLを取る_());
  Logger.log('  知らせ先: ' + CONFIG.GMAIL_TO);
  Logger.log('');
  Logger.log('  約30分止まったら1通、直ったらもう1通だけ届きます。');
}

function 見張りを止める() {
  var 消した = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 見張り_扱う関数名) {
      ScriptApp.deleteTrigger(t);
      消した += 1;
    }
  });
  Logger.log('■ 見張りを ' + 消した + '件 外しました');
}

function 見張り_URLを取る_() {
  return String(
    PropertiesService.getScriptProperties().getProperty('DB_HEALTH_URL') || 見張り_既定のURL
  ).trim();
}

function 見張り_叩く_() {
  var 始め = Date.now();
  try {
    var res = UrlFetchApp.fetch(見張り_URLを取る_(), {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      // 自宅回線なので、遅いときは待たずに諦める。待ち続けると次の見張りと重なる。
      validateHttpsCertificates: true,
    });
    var 本文 = String(res.getContentText() || '').slice(0, 200);
    var 番号 = res.getResponseCode();
    if (番号 !== 200) {
      return { ok: false, 理由: 'HTTP ' + 番号 + ' が返りました', かかった: Date.now() - 始め, 本文: 本文 };
    }
    // 中身まで見る。Funnel だけ生きていて中の Django が落ちている場合を拾うため。
    if (本文.indexOf('ok') < 0) {
      return { ok: false, 理由: '応答の中身が想定と違います: ' + 本文, かかった: Date.now() - 始め, 本文: 本文 };
    }
    return { ok: true, 理由: '', かかった: Date.now() - 始め, 本文: 本文 };
  } catch (e) {
    return { ok: false, 理由: String(e && e.message ? e.message : e), かかった: Date.now() - 始め, 本文: '' };
  }
}

function 見張り_状態を読む_() {
  try {
    var 生 = PropertiesService.getScriptProperties().getProperty(見張り_状態の鍵);
    var v = 生 ? JSON.parse(生) : null;
    if (!v || typeof v !== 'object') return { 最後: 'ok', 連続失敗: 0, 落ちた時刻: '' };
    return v;
  } catch (e) {
    return { 最後: 'ok', 連続失敗: 0, 落ちた時刻: '' };
  }
}

function 見張り_状態を書く_(v) {
  PropertiesService.getScriptProperties().setProperty(見張り_状態の鍵, JSON.stringify(v || {}));
}

function 見張り_経過を書く_(落ちた時刻) {
  if (!落ちた時刻) return '不明';
  var 分 = Math.round((Date.now() - new Date(落ちた時刻).getTime()) / 60000);
  if (分 < 60) return 見張り_分で書く_(分);
  var 時 = Math.floor(分 / 60);
  return 時 + '時間' + (分 % 60 ? 見張り_分で書く_(分 % 60) : '');
}

function 見張り_分で書く_(分) {
  return 分 + '分';
}
