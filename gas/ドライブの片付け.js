// 作業でできたファイルを片付ける道具。
//
// データを消す関数は必ず控えを取るようにしてあるので、作業のたびに
// 【控え】…が増える。放っておくとマイドライブの一番上に溜まり、
// どれが本番のファイルなのか分からなくなる。
//
// ここでやるのは「まとめる」まで。消すのは別の関数にしてある。
//   作業ファイルの下見()      … どれを動かすか見るだけ
//   作業ファイルをまとめる()  … 「アプリ/作業の控え」へ移す
//   古い作業ファイルをゴミ箱へ() … 日数を決めてゴミ箱に入れる（戻せる）

// エディタは「ファイルの先頭の関数」を最初に選ぶ。
// 既定は下見にしておく。実行するときだけ、ここを呼びたい関数に書き換える。
function いま実行する() {
  作業ファイルの下見();
}

// 「まゆみ助産院 / アプリ」フォルダ。ここの下にまとめる。
var 片付け_アプリフォルダID = '1NrHV4IYrzY_95Btpp7vqgieH091hpZTl';
var 片付け_まとめ先の名前 = '作業の控え';

// 動かす対象。お名前の頭で見分ける。
// 本番のファイル（「まゆみ助産院_管理」「ビジリス 会員別まとめ」など）は
// この形に当てはまらないので、間違って動かすことはない。
var 片付け_対象の頭 = [
  '【控え】',
  '移行の下見_',
  '会員一覧_',
  '測定履歴_',
  'まゆみ助産院_管理_退避_'
];

// 自動バックアップの置き場。ここの中は触らない（毎晩の仕組みが使っている）。
var 片付け_触らないフォルダ = ['まゆみ助産院_バックアップ', 'system-backups'];

function 片付け_対象か_(名前) {
  var s = String(名前 || '');
  return 片付け_対象の頭.some(function (t) { return s.indexOf(t) === 0; });
}

function 片付け_親の名前_(file) {
  var it = file.getParents();
  return it.hasNext() ? it.next().getName() : '（マイドライブ直下）';
}

// 対象のファイルを集める。名前で探すより、頭の一致で確かめる方が取りこぼさない。
function 片付け_集める_() {
  var 見つけた = [];
  var 済み = {};
  片付け_対象の頭.forEach(function (頭) {
    var it = DriveApp.searchFiles("title contains '" + 頭.replace(/'/g, "\\'") + "'");
    while (it.hasNext()) {
      var f = it.next();
      if (済み[f.getId()]) continue;
      if (!片付け_対象か_(f.getName())) continue;
      var 親 = 片付け_親の名前_(f);
      if (片付け_触らないフォルダ.indexOf(親) >= 0) continue;
      if (親 === 片付け_まとめ先の名前) continue;   // もう片付いている
      済み[f.getId()] = true;
      見つけた.push(f);
    }
  });
  見つけた.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  return 見つけた;
}

function 片付け_日数_(file) {
  return Math.floor((new Date() - file.getDateCreated()) / 86400000);
}

function 作業ファイルの下見() {
  var files = 片付け_集める_();
  Logger.log('■ 片付ける対象（まだ動かしていません）');
  Logger.log('');
  if (!files.length) {
    Logger.log('  散らかっているファイルはありません。');
    return;
  }
  files.forEach(function (f) {
    Logger.log('  ' + f.getName());
    Logger.log('    いまの場所: ' + 片付け_親の名前_(f) + ' ／ ' + 片付け_日数_(f) + '日前');
  });
  Logger.log('');
  Logger.log('  合計 ' + files.length + '件');
  Logger.log('  「アプリ/' + 片付け_まとめ先の名前 + '」へ移すには 作業ファイルをまとめる() を実行してください。');
}

function 片付け_まとめ先_() {
  var 親 = DriveApp.getFolderById(片付け_アプリフォルダID);
  var it = 親.getFoldersByName(片付け_まとめ先の名前);
  return it.hasNext() ? it.next() : 親.createFolder(片付け_まとめ先の名前);
}

function 作業ファイルをまとめる() {
  var files = 片付け_集める_();
  if (!files.length) { Logger.log('散らかっているファイルはありません。'); return; }

  var 先 = 片付け_まとめ先_();
  var 動かした = 0;
  files.forEach(function (f) {
    f.moveTo(先);
    Logger.log('  移しました: ' + f.getName());
    動かした += 1;
  });

  Logger.log('');
  Logger.log('■ ' + 動かした + '件を「アプリ/' + 片付け_まとめ先の名前 + '」へ移しました');
  Logger.log('  ' + 先.getUrl());
  Logger.log('');
  Logger.log('  中身は消していません。要らなくなったら 古い作業ファイルをゴミ箱へ() を使ってください。');
}

// 何日より古いものをゴミ箱に入れるか。ゴミ箱なので30日は戻せる。
var 片付け_残す日数 = 14;

function 古い作業ファイルの下見() {
  var 先 = 片付け_まとめ先_();
  var it = 先.getFiles();
  var 対象 = [], 残す = [];
  while (it.hasNext()) {
    var f = it.next();
    (片付け_日数_(f) > 片付け_残す日数 ? 対象 : 残す).push(f);
  }
  Logger.log('■ ' + 片付け_残す日数 + '日より古いもの（まだ消していません）');
  Logger.log('');
  対象.forEach(function (f) { Logger.log('  ゴミ箱へ: ' + f.getName() + '（' + 片付け_日数_(f) + '日前）'); });
  Logger.log('');
  残す.forEach(function (f) { Logger.log('  残す: ' + f.getName() + '（' + 片付け_日数_(f) + '日前）'); });
  Logger.log('');
  Logger.log('  ゴミ箱へ ' + 対象.length + '件 / 残す ' + 残す.length + '件');
}

function 古い作業ファイルをゴミ箱へ() {
  var 先 = 片付け_まとめ先_();
  var it = 先.getFiles();
  var 消した = 0;
  while (it.hasNext()) {
    var f = it.next();
    if (片付け_日数_(f) <= 片付け_残す日数) continue;
    Logger.log('  ゴミ箱へ: ' + f.getName());
    f.setTrashed(true);
    消した += 1;
  }
  Logger.log('');
  Logger.log('■ ' + 消した + '件をゴミ箱に入れました（30日は元に戻せます）');
}
