// 控えの保存先フォルダを、IDで固定する道具。
//
//   フォルダの下見()     … 同じ名前のフォルダがいくつあるかを見るだけ
//   フォルダを固定する() … 控えが実際に入っているフォルダを設定に書く
//
// なぜ要るのか:
//   保存先を名前で探していたため、同じ名前のフォルダが2つでき、
//   受け口が保存する先と、画面で数える先が食い違った。
//   「送れているのに0件」に見え、原因を掴むまでに遠回りをした（2026-08-17）。
//   DriveApp.getFoldersByName は、どれが返るか保証されない。

// 受け口の記録から分かった、控えが実際に入っているフォルダ。
var 固定_控えのフォルダID = '1hPCpks9-GGjmtd8ewP469b-I2YEYOpq3';

function フォルダの下見() {
  var it = DriveApp.getFoldersByName(控え受取_フォルダ名);
  var 一覧 = [];
  while (it.hasNext()) {
    var f = it.next();
    var n = 0;
    var files = f.getFiles();
    while (files.hasNext()) { files.next(); n += 1; }
    一覧.push({ id: f.getId(), 件数: n, url: f.getUrl() });
  }

  Logger.log('■ 「' + 控え受取_フォルダ名 + '」という名前のフォルダ: ' + 一覧.length + '個');
  Logger.log('');
  一覧.forEach(function (x) {
    var 印 = x.id === 固定_控えのフォルダID ? '★ これを使う  ' : '   ';
    Logger.log(印 + x.id + '  ' + x.件数 + '件');
    Logger.log('      ' + x.url);
  });
  Logger.log('');

  var いま = PropertiesService.getScriptProperties().getProperty(控え受取_フォルダIDの鍵);
  Logger.log('■ いま設定されているID: ' + (いま || '（未設定）'));
  Logger.log('');
  if (いま === 固定_控えのフォルダID) {
    Logger.log('  すでに固定されています。');
  } else {
    Logger.log('  よければ フォルダを固定する() を実行してください。');
    Logger.log('  ※ 空のほうのフォルダは、中身が無いことを確かめてから手で消してください。');
  }
}

function フォルダを固定する() {
  // 実在するか・中身があるかを確かめてから書く。
  var folder;
  try {
    folder = DriveApp.getFolderById(固定_控えのフォルダID);
  } catch (e) {
    Logger.log('■ **そのIDのフォルダが見つかりません。**設定は変えていません。');
    return;
  }

  var n = 0;
  var files = folder.getFiles();
  while (files.hasNext()) { files.next(); n += 1; }

  PropertiesService.getScriptProperties().setProperty(控え受取_フォルダIDの鍵, 固定_控えのフォルダID);

  Logger.log('■ 保存先を固定しました');
  Logger.log('  ' + folder.getName() + '（' + n + '件）');
  Logger.log('  ' + folder.getUrl());
  Logger.log('');
  Logger.log('  これ以降、受け口も画面も同じフォルダを見ます。');
}
