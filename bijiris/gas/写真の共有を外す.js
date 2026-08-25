// お客様の計測写真の「リンクを知っている全員が閲覧可」を外す道具。
//
//   写真の共有の下見()  … 何件が公開になっているかを数えるだけ（**変えない**）
//   写真の共有を外す()  … 実際に外す
//
// ## 実行する順番を間違えないこと
//
// **アプリ側の更新を先に公開してから、これを実行する。**
//
//   1. お客様アプリ・管理アプリを新しい版で公開する
//      （写真を `?action=photoData` 経由で出す作りになっている）
//   2. 自分の端末で、写真が出ることを確かめる
//   3. **そのあとで**この道具を実行する
//
// 順番を逆にすると、**アプリが古い間は写真が出なくなる。**
//
// ## 外さないもの
//
// ビジリス投稿の画像・PDF・表紙は**公開のままにする。**
// お客様みなさまに見ていただくためのもので、性質が違う。
//
//   Bijiris/投稿/  … 触らない
//   Bijiris/計測時/・測定/ … ここのお客様の写真を外す
//
// ## 途中で止まっても大丈夫
//
// Drive は件数が多いと時間切れになる。**何度実行してもよい作り**にしてある
// （すでに外れているものは飛ばす）。止まったら、もう一度実行する。

// **除外するフォルダだけを名指しする。**残りは全部たどる。
//
// 「触るフォルダを並べる」やり方にすると、知らないフォルダを取りこぼす。
// 守りたいのはお客様の写真なので、**取りこぼすより、除外を明示する。**
//
//   ビジリス通信 … お客様みなさまに見ていただく投稿。**公開のままにする**
//   分析シート   … 画像ではない（触っても害はないが、素通りさせる）
var 共有外し_除外 = ['ビジリス通信', '分析シート'];

function 写真の共有の下見() {
  var r = 共有外し_数える_(false);
  共有外し_出す_(r, false);
}

function 写真の共有を外す() {
  var r = 共有外し_数える_(true);
  共有外し_出す_(r, true);
}

function 共有外し_数える_(外すか) {
  // Bijiris フォルダ。Code.gs が持っている取得口を使う。
  var 根 = getRootPhotoFolder_();
  var 結果 = { 見た: 0, 公開だった: 0, 外した: 0, 失敗: 0, フォルダ: {}, 時間切れ: false };
  var 始め = new Date().getTime();

  // 根の直下にあるファイルも見る。
  結果.フォルダ['（Bijiris直下）'] = 共有外し_ファイルだけ_(根, 外すか, 結果, 始め) + '件';

  var subs = 根.getFolders();
  while (subs.hasNext()) {
    var f = subs.next();
    var 名 = f.getName();
    if (共有外し_除外.indexOf(名) >= 0) {
      結果.フォルダ[名] = '**除外（公開のまま）**';
      continue;
    }
    結果.フォルダ[名] = 共有外し_歩く_(f, 外すか, 結果, 始め) + '件';
  }

  return 結果;
}

// そのフォルダ直下のファイルだけを見る（子フォルダはたどらない）。
function 共有外し_ファイルだけ_(folder, 外すか, 結果, 始め) {
  var 数 = 0;
  var files = folder.getFiles();
  while (files.hasNext()) {
    if (new Date().getTime() - 始め > 5 * 60 * 1000) { 結果.時間切れ = true; return 数; }
    数 += 共有外し_1件_(files.next(), 外すか, 結果) ? 1 : 1;
  }
  return 数;
}

// 1件を見て、必要なら外す。見たら true。
function 共有外し_1件_(f, 外すか, 結果) {
  結果.見た++;
  var 公開 = false;
  try {
    公開 = f.getSharingAccess() === DriveApp.Access.ANYONE_WITH_LINK;
  } catch (e) {
    結果.失敗++;
    return true;
  }
  if (!公開) return true;
  結果.公開だった++;
  if (!外すか) return true;
  try {
    // **持ち主だけが見られる状態に戻す。**
    f.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    結果.外した++;
  } catch (e) {
    結果.失敗++;
  }
  return true;
}

function 共有外し_歩く_(folder, 外すか, 結果, 始め) {
  // 5分で切り上げる。GASは6分で強制終了され、そこまでの記録が残らない。
  if (new Date().getTime() - 始め > 5 * 60 * 1000) {
    結果.時間切れ = true;
    return 0;
  }

  var 数 = 0;
  var files = folder.getFiles();
  while (files.hasNext()) {
    if (new Date().getTime() - 始め > 5 * 60 * 1000) { 結果.時間切れ = true; return 数; }
    共有外し_1件_(files.next(), 外すか, 結果);
    数++;
  }

  var subs = folder.getFolders();
  while (subs.hasNext()) {
    数 += 共有外し_歩く_(subs.next(), 外すか, 結果, 始め);
  }
  return 数;
}

function 共有外し_出す_(r, 外したか) {
  Logger.log('■ お客様の計測写真の共有ぐあい');
  Logger.log('');
  Object.keys(r.フォルダ).forEach(function (k) {
    Logger.log('   ' + k + ': ' + r.フォルダ[k]);
  });
  Logger.log('');
  Logger.log('   見たファイル:        ' + r.見た + '件');
  Logger.log('   **公開になっていた: ' + r.公開だった + '件**');
  if (外したか) {
    Logger.log('   **外した:           ' + r.外した + '件**');
  }
  if (r.失敗) {
    Logger.log('   触れなかった:        ' + r.失敗 + '件（持ち主が違う可能性があります）');
  }
  Logger.log('');

  if (r.時間切れ) {
    Logger.log('■ **5分で切り上げました。まだ残っています。**');
    Logger.log('   もう一度実行してください。すでに外したものは飛ばします。');
    Logger.log('');
  }

  if (!外したか) {
    Logger.log('■ **下見なので、何も変えていません。**');
    Logger.log('');
    Logger.log('   ※ 外す前に、**新しい版のアプリを公開して、写真が出ることを**');
    Logger.log('      **確かめてください。**順番を逆にすると写真が出なくなります。');
  } else if (r.公開だった === 0) {
    Logger.log('■ 公開のものはありませんでした。');
  } else {
    Logger.log('■ 外し終わりました。アプリで写真が出ることを確かめてください。');
  }
}
