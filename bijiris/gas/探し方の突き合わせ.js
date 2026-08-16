// お名前で探した場合と、会員番号で探した場合で、同じ記録に行き着くかを見る道具。
//
//   探し方を突き合わせる()   … 読むだけ。何も書かない。
//
// 記録の手がかりをお名前から会員番号に移すにあたり、切り替えても全員が
// これまでと同じ記録に行き着くことを、先に確かめるために作った。
//
// 見るところは3つ。
//   ・会員番号で引けるか（番号が入っていない方は引けない）
//   ・お名前で引いた記録と、会員番号で引いた記録が同じ方か
//   ・同じ会員番号を2人が持っていないか

function 探し方を突き合わせる() {
  var profiles = getCustomerProfiles_() || {};
  var 名前一覧 = Object.keys(profiles).map(function (k) {
    return String((profiles[k] || {}).name || '');
  }).filter(Boolean).sort();

  var 一致 = [], 番号なし = [], 食い違い = [];
  var 番号ごと = {};

  名前一覧.forEach(function (名) {
    var お名前で = findCustomerProfileByName_(profiles, 名, '');
    var 記録 = お名前で && お名前で.profile ? お名前で.profile : null;
    var 番号 = normalizeMemberNumber_(記録 && 記録.memberNumber);

    if (!番号) { 番号なし.push(名); return; }

    (番号ごと[番号] = 番号ごと[番号] || []).push(名);

    var 番号で = findCustomerProfileByMemberNumber_(profiles, 番号);
    var 行き着く先 = 番号で && 番号で.profile ? String(番号で.profile.name || '') : '';

    if (行き着く先 === 名) {
      一致.push(名 + '（' + 番号 + '）');
    } else {
      食い違い.push(名 + '（' + 番号 + '） → お名前では「' + 名 + '」／番号では「' + (行き着く先 || '見つからない') + '」');
    }
  });

  var 番号の重複 = Object.keys(番号ごと).filter(function (n) { return 番号ごと[n].length > 1; });

  Logger.log('■ 顧客管理の登録: ' + 名前一覧.length + '名');
  Logger.log('');
  Logger.log('■ お名前でも会員番号でも同じ記録: ' + 一致.length + '名');
  一致.forEach(function (x) { Logger.log('    ' + x); });
  Logger.log('');

  if (食い違い.length) {
    Logger.log('■ **食い違う方: ' + 食い違い.length + '名** … 切り替える前に直すこと');
    食い違い.forEach(function (x) { Logger.log('    ' + x); });
    Logger.log('');
  }

  if (番号の重複.length) {
    Logger.log('■ **同じ会員番号を複数の方が持っています** … 切り替える前に直すこと');
    番号の重複.forEach(function (n) { Logger.log('    ' + n + ': ' + 番号ごと[n].join('・')); });
    Logger.log('');
  }

  if (番号なし.length) {
    Logger.log('■ 会員番号が入っていない方: ' + 番号なし.length + '名 … ' + 番号なし.join('・'));
    Logger.log('    この方々は今までどおりお名前で探します。まゆみアプリからお入りになれば番号が入ります。');
    Logger.log('');
  }

  if (!食い違い.length && !番号の重複.length) {
    Logger.log('■ 切り替えて差し支えありません。');
  }
}
