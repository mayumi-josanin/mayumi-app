// サイドバーの開閉（スマートフォン）。KEM_DDENKI の app.js と同じ動き。
document.addEventListener('DOMContentLoaded', function () {
  var hamburger = document.querySelector('.hamburger');
  var sidebar = document.querySelector('.sidebar');
  if (hamburger && sidebar) {
    hamburger.addEventListener('click', function () {
      sidebar.classList.toggle('open');
    });
    document.addEventListener('click', function (e) {
      if (window.innerWidth <= 768 && !sidebar.contains(e.target) && !hamburger.contains(e.target)) {
        sidebar.classList.remove('open');
      }
    });
  }
});

// 左メニューの段（アプリ管理・公式サイト・予約管理）は、見出しをタップすると一覧が開く。
// 院長の希望（2026-09-16）: 段ごとに閉じておき、必要な段だけ開く。
// いま開いている画面が入っている段だけは、最初から開いておく（自分がどこにいるか分かるように）。
// 開け閉めは端末ごとに覚える（localStorage）。
document.addEventListener('DOMContentLoaded', function () {
  var sections = document.querySelectorAll('.sidebar-nav .nav-section');
  sections.forEach(function (sec) {
    var links = [];
    var el = sec.nextElementSibling;
    while (el && !el.classList.contains('nav-section')) { links.push(el); el = el.nextElementSibling; }
    if (!links.length) return;
    var key = 'nav-open:' + sec.textContent.trim();
    var hasActive = links.some(function (a) { return a.classList.contains('active'); });
    var saved = null;
    try { saved = localStorage.getItem(key); } catch (e) {}
    var open = hasActive ? true : (saved === '1');
    function apply() {
      sec.classList.toggle('collapsed', !open);
      links.forEach(function (a) { a.hidden = !open; });
    }
    sec.setAttribute('role', 'button');
    sec.setAttribute('tabindex', '0');
    sec.addEventListener('click', function () {
      open = !open; apply();
      try { localStorage.setItem(key, open ? '1' : '0'); } catch (e) {}
    });
    sec.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); sec.click(); } });
    apply();
  });
});
