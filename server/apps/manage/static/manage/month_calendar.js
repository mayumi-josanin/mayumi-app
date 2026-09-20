/* 管理画面の「月の表」。一覧（カレンダー管理）と追加の画面の両方で使う。
   同じ見た目・同じ考え方を2か所に書くと片方だけ直して食い違うので、ここに1つだけ置く。 */
(function (global) {
  'use strict';

  var WEEK = ['日', '月', '火', '水', '木', '金', '土'];
  // 同じ日に複数あるときの印の並び順（旧アプリと同じ）。
  var ORDER = { event: 0, holiday: 1, visit: 2, postpartum: 3 };

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function icon(kind) {
    if (kind === 'holiday') return '<span class="cal-icon-badge icon-holiday" title="休診">休</span>';
    if (kind === 'visit') return '<span class="cal-icon-badge icon-visit" title="往診">往診</span>';
    if (kind === 'postpartum') return '<span class="cal-icon-badge icon-postpartum" title="訪問産後ケア">訪問</span>';
    return '<span class="cal-icon-badge icon-event" title="イベント">イ</span>';
  }

  function typeLabel(kind) {
    return kind === 'holiday' ? '休診' : (kind === 'visit' ? '往診' : (kind === 'postpartum' ? '訪問産後ケア' : 'イベント'));
  }

  function pad2(n) { return String(n).padStart(2, '0'); }

  function dateStr(year, month, day) { return year + '-' + pad2(month) + '-' + pad2(day); }

  /* 月を1つ動かす（前月・翌月）。year/month を持つ入れ物を返す。 */
  function shift(year, month, delta) {
    var y = year, m = month + delta;
    while (m < 1) { m += 12; y -= 1; }
    while (m > 12) { m -= 12; y += 1; }
    return { year: y, month: m };
  }

  /* 表示中の月の、その曜日（0=日）の日を全部返す。 */
  function daysOfWeekday(year, month, dow) {
    var out = [], last = new Date(year, month, 0).getDate();
    for (var d = 1; d <= last; d++) {
      if (new Date(year, month - 1, d).getDay() === dow) out.push(dateStr(year, month, d));
    }
    return out;
  }

  /* 月の表を組み立てて入れる。
     grid : 入れ先の要素
     opts : {
       year, month        表示する年月
       events             [{date,title,color,kind,...}]（省略可）
       today              'YYYY-MM-DD'（今日の枠。省略可）
       selected           ['YYYY-MM-DD']（選んだ印。省略可）
       showChips          予定の名前を出すか（既定 true）
       onDay(dateStr)     日を押したとき（渡せば全部の日が押せる）
     } */
  function render(grid, opts) {
    if (!grid) return;
    var year = opts.year, month = opts.month;
    var monthStr = pad2(month);
    var selected = opts.selected || [];
    var showChips = opts.showChips !== false;
    var byDay = {};
    (opts.events || []).forEach(function (e) {
      if (!e.date) return;
      var p = String(e.date).split('-');
      if (p[0] !== String(year) || p[1] !== monthStr) return;
      var d = parseInt(p[2], 10);
      (byDay[d] = byDay[d] || []).push(e);
    });

    var first = new Date(year, month - 1, 1);
    var daysInMonth = new Date(year, month, 0).getDate();
    var startDow = first.getDay();
    var prevLast = new Date(year, month - 1, 0).getDate();
    var html = '';
    WEEK.forEach(function (w, i) {
      html += '<div class="month-calendar-weekday ' + (i === 0 ? 'sun' : (i === 6 ? 'sat' : '')) + '">' + w + '</div>';
    });
    for (var i = 0; i < startDow; i++) {
      html += '<div class="month-calendar-day is-other-month"><span class="day-number">' + (prevLast - startDow + i + 1) + '</span></div>';
    }
    for (var day = 1; day <= daysInMonth; day++) {
      var dow = (startDow + day - 1) % 7;
      var ds = dateStr(year, month, day);
      var evs = byDay[day] || [];
      var kinds = [];
      evs.forEach(function (e) { if (kinds.indexOf(e.kind) === -1) kinds.push(e.kind); });
      kinds.sort(function (a, b) { return ORDER[a] - ORDER[b]; });
      var chips = '';
      if (showChips) {
        evs.slice(0, 2).forEach(function (e) {
          chips += '<div class="day-event-chip" style="border-left-color:' + esc(e.color) + '" title="' + esc(e.title) + '">' + esc(e.title) + '</div>';
        });
        if (evs.length > 2) chips += '<div class="day-event-more">+' + (evs.length - 2) + '件</div>';
      }
      var cls = 'month-calendar-day';
      if (ds === opts.today) cls += ' is-today';
      if (evs.length) cls += ' has-events';
      if (selected.indexOf(ds) !== -1) cls += ' is-picked';
      if (opts.onDay) cls += ' is-clickable';
      html += '<div class="' + cls + '" data-date="' + ds + '"' +
        (opts.onDay ? ' role="button" tabindex="0" aria-pressed="' + (selected.indexOf(ds) !== -1) + '"' : '') + '>';
      html += '<span class="day-number ' + (dow === 0 ? 'sun' : (dow === 6 ? 'sat' : '')) + '">' + day + '</span>';
      if (kinds.length) html += '<div class="day-icons">' + kinds.map(icon).join('') + '</div>';
      if (chips) html += '<div class="day-events">' + chips + '</div>';
      html += '</div>';
    }
    var trailing = (7 - ((startDow + daysInMonth) % 7)) % 7;
    for (var t = 1; t <= trailing; t++) {
      html += '<div class="month-calendar-day is-other-month"><span class="day-number">' + t + '</span></div>';
    }
    grid.innerHTML = html;

    if (opts.onDay) {
      grid.querySelectorAll('.month-calendar-day[data-date]').forEach(function (cell) {
        cell.addEventListener('click', function () { opts.onDay(cell.dataset.date); });
        // キーボードでも押せるように（Enter / スペース）。
        cell.addEventListener('keydown', function (ev) {
          if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); opts.onDay(cell.dataset.date); }
        });
      });
    }
  }

  global.MonthCalendar = {
    render: render, esc: esc, icon: icon, typeLabel: typeLabel,
    shift: shift, daysOfWeekday: daysOfWeekday, dateStr: dateStr, WEEK: WEEK
  };
})(window);
