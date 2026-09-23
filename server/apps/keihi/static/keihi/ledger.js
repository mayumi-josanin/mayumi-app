/* 現金出納帳の入力まわり。
 *
 *  - Enter で右のセル、行の最後なら次の行の先頭へ（Tab と同じ並び）
 *  - 上下の矢印で同じ列の行移動
 *  - 金額は入力中は素の数字、セルを離れると 3 桁カンマ
 *  - 合計は、入力しているそばから計算して出す
 *    （保存すればサーバー側でも同じ計算をする。正は常にサーバー側）
 *  - 摘要が長い行は、枠に収まるまで文字を小さくする
 */
(function () {
  const table = document.querySelector('.ledger');
  if (!table) return;

  const COLUMN_COUNT = 6; // 月・日・摘要・相手科目・収入金額・支払金額

  const cell = (row, col) =>
    table.querySelector('[data-row="' + row + '"][data-col="' + col + '"]');

  const lastRow = () => {
    const cells = table.querySelectorAll('[data-col="0"]');
    return cells.length - 1;
  };

  function focusCell(row, col) {
    const el = cell(row, col);
    if (el) { el.focus(); el.select(); }
  }

  function toNumber(value) {
    const cleaned = (value || '').replace(/,/g, '').replace(/[^0-9-]/g, '');
    return cleaned === '' ? 0 : parseInt(cleaned, 10);
  }

  function format(n) {
    return n.toLocaleString('ja-JP');
  }

  /** 収入と支払の合計を出し直す。
   *
   * **差引残高は出さない**（院長の希望 2026-09-22「差引残高を削除」）。
   * 前葉繰越の欄も画面から外したので、ここでも読まない。
   */
  function recalc() {
    let incomeTotal = 0;
    let paymentTotal = 0;

    table.querySelectorAll('tbody tr').forEach(function (tr) {
      incomeTotal += toNumber(tr.querySelector('[data-col="4"]').value);
      paymentTotal += toNumber(tr.querySelector('[data-col="5"]').value);
    });

    document.getElementById('income-total').textContent = format(incomeTotal);
    document.getElementById('payment-total').textContent = format(paymentTotal);
  }

  /* ------------------------------------------------------------------
     摘要を枠の中に収める（院長の希望 2026-09-22）
     ------------------------------------------------------------------ */

  // これ以上は小さくしない。読めなくなっては、収まっても意味がない。
  // この大きさで、全角なら 31 文字ほどが 1 行に入る（56mm の列で確かめた）
  const 摘要の下限 = 7;

  /** 摘要の欄が枠からはみ出していたら、収まるまで文字を少し小さくする。
   *
   * 列の幅は紙の様式（様式 777）で決まっているので広げられない。入力欄は、
   * はみ出した分が横に隠れるだけなので、**画面でも紙でも読めなくなる。**
   * そこで、はみ出した行だけ文字の方を縮める。はみ出していない行は触らない
   * （全部を小さくすると、短い摘要まで読みにくくなる）。
   *
   * 下限まで小さくしても入らないほど長い摘要は、そこで止める。
   */
  function 摘要を収める(el) {
    if (!el) return;
    el.style.fontSize = '';                       // いったん元に戻してから測り直す
    if (el.scrollWidth <= el.clientWidth) return; // 収まっているなら触らない

    // 元の大きさは ledger.css が決めている。ここには書き写さない（2か所になると必ずずれる）
    let size = parseFloat(getComputedStyle(el).fontSize);
    while (size > 摘要の下限 && el.scrollWidth > el.clientWidth) {
      size = Math.max(摘要の下限, size - 0.25);
      el.style.fontSize = size + 'px';
    }
  }

  const 摘要たち = () => table.querySelectorAll('[data-col="2"]');

  function 摘要を全部収める() {
    摘要たち().forEach(摘要を収める);
  }

  table.addEventListener('keydown', function (e) {
    const el = e.target;
    if (!el.dataset || el.dataset.col === undefined) return;
    const row = parseInt(el.dataset.row, 10);
    const col = parseInt(el.dataset.col, 10);

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      focusCell(Math.min(row + 1, lastRow()), col);
      return;
    }
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      focusCell(Math.max(row - 1, 0), col);
      return;
    }
    if (e.key !== 'Enter') return;

    // Enter は Tab と同じく次のセルへ。うっかりフォームが送られないように止める。
    e.preventDefault();
    if (col < COLUMN_COUNT - 1) {
      focusCell(row, col + 1);
    } else if (row < lastRow()) {
      focusCell(row + 1, 0);
    } else {
      // 最終行の最後まで来たら、行を足す（今の入力も一緒に保存される）
      document.querySelector('[name="add_row"]').click();
    }
  });

  // 金額のセル: 触っているあいだは素の数字、離れたらカンマ付き
  table.addEventListener('focusin', function (e) {
    const el = e.target;
    if (!el.classList || !el.classList.contains('ledger-amount')) return;
    el.value = (el.value || '').replace(/,/g, '');
    el.select();
  });

  table.addEventListener('focusout', function (e) {
    const el = e.target;
    if (!el.classList || !el.classList.contains('ledger-amount')) return;
    const digits = (el.value || '').replace(/[^0-9]/g, '');
    el.value = digits === '' ? '' : format(parseInt(digits, 10));
    recalc();
  });

  table.addEventListener('input', function (e) {
    recalc();
    // 書いているそばから収める。書き終わってからでは、途中が見えないまま進む
    if (e.target && e.target.dataset && e.target.dataset.col === '2') 摘要を収める(e.target);
  });
  // 前葉繰越の欄は外した（2026-09-22）。あっても無くても落ちないようにしておく
  const openingEl = document.getElementById('opening');
  if (openingEl) openingEl.addEventListener('input', recalc);

  recalc();
  摘要を全部収める();

  // 印刷の直前にも測り直す。**2通りで見張る。**iPad の Safari は beforeprint を出さず、
  // 代わりに print のメディアが切り替わる。どちらか片方だけだと、紙にだけはみ出しが残る
  if (window.matchMedia) {
    const 紙 = window.matchMedia('print');
    if (紙.addEventListener) 紙.addEventListener('change', 摘要を全部収める);
  }
  window.addEventListener('beforeprint', 摘要を全部収める);
})();
