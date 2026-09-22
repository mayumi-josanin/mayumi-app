/* 現金出納帳の入力まわり。
 *
 *  - Enter で右のセル、行の最後なら次の行の先頭へ（Tab と同じ並び）
 *  - 上下の矢印で同じ列の行移動
 *  - 金額は入力中は素の数字、セルを離れると 3 桁カンマ
 *  - 差引残高と合計は、入力しているそばから計算して出す
 *    （保存すればサーバー側でも同じ計算をする。正は常にサーバー側）
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

  table.addEventListener('input', recalc);
  // 前葉繰越の欄は外した（2026-09-22）。あっても無くても落ちないようにしておく
  const openingEl = document.getElementById('opening');
  if (openingEl) openingEl.addEventListener('input', recalc);

  recalc();
})();
