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

  /** 前葉繰越から順に足し引きして、残高と合計を出し直す。 */
  function recalc() {
    const opening = toNumber(document.getElementById('opening').value);
    let balance = opening;
    let incomeTotal = 0;
    let paymentTotal = 0;

    table.querySelectorAll('tbody tr').forEach(function (tr) {
      const income = toNumber(tr.querySelector('[data-col="4"]').value);
      const payment = toNumber(tr.querySelector('[data-col="5"]').value);
      incomeTotal += income;
      paymentTotal += payment;
      balance += income - payment;

      // 何も書かれていない行は、紙と同じく残高も空欄のままにする
      const written = ['1', '2', '3', '4', '5'].some(function (col) {
        return (tr.querySelector('[data-col="' + col + '"]').value || '').trim() !== '';
      });
      tr.querySelector('.balance').textContent = written ? format(balance) : '';
    });

    document.getElementById('income-total').textContent = format(incomeTotal);
    document.getElementById('payment-total').textContent = format(paymentTotal);
    document.getElementById('closing-balance').textContent = format(opening + incomeTotal - paymentTotal);
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
  document.getElementById('opening').addEventListener('input', recalc);

  recalc();
})();
