"""現金出納帳の計算。

差引残高と合計はどこにも保存せず、常にここで計算する。
紙の様式（JDL 様式 777）に合わせて、1ページは 25 行を基本とする。
"""

# 紙の様式 1 ページ分の行数。帳簿を開いたときはこの数だけ空行を用意する。
DEFAULT_ROWS = 25

# 相手科目のサジェスト。自由入力もできるので、あくまで候補。
SUGGESTED_ACCOUNTS = [
    "売上高", "普通預金", "当座預金", "現金", "売掛金", "買掛金", "仕入高",
    "消耗品費", "事務用品費", "水道光熱費", "通信費", "旅費交通費", "荷造運賃",
    "広告宣伝費", "接待交際費", "会議費", "新聞図書費", "修繕費", "地代家賃",
    "支払手数料", "租税公課", "法定福利費", "福利厚生費", "給料手当", "雑費",
    "前払金", "仮払金", "事業主貸", "事業主借",
]


def is_blank(entry) -> bool:
    """まだ何も書かれていない行か。

    月は帳簿から自動で入るので、書かれたかどうかの判断には使わない。
    """
    return not any([entry.day, entry.description, entry.counter_account, entry.income, entry.payment])


def entry_errors(entry) -> list[str]:
    """入力の不備を返す。保存は拒否せず、画面で知らせるだけに使う。

    書きかけで手が止まらないよう、止めるのではなく知らせるだけにしている。
    """
    if is_blank(entry):
        return []

    errors = []
    if entry.day is None:
        errors.append("日付が入っていません")
    if entry.income and entry.payment:
        errors.append("収入金額と支払金額の両方に入っています")
    if not entry.income and not entry.payment:
        errors.append("収入金額と支払金額のどちらも入っていません")
    if (entry.income or 0) < 0 or (entry.payment or 0) < 0:
        errors.append("金額がマイナスです")
    return errors


def running_balances(opening_balance: int, entries) -> list[int]:
    """各行の差引残高を、渡された順に計算して返す。1 行目は前葉繰越が起点。"""
    balances = []
    balance = opening_balance
    for entry in entries:
        balance += (entry.income or 0) - (entry.payment or 0)
        balances.append(balance)
    return balances


def totals(opening_balance: int, entries) -> dict[str, int]:
    """合計行の内容（収入合計・支払合計・最終残高）を返す。"""
    income_total = sum(e.income or 0 for e in entries)
    payment_total = sum(e.payment or 0 for e in entries)
    return {
        "income_total": income_total,
        "payment_total": payment_total,
        "closing_balance": opening_balance + income_total - payment_total,
    }


def sorted_by_date(entries):
    """日付順に並べ替える。日付が未入力の行は、元の順のまま末尾に送る。"""
    dated = [e for e in entries if e.day is not None]
    undated = [e for e in entries if e.day is None]
    dated.sort(key=lambda e: (e.month or 0, e.day or 0))
    return dated + undated


def rows_for_display(book, entries):
    """画面に渡す 1 行分の束。残高と不備をここで付ける。"""
    balances = running_balances(book.opening_balance, entries)
    return [
        {
            "entry": entry,
            "no": index + 1,
            # 紙と同じく 5 行ごとに目盛りの数字を出す
            "tick": (index + 1) if (index + 1) % 5 == 0 else "",
            "balance": "" if is_blank(entry) else balances[index],
            "errors": entry_errors(entry),
        }
        for index, entry in enumerate(entries)
    ]
