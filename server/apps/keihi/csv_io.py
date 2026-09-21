"""現金出納帳の CSV。列は 日付, 摘要, 相手科目, 収入金額, 支払金額 の 5 つ。

差引残高は計算値なので CSV には入れない。
"""

import csv
import io

HEADER = ["日付", "摘要", "相手科目", "収入金額", "支払金額"]


def write_csv(entries) -> str:
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(HEADER)
    for entry in entries:
        writer.writerow([
            f"{entry.month}/{entry.day}" if entry.day is not None else "",
            entry.description or "",
            entry.counter_account or "",
            entry.income if entry.income is not None else "",
            entry.payment if entry.payment is not None else "",
        ])
    return output.getvalue()


def parse_date(value: str) -> tuple[int | None, int | None]:
    """「9/3」「2026/9/3」「09-03」などから月と日を取り出す。"""
    parts = [p for p in value.replace("-", "/").split("/") if p.strip()]
    try:
        numbers = [int(p) for p in parts]
    except ValueError:
        return None, None
    if len(numbers) >= 3:
        return numbers[1], numbers[2]
    if len(numbers) == 2:
        return numbers[0], numbers[1]
    return None, None


def parse_amount(value: str) -> int | None:
    """「1,200」「¥1200」「1200円」「１２００」などから整数を取り出す。"""
    cleaned = value.replace(",", "").replace("¥", "").replace("円", "").strip()
    cleaned = cleaned.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    if not cleaned:
        return None
    try:
        return int(float(cleaned))
    except ValueError:
        return None


def decode(raw: bytes) -> str:
    """Excel が書いた Shift_JIS も読めるようにする。"""
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw.decode("cp932", errors="replace")


def read_csv(content: str) -> list[dict]:
    """CSV から行の中身を取り出す。ヘッダー行があってもなくても読める。"""
    reader = csv.reader(io.StringIO(content.lstrip("﻿")))
    rows = [row for row in reader if any(cell.strip() for cell in row)]
    if rows and rows[0][: len(HEADER)] == HEADER:
        rows = rows[1:]

    entries = []
    for row in rows:
        cells = (row + [""] * len(HEADER))[: len(HEADER)]
        month, day = parse_date(cells[0])
        entries.append({
            "month": month,
            "day": day,
            "description": cells[1].strip(),
            "counter_account": cells[2].strip(),
            "income": parse_amount(cells[3]),
            "payment": parse_amount(cells[4]),
        })
    return entries
