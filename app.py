# -*- coding: utf-8 -*-
"""
Bank/PAYA Reconciliation - Windows Desktop Application
No virtual environment required.

Install once:
    py -m pip install -r requirements.txt

Run:
    run.bat

Build EXE:
    build_exe.bat

Features:
- Windows GUI with two tabs: Reconciliation / AI Settings
- Import Bank Excel + PAYA Excel
- Automatic sheet/header/column detection
- Sends ALL rows from BOTH workbooks to the configured OpenAI-compatible API,
  in numbered chunks, so the complete source data is transmitted.
- Local deterministic reconciliation is the final source of truth.
- AI is used as a secondary analytical layer and can explain ambiguous matches.
- Output Excel contains summary, matched, bank-only, PAYA-only, amount mismatches,
  duplicates, daily balances, AI analysis and source data.
- API settings are saved locally in config.json.
"""

import json
import os
import re
import threading
import traceback
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import requests
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


APP_DIR = Path(__file__).resolve().parent
CONFIG_FILE = APP_DIR / "config.json"
OUTPUT_DIR = APP_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

DEFAULT_CONFIG = {
    "base_url": "https://api.gapgpt.app/v1",
    "api_key": "",
    "model": "gpt-4o",
    "temperature": 0.0,
    "max_tokens": 4096,
    "timeout": 120,
    "send_all_rows": True,
    "ai_enabled": True,
    "chunk_rows": 150,
}

PERSIAN_DIGITS = str.maketrans(
    "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩",
    "01234567890123456789"
)

DATE_WORDS = [
    "date", "transaction date", "posting date", "value date", "operation date",
    "تاریخ", "تاریخ تراکنش", "تاریخ عملیات", "تاریخ ثبت", "تاریخ موثر",
    "تاريخ", "تاريخ تراکنش", "تاريخ عمليات"
]
AMOUNT_WORDS = [
    "amount", "transaction amount", "operation amount", "مبلغ",
    "مبلغ تراکنش", "مبلغ عملیات", "مبلغ ریالی", "مبلغ ریال",
    "amount rial", "transaction value", "value"
]
DEBIT_WORDS = [
    "debit", "debit amount", "بدهکار", "بدهکاری", "مبلغ بدهکار",
    "برداشت", "برداشتی", "مبلغ برداشت", "برداشت ریالی", "برداشت ریال",
    "برداشت از حساب", "خروج"
]
CREDIT_WORDS = [
    "credit", "credit amount", "بستانکار", "بستانکاری", "مبلغ بستانکار",
    "واریز", "واریزی", "مبلغ واریز", "واریز ریالی", "واریز ریال",
    "واریز به حساب", "ورود"
]
BALANCE_WORDS = [
    "balance", "account balance", "closing balance", "ending balance",
    "مانده", "مانده حساب", "مانده نهایی", "مانده پایان", "مانده جاری",
    "موجودی", "موجودی حساب"
]
ID_WORDS = [
    "reference", "ref", "trace", "tracking", "tracking number",
    "reference number", "transaction id", "transaction number",
    "شماره پیگیری", "شماره رهگیری", "شناسه", "شناسه تراکنش",
    "شماره سند", "سند", "شماره تراکنش", "کد رهگیری", "رفرنس"
]
DESC_WORDS = [
    "description", "desc", "details", "narration", "شرح", "شرح تراکنش",
    "شرح عملیات", "توضیحات", "بابت", "شرح واریز", "شرح برداشت"
]
TYPE_WORDS = [
    "type", "transaction type", "operation type", "نوع", "نوع تراکنش",
    "نوع عملیات", "ماهیت", "وضعیت تراکنش"
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
SUB_FILL = PatternFill("solid", fgColor="D9EAF7")
OK_FILL = PatternFill("solid", fgColor="C6EFCE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WHITE_FONT = Font(color="FFFFFF", bold=True)
BOLD_FONT = Font(bold=True)
THIN = Side(style="thin", color="D9D9D9")


def load_config():
    if not CONFIG_FILE.exists():
        save_config(DEFAULT_CONFIG)
        return dict(DEFAULT_CONFIG)
    try:
        data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        cfg = dict(DEFAULT_CONFIG)
        cfg.update(data)
        return cfg
    except Exception:
        return dict(DEFAULT_CONFIG)


def save_config(cfg):
    CONFIG_FILE.write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def norm_text(value):
    if value is None:
        return ""
    s = str(value).strip().translate(PERSIAN_DIGITS)
    s = re.sub(r"\s+", " ", s)
    return s


def norm_header(value):
    s = norm_text(value).lower()
    s = s.replace("_", " ").replace("-", " ")
    return re.sub(r"\s+", " ", s).strip()


def parse_decimal(value):
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        try:
            return Decimal(str(value))
        except Exception:
            return None
    s = norm_text(value)
    if not s:
        return None
    s = s.replace(",", "").replace("٬", "").replace(" ", "")
    s = s.replace("ریال", "").replace("تومان", "")
    s = s.replace("٫", ".")
    s = re.sub(r"[^0-9.\-+]", "", s)
    if not s or s in ("-", "+", "."):
        return None
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def serializable_value(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.isoformat(sep=" ")
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, float):
        return round(v, 10)
    return str(v)


def parse_date(value):
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    s = norm_text(value)
    if not s:
        return ""
    # Keep Jalali dates as text; normalize separators.
    s = s.replace("/", "-").replace(".", "-")
    # Gregorian common formats.
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return s


def choose_sheet(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    best = None
    best_score = -1
    for ws in wb.worksheets:
        rows = ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 15), values_only=True)
        sample = list(rows)
        if not sample:
            continue
        nonempty = sum(1 for r in sample for v in r if v not in (None, ""))
        score = nonempty + min(ws.max_row or 0, 100) * 0.01
        if score > best_score:
            best_score = score
            best = ws.title
    wb.close()
    return best


def detect_header_row(path, sheet_name):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name]
    max_cols = min(ws.max_column or 1, 100)
    best_row = 1
    best_score = -1
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 20),
                     max_col=max_cols, values_only=True), start=1
    ):
        vals = [norm_header(v) for v in row]
        score = 0
        joined = " | ".join(vals)
        for words, weight in [
            (DATE_WORDS, 4), (AMOUNT_WORDS, 4), (BALANCE_WORDS, 3),
            (ID_WORDS, 3), (DESC_WORDS, 2), (DEBIT_WORDS, 2), (CREDIT_WORDS, 2)
        ]:
            if any(w in joined for w in words):
                score += weight
        score += sum(1 for v in vals if v) * 0.1
        if score > best_score:
            best_score = score
            best_row = row_idx
    wb.close()
    return best_row


def read_excel(path):
    sheet = choose_sheet(path)
    header_row = detect_header_row(path, sheet)
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]

    raw_headers = []
    for c in range(1, (ws.max_column or 1) + 1):
        v = ws.cell(header_row, c).value
        raw_headers.append(norm_text(v) or f"Column_{c}")

    # Ensure unique headers.
    headers = []
    seen = {}
    for h in raw_headers:
        n = seen.get(h, 0)
        seen[h] = n + 1
        headers.append(h if n == 0 else f"{h}_{n+1}")

    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not any(v not in (None, "") for v in row):
            continue
        d = {}
        for i, h in enumerate(headers):
            d[h] = serializable_value(row[i] if i < len(row) else None)
        rows.append(d)

    wb.close()
    return {
        "path": str(path),
        "sheet": sheet,
        "header_row": header_row,
        "headers": headers,
        "rows": rows,
    }


def best_column(headers, words):
    scored = []
    for h in headers:
        nh = norm_header(h)
        score = 0
        for w in words:
            nw = norm_header(w)
            if nh == nw:
                score += 100
            elif nw and nw in nh:
                score += 40
        if score:
            scored.append((score, h))
    return max(scored, default=(0, None))[1]


def infer_columns(data, overrides=None):
    headers = data["headers"]
    overrides = overrides or {}
    result = {
        "date": best_column(headers, DATE_WORDS),
        "amount": best_column(headers, AMOUNT_WORDS),
        "debit": best_column(headers, DEBIT_WORDS),
        "credit": best_column(headers, CREDIT_WORDS),
        "balance": best_column(headers, BALANCE_WORDS),
        "id": best_column(headers, ID_WORDS),
        "description": best_column(headers, DESC_WORDS),
        "type": best_column(headers, TYPE_WORDS),
    }
    for k, v in overrides.items():
        if v in headers or v == "":
            result[k] = v or None

    # If the workbook has debit/credit (or withdrawal/deposit) columns,
    # do not force a generic "amount" column.
    if not result["amount"] and (result["debit"] or result["credit"]):
        result["amount"] = None
    return result


def row_amount(row, cols):
    # Explicit amount wins.
    if cols.get("amount"):
        a = parse_decimal(row.get(cols["amount"]))
        if a is not None:
            return abs(a)
    debit = parse_decimal(row.get(cols.get("debit"))) if cols.get("debit") else None
    credit = parse_decimal(row.get(cols.get("credit"))) if cols.get("credit") else None
    if debit is not None and debit != 0:
        return abs(debit)
    if credit is not None and credit != 0:
        return abs(credit)
    return None


def row_direction(row, cols):
    debit = parse_decimal(row.get(cols.get("debit"))) if cols.get("debit") else None
    credit = parse_decimal(row.get(cols.get("credit"))) if cols.get("credit") else None
    if debit is not None and debit != 0:
        return "DEBIT"
    if credit is not None and credit != 0:
        return "CREDIT"

    type_value = norm_text(row.get(cols.get("type"), "")).lower() if cols.get("type") else ""
    desc = norm_text(row.get(cols.get("description"), "")).lower() if cols.get("description") else ""
    text = type_value + " " + desc

    debit_terms = ["برداشت", "برداشتی", "بدهکار", "debit", "withdrawal", "withdraw", "خروج"]
    credit_terms = ["واریز", "واریزی", "بستانکار", "credit", "deposit", "deposited", "ورود"]

    if any(x in text for x in debit_terms):
        return "DEBIT"
    if any(x in text for x in credit_terms):
        return "CREDIT"
    return ""


def make_tx(row, idx, cols, source):
    return {
        "index": idx,
        "source": source,
        "date": parse_date(row.get(cols.get("date"))) if cols.get("date") else "",
        "amount": row_amount(row, cols),
        "direction": row_direction(row, cols),
        "id": norm_text(row.get(cols.get("id"))) if cols.get("id") else "",
        "description": norm_text(row.get(cols.get("description"))) if cols.get("description") else "",
        "balance": parse_decimal(row.get(cols.get("balance"))) if cols.get("balance") else None,
        "raw": row,
    }


def same_amount(a, b, tolerance=Decimal("0")):
    if a is None or b is None:
        return False
    return abs(a - b) <= tolerance


def similarity_score(bank, paya):
    score = 0
    if bank["date"] and paya["date"]:
        if bank["date"] == paya["date"]:
            score += 50
    if bank["id"] and paya["id"]:
        if bank["id"] == paya["id"]:
            score += 100
        elif bank["id"] in paya["id"] or paya["id"] in bank["id"]:
            score += 30
    if bank["amount"] is not None and paya["amount"] is not None:
        if bank["amount"] == paya["amount"]:
            score += 50
        else:
            diff = abs(bank["amount"] - paya["amount"])
            if diff <= max(bank["amount"], paya["amount"]) * Decimal("0.01"):
                score += 15
    if bank["direction"] and paya["direction"] and bank["direction"] == paya["direction"]:
        score += 20
    if bank["description"] and paya["description"]:
        bwords = set(re.findall(r"\w+", bank["description"].lower()))
        pwords = set(re.findall(r"\w+", paya["description"].lower()))
        if bwords and pwords:
            score += int(20 * len(bwords & pwords) / max(1, len(bwords | pwords)))
    return score


def reconcile(bank_data, paya_data, date_tolerance=0, bank_mapping=None, paya_mapping=None):
    bank_cols = infer_columns(bank_data, bank_mapping)
    paya_cols = infer_columns(paya_data, paya_mapping)

    bank_txs = [make_tx(r, i + 2, bank_cols, "BANK")
                for i, r in enumerate(bank_data["rows"])]
    paya_txs = [make_tx(r, i + 2, paya_cols, "PAYA")
                for i, r in enumerate(paya_data["rows"])]

    # Candidate matching. Exact ID, then exact date+amount, then best score.
    unmatched_paya = set(range(len(paya_txs)))
    matched = []
    bank_only = []
    used_paya = set()

    for bi, b in enumerate(bank_txs):
        candidates = []

        for pi in list(unmatched_paya):
            p = paya_txs[pi]
            s = similarity_score(b, p)
            if s >= 100:
                candidates.append((s, pi))

        if not candidates:
            for pi in list(unmatched_paya):
                p = paya_txs[pi]
                if b["date"] and p["date"] and b["date"] != p["date"]:
                    continue
                if same_amount(b["amount"], p["amount"]):
                    candidates.append((similarity_score(b, p), pi))

        if not candidates:
            # Fuzzy candidate on same date, strongest score.
            for pi in list(unmatched_paya):
                p = paya_txs[pi]
                if b["date"] and p["date"] and b["date"] != p["date"]:
                    continue
                s = similarity_score(b, p)
                if s >= 50:
                    candidates.append((s, pi))

        if candidates:
            candidates.sort(reverse=True)
            score, pi = candidates[0]
            p = paya_txs[pi]
            used_paya.add(pi)
            unmatched_paya.discard(pi)

            if same_amount(b["amount"], p["amount"]):
                status = "MATCHED"
            else:
                status = "AMOUNT_MISMATCH"

            matched.append({
                "bank": b,
                "paya": p,
                "score": score,
                "status": status,
                "amount_diff": (
                    (b["amount"] - p["amount"])
                    if b["amount"] is not None and p["amount"] is not None else None
                )
            })
        else:
            bank_only.append(b)

    paya_only = [paya_txs[i] for i in sorted(unmatched_paya)]

    # Duplicate detection by date + amount + direction + description.
    groups = {}
    for tx in bank_txs:
        key = (
            tx["date"], str(tx["amount"]), tx["direction"],
            tx["description"][:100]
        )
        groups.setdefault(key, []).append(tx)
    duplicates = [v for v in groups.values() if len(v) > 1]

    return {
        "bank_cols": bank_cols,
        "paya_cols": paya_cols,
        "bank_txs": bank_txs,
        "paya_txs": paya_txs,
        "matched": matched,
        "bank_only": bank_only,
        "paya_only": paya_only,
        "duplicates": duplicates,
    }



def balance_control(bank_txs):
    """
    Attempts an arithmetic control when opening/closing balances and signed
    transaction directions are available. It never invents an opening balance.
    """
    balances = [t["balance"] for t in bank_txs if t["balance"] is not None]
    if not balances:
        return {"available": False, "reason": "No balance column was identified."}

    closing = balances[-1]
    opening = balances[0]
    signed_sum = Decimal("0")
    for tx in bank_txs:
        if tx["amount"] is None:
            continue
        if tx["direction"] == "CREDIT":
            signed_sum += tx["amount"]
        elif tx["direction"] == "DEBIT":
            signed_sum -= tx["amount"]

    expected_change = closing - opening
    difference = expected_change - signed_sum
    return {
        "available": True,
        "opening_balance": opening,
        "closing_balance": closing,
        "signed_transaction_change": signed_sum,
        "difference": difference,
        "status": "OK" if difference == 0 else "REVIEW"
    }

def daily_balance_check(bank_txs):
    by_date = {}
    for tx in bank_txs:
        d = tx["date"] or "(بدون تاریخ)"
        by_date.setdefault(d, []).append(tx)

    result = []
    for d, txs in sorted(by_date.items(), key=lambda x: x[0]):
        balances = [t["balance"] for t in txs if t["balance"] is not None]
        result.append({
            "date": d,
            "transactions": len(txs),
            "last_balance": balances[-1] if balances else None,
            "first_balance": balances[0] if balances else None,
        })
    return result


def fmt_num(v):
    if v is None:
        return ""
    if isinstance(v, Decimal):
        return str(v)
    return str(v)


def tx_export_row(tx):
    return [
        tx["index"], tx["date"], fmt_num(tx["amount"]), tx["direction"],
        tx["id"], tx["description"], fmt_num(tx["balance"])
    ]


def api_chat(cfg, messages):
    base = cfg["base_url"].rstrip("/")
    url = base + "/chat/completions"
    headers = {
        "Authorization": "Bearer " + cfg["api_key"],
        "Content-Type": "application/json",
    }
    payload = {
        "model": cfg["model"],
        "messages": messages,
        "temperature": float(cfg["temperature"]),
        "max_tokens": int(cfg["max_tokens"]),
    }
    r = requests.post(url, headers=headers, json=payload,
                      timeout=int(cfg["timeout"]))
    if r.status_code >= 400:
        raise RuntimeError(f"API HTTP {r.status_code}: {r.text[:2000]}")
    data = r.json()
    try:
        return data["choices"][0]["message"]["content"]
    except Exception:
        return json.dumps(data, ensure_ascii=False)


def build_all_rows_payload(bank_data, paya_data):
    # Every row is included. The application splits the complete dataset into
    # numbered chunks because one HTTP request may exceed context limits.
    bank_rows = []
    for i, row in enumerate(bank_data["rows"], start=1):
        bank_rows.append({"row_number": i, "values": row})
    paya_rows = []
    for i, row in enumerate(paya_data["rows"], start=1):
        paya_rows.append({"row_number": i, "values": row})
    return bank_rows, paya_rows


def send_all_rows_to_ai(cfg, bank_data, paya_data, log):
    """
    Sends every source row to the model in chunks.
    The model is told that chunks are parts of one complete dataset.
    Each call returns a compact acknowledgement/observations payload.
    Then a final synthesis call receives the deterministic reconciliation
    plus all chunk observations.
    """
    if not cfg.get("ai_enabled") or not cfg.get("api_key"):
        return {"enabled": False, "chunks": [], "final": "AI disabled or API key empty."}

    bank_rows, paya_rows = build_all_rows_payload(bank_data, paya_data)
    all_records = (
        [{"source": "BANK", **x} for x in bank_rows] +
        [{"source": "PAYA", **x} for x in paya_rows]
    )
    chunk_size = max(20, int(cfg.get("chunk_rows", 150)))
    chunks = [
        all_records[i:i + chunk_size]
        for i in range(0, len(all_records), chunk_size)
    ]

    observations = []
    total = len(chunks)

    for n, chunk in enumerate(chunks, start=1):
        log(f"ارسال بخش {n} از {total} به مدل؛ تعداد ردیف: {len(chunk)}")
        prompt = {
            "instruction": (
                "You are a bank reconciliation analyst. This is chunk "
                f"{n}/{total} of a COMPLETE dataset. Do not invent rows. "
                "Inspect every supplied row. Return compact JSON-like text "
                "with observations, suspicious duplicates, date/amount anomalies, "
                "and useful candidate relationships. Do not claim final reconciliation "
                "because other chunks may contain matching rows."
            ),
            "bank_file": {
                "sheet": bank_data["sheet"],
                "headers": bank_data["headers"],
            },
            "paya_file": {
                "sheet": paya_data["sheet"],
                "headers": paya_data["headers"],
            },
            "records": chunk,
        }
        content = api_chat(cfg, [
            {"role": "system", "content": "You analyze financial reconciliation data. Be exact and concise."},
            {"role": "user", "content": json.dumps(prompt, ensure_ascii=False, default=str)}
        ])
        observations.append({"chunk": n, "rows": len(chunk), "analysis": content})

    # Final synthesis gets the deterministic engine result separately.
    log("در حال ارسال مرحله جمع‌بندی نهایی به مدل...")
    return {"enabled": True, "chunks": observations}


def final_ai_synthesis(cfg, ai_state, recon, log):
    if not ai_state.get("enabled"):
        return "AI disabled."

    def simple_tx(tx):
        return {
            "row": tx["index"],
            "date": tx["date"],
            "amount": fmt_num(tx["amount"]),
            "direction": tx["direction"],
            "id": tx["id"],
            "description": tx["description"],
        }

    deterministic = {
        "matched": len(recon["matched"]),
        "bank_only": len(recon["bank_only"]),
        "paya_only": len(recon["paya_only"]),
        "amount_mismatch": sum(1 for x in recon["matched"] if x["status"] == "AMOUNT_MISMATCH"),
        "duplicates": sum(len(x) for x in recon["duplicates"]),
        "bank_only_rows": [simple_tx(x) for x in recon["bank_only"][:500]],
        "paya_only_rows": [simple_tx(x) for x in recon["paya_only"][:500]],
        "amount_mismatches": [
            {
                "bank": simple_tx(x["bank"]),
                "paya": simple_tx(x["paya"]),
                "difference": fmt_num(x["amount_diff"]),
                "score": x["score"],
            }
            for x in recon["matched"] if x["status"] == "AMOUNT_MISMATCH"
        ][:500],
    }

    compact_obs = ai_state["chunks"]
    # Keep all chunk analyses but cap individual strings to avoid an accidental
    # oversized synthesis request.
    compact_obs = [
        {"chunk": x["chunk"], "rows": x["rows"], "analysis": x["analysis"][:6000]}
        for x in compact_obs
    ]

    prompt = {
        "task": (
            "Produce the final analytical reconciliation report. The deterministic "
            "engine is the authoritative arithmetic layer. Use the AI observations "
            "only to explain ambiguities and identify likely false matches. Never "
            "change a numeric fact without evidence. State clearly that the final "
            "balance control must equal the bank statement balance."
        ),
        "deterministic_reconciliation": deterministic,
        "ai_chunk_observations": compact_obs,
        "required_sections": [
            "Executive summary",
            "Bank-only transactions",
            "PAYA-only transactions",
            "Amount mismatches",
            "Potential duplicate issues",
            "Ambiguous/high-risk cases",
            "Recommended review actions",
            "Final control conclusion"
        ]
    }

    return api_chat(cfg, [
        {"role": "system", "content": "You are a senior financial reconciliation reviewer."},
        {"role": "user", "content": json.dumps(prompt, ensure_ascii=False, default=str)}
    ])


def write_sheet(ws, headers, rows, widths=None):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(1, c, h)
        cell.fill = HEADER_FILL
        cell.font = WHITE_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r, row in enumerate(rows, start=2):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(r, c, value)
            cell.border = Border(bottom=THIN)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    if widths:
        for i, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = width
    else:
        for col in ws.columns:
            max_len = 10
            for cell in col[:300]:
                max_len = max(max_len, len(str(cell.value or "")))
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 2, 50)


def export_result(path, bank_data, paya_data, recon, ai_text):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    total_amount_mismatch = sum(
        (x["amount_diff"] or Decimal("0")) for x in recon["matched"]
        if x["status"] == "AMOUNT_MISMATCH"
    )

    summary_rows = [
        ["Report generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["Bank file", bank_data["path"]],
        ["Bank sheet", bank_data["sheet"]],
        ["PAYA file", paya_data["path"]],
        ["PAYA sheet", paya_data["sheet"]],
        ["Bank rows", len(recon["bank_txs"])],
        ["PAYA rows", len(recon["paya_txs"])],
        ["Matched", len(recon["matched"])],
        ["Bank only", len(recon["bank_only"])],
        ["PAYA only", len(recon["paya_only"])],
        ["Amount mismatches", sum(1 for x in recon["matched"] if x["status"] == "AMOUNT_MISMATCH")],
        ["Duplicate groups", len(recon["duplicates"])],
        ["Total amount difference", fmt_num(total_amount_mismatch)],
    ]
    for r, row in enumerate(summary_rows, start=1):
        ws.cell(r, 1, row[0]).font = BOLD_FONT
        ws.cell(r, 2, row[1])
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 80

    # Matched
    ws = wb.create_sheet("Matched")
    rows = []
    for x in recon["matched"]:
        b, p = x["bank"], x["paya"]
        rows.append([
            b["index"], b["date"], fmt_num(b["amount"]), b["direction"], b["id"],
            b["description"], p["index"], p["date"], fmt_num(p["amount"]),
            p["direction"], p["id"], p["description"], x["score"], x["status"],
            fmt_num(x["amount_diff"])
        ])
    write_sheet(ws, [
        "Bank Row", "Bank Date", "Bank Amount", "Bank Direction", "Bank ID",
        "Bank Description", "PAYA Row", "PAYA Date", "PAYA Amount",
        "PAYA Direction", "PAYA ID", "PAYA Description", "Match Score",
        "Status", "Amount Difference"
    ], rows)

    # Bank only
    ws = wb.create_sheet("Bank Only")
    write_sheet(ws, ["Bank Row", "Date", "Amount", "Direction", "ID", "Description", "Balance"],
                [tx_export_row(x) for x in recon["bank_only"]])

    # PAYA only
    ws = wb.create_sheet("PAYA Only")
    write_sheet(ws, ["PAYA Row", "Date", "Amount", "Direction", "ID", "Description", "Balance"],
                [tx_export_row(x) for x in recon["paya_only"]])

    # Amount mismatches
    ws = wb.create_sheet("Amount Mismatch")
    mm_rows = []
    for x in recon["matched"]:
        if x["status"] == "AMOUNT_MISMATCH":
            b, p = x["bank"], x["paya"]
            mm_rows.append([
                b["index"], b["date"], fmt_num(b["amount"]), p["index"],
                p["date"], fmt_num(p["amount"]), fmt_num(x["amount_diff"]),
                b["description"], p["description"], x["score"]
            ])
    write_sheet(ws, [
        "Bank Row", "Bank Date", "Bank Amount", "PAYA Row", "PAYA Date",
        "PAYA Amount", "Difference", "Bank Description", "PAYA Description",
        "Match Score"
    ], mm_rows)

    # Duplicates
    ws = wb.create_sheet("Duplicates")
    dup_rows = []
    for group_no, group in enumerate(recon["duplicates"], start=1):
        for tx in group:
            dup_rows.append([
                group_no, tx["index"], tx["date"], fmt_num(tx["amount"]),
                tx["direction"], tx["id"], tx["description"]
            ])
    write_sheet(ws, ["Group", "Bank Row", "Date", "Amount", "Direction", "ID", "Description"], dup_rows)

    # Daily balances
    ws = wb.create_sheet("Daily Balance")
    daily = daily_balance_check(recon["bank_txs"])
    write_sheet(ws, ["Date", "Transactions", "First Balance", "Last Balance"],
                [[x["date"], x["transactions"], fmt_num(x["first_balance"]), fmt_num(x["last_balance"])]
                 for x in daily])

    # AI analysis
    ws = wb.create_sheet("AI Analysis")
    ws["A1"] = "AI Final Analysis"
    ws["A1"].fill = HEADER_FILL
    ws["A1"].font = WHITE_FONT
    ws["A2"] = ai_text
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 120

    # Source data, exactly as imported
    ws = wb.create_sheet("Source Bank")
    write_sheet(ws, bank_data["headers"], [list(r.values()) for r in bank_data["rows"]])

    ws = wb.create_sheet("Source PAYA")
    write_sheet(ws, paya_data["headers"], [list(r.values()) for r in paya_data["rows"]])

    # Final control
    ws = wb.create_sheet("Final Control")
    bc = balance_control(recon["bank_txs"])
    control = [
        ["Control", "Value", "Status"],
        ["Bank transaction count", len(recon["bank_txs"]), "OK"],
        ["PAYA transaction count", len(recon["paya_txs"]), "INFO"],
        ["Matched", len(recon["matched"]), "OK"],
        ["Bank-only", len(recon["bank_only"]), "REVIEW" if recon["bank_only"] else "OK"],
        ["PAYA-only", len(recon["paya_only"]), "REVIEW" if recon["paya_only"] else "OK"],
        ["Amount mismatches",
         sum(1 for x in recon["matched"] if x["status"] == "AMOUNT_MISMATCH"),
         "REVIEW" if any(x["status"] == "AMOUNT_MISMATCH" for x in recon["matched"]) else "OK"],
        ["Total amount difference", fmt_num(total_amount_mismatch),
         "OK" if total_amount_mismatch == 0 else "REVIEW"],
        ["Bank opening balance", fmt_num(bc.get("opening_balance")) if bc.get("available") else "",
         "OK" if bc.get("available") else "INFO"],
        ["Bank closing balance", fmt_num(bc.get("closing_balance")) if bc.get("available") else "",
         "OK" if bc.get("available") else "INFO"],
        ["Signed transaction change", fmt_num(bc.get("signed_transaction_change")) if bc.get("available") else "",
         "OK" if bc.get("available") else "INFO"],
        ["Balance arithmetic difference", fmt_num(bc.get("difference")) if bc.get("available") else "",
         bc.get("status", "INFO")],
    ]
    for r, row in enumerate(control, start=1):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(r, c, value)
            cell.border = Border(bottom=THIN)
            if r == 1:
                cell.fill = HEADER_FILL
                cell.font = WHITE_FONT
            elif c == 3:
                if value == "OK":
                    cell.fill = OK_FILL
                elif value == "REVIEW":
                    cell.fill = WARN_FILL
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 20

    wb.save(path)

class ColumnMappingDialog(tk.Toplevel):
    def __init__(self, parent, data, detected, title):
        super().__init__(parent)
        self.title(title)
        self.geometry("720x560")
        self.transient(parent)
        self.grab_set()
        self.result = None
        self.data = data
        self.detected = detected
        self.vars = {}

        ttk.Label(
            self,
            text="ساختار فایل را بررسی کنید. ستون‌های تشخیص‌داده‌شده را در صورت نیاز اصلاح کنید.",
            wraplength=650, justify="right"
        ).pack(fill="x", padx=15, pady=12)

        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True, padx=15)

        fields = [
            ("date", "تاریخ تراکنش", DATE_WORDS),
            ("amount", "مبلغ واحد", AMOUNT_WORDS),
            ("debit", "برداشت / بدهکار", DEBIT_WORDS),
            ("credit", "واریز / بستانکار", CREDIT_WORDS),
            ("balance", "مانده", BALANCE_WORDS),
            ("id", "شناسه / شماره پیگیری", ID_WORDS),
            ("description", "شرح", DESC_WORDS),
            ("type", "نوع تراکنش", TYPE_WORDS),
        ]

        headers = [""] + data["headers"]
        for r, (key, label, _) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=r, column=0, sticky="e", padx=8, pady=7)
            var = tk.StringVar(value=detected.get(key) or "")
            self.vars[key] = var
            combo = ttk.Combobox(frame, textvariable=var, values=headers[1:],
                                 state="readonly", width=55)
            combo.grid(row=r, column=1, sticky="ew", padx=8, pady=7)

        ttk.Label(
            frame,
            text="اگر فایل شما «واریز/برداشت» دارد، آن‌ها را در فیلدهای مربوط انتخاب کنید. "
                 "اگر فقط «مبلغ» دارید، Amount را انتخاب کنید و Direction از نوع/شرح تشخیص داده می‌شود.",
            wraplength=620, justify="right"
        ).grid(row=len(fields), column=0, columnspan=2, sticky="e", pady=12)

        buttons = ttk.Frame(self)
        buttons.pack(fill="x", padx=15, pady=12)
        ttk.Button(buttons, text="لغو", command=self.destroy).pack(side="left")
        ttk.Button(buttons, text="تأیید ساختار", command=self.accept).pack(side="right")

        frame.columnconfigure(1, weight=1)

    def accept(self):
        self.result = {k: v.get().strip() for k, v in self.vars.items()}
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("مغایرت‌گیری بانک و پایا - Enterprise")
        self.geometry("1200x780")
        self.minsize(1000, 680)
        self.cfg = load_config()
        self.bank_data = None
        self.paya_data = None
        self.recon = None
        self.bank_mapping = {}
        self.paya_mapping = {}

        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except Exception:
            pass

        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Label(
            top, text="سامانه مغایرت‌گیری بانک و پایا",
            font=("Segoe UI", 18, "bold")
        ).pack(side="right")
        ttk.Label(
            top, text="پردازش محلی + تحلیل هوش مصنوعی",
            font=("Segoe UI", 10)
        ).pack(side="left")

        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=5)

        self.main_tab = ttk.Frame(self.tabs, padding=12)
        self.settings_tab = ttk.Frame(self.tabs, padding=12)
        self.tabs.add(self.main_tab, text="  مغایرت‌گیری  ")
        self.tabs.add(self.settings_tab, text="  تنظیمات مدل و API  ")

        self._build_main()
        self._build_settings()

    def _build_main(self):
        files = ttk.LabelFrame(self.main_tab, text="انتخاب فایل‌ها", padding=12)
        files.pack(fill="x")

        self.bank_var = tk.StringVar()
        self.paya_var = tk.StringVar()

        ttk.Label(files, text="صورتحساب بانک:").grid(row=0, column=0, sticky="e", padx=5, pady=8)
        ttk.Entry(files, textvariable=self.bank_var, width=85).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(files, text="انتخاب Excel", command=self.pick_bank).grid(row=0, column=2, padx=5)

        ttk.Label(files, text="گزارش پایا:").grid(row=1, column=0, sticky="e", padx=5, pady=8)
        ttk.Entry(files, textvariable=self.paya_var, width=85).grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(files, text="انتخاب Excel", command=self.pick_paya).grid(row=1, column=2, padx=5)

        files.columnconfigure(1, weight=1)

        info = ttk.LabelFrame(self.main_tab, text="تشخیص ساختار", padding=10)
        info.pack(fill="x", pady=10)
        self.bank_info = tk.StringVar(value="فایل بانک انتخاب نشده است.")
        self.paya_info = tk.StringVar(value="فایل پایا انتخاب نشده است.")
        ttk.Label(info, textvariable=self.bank_info).pack(anchor="e", fill="x")
        ttk.Label(info, textvariable=self.paya_info).pack(anchor="e", fill="x")

        actions = ttk.Frame(self.main_tab)
        actions.pack(fill="x", pady=5)

        self.run_btn = ttk.Button(actions, text="شروع مغایرت‌گیری و تحلیل", command=self.start)
        self.run_btn.pack(side="right", padx=5)
        ttk.Button(actions, text="پاک کردن", command=self.clear).pack(side="right", padx=5)
        ttk.Button(actions, text="باز کردن پوشه خروجی", command=self.open_output).pack(side="left", padx=5)

        progress_frame = ttk.LabelFrame(self.main_tab, text="وضعیت عملیات", padding=8)
        progress_frame.pack(fill="both", expand=True, pady=10)

        self.progress = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress.pack(fill="x", pady=5)
        self.log = tk.Text(progress_frame, height=25, wrap="word", font=("Consolas", 10))
        self.log.pack(fill="both", expand=True)
        self.log.configure(state="disabled")

    def _build_settings(self):
        frm = ttk.Frame(self.settings_tab)
        frm.pack(fill="both", expand=True)

        fields = [
            ("Base URL", "base_url"),
            ("API Key", "api_key"),
            ("Model", "model"),
            ("Temperature", "temperature"),
            ("Max Tokens", "max_tokens"),
            ("Timeout (sec)", "timeout"),
            ("Chunk Rows", "chunk_rows"),
        ]

        self.settings_vars = {}
        for r, (label, key) in enumerate(fields):
            ttk.Label(frm, text=label).grid(row=r, column=0, sticky="e", padx=10, pady=9)
            show = "*" if key == "api_key" else ""
            var = tk.StringVar(value=str(self.cfg.get(key, "")))
            self.settings_vars[key] = var
            ttk.Entry(frm, textvariable=var, show=show, width=70).grid(
                row=r, column=1, sticky="ew", padx=10, pady=9
            )

        self.ai_enabled_var = tk.BooleanVar(value=bool(self.cfg.get("ai_enabled", True)))
        self.send_all_var = tk.BooleanVar(value=bool(self.cfg.get("send_all_rows", True)))

        ttk.Checkbutton(
            frm, text="فعال بودن تحلیل AI",
            variable=self.ai_enabled_var
        ).grid(row=7, column=1, sticky="e", padx=10, pady=5)

        ttk.Checkbutton(
            frm, text="ارسال تمام ردیف‌های هر دو فایل به مدل (Chunked)",
            variable=self.send_all_var
        ).grid(row=8, column=1, sticky="e", padx=10, pady=5)

        buttons = ttk.Frame(frm)
        buttons.grid(row=9, column=1, sticky="e", padx=10, pady=20)
        ttk.Button(buttons, text="ذخیره تنظیمات", command=self.save_settings).pack(side="right", padx=5)
        ttk.Button(buttons, text="تست اتصال API", command=self.test_api).pack(side="right", padx=5)

        note = (
            "نکته: API Key فقط در config.json کنار برنامه ذخیره می‌شود و به مرورگر ارسال نمی‌شود. "
            "برای APIهای OpenAI-compatible باید Base URL تا /v1 باشد."
        )
        ttk.Label(frm, text=note, wraplength=800, justify="right").grid(
            row=10, column=1, sticky="e", padx=10, pady=15
        )
        frm.columnconfigure(1, weight=1)

    def save_settings(self):
        try:
            cfg = dict(self.cfg)
            cfg["base_url"] = self.settings_vars["base_url"].get().strip()
            cfg["api_key"] = self.settings_vars["api_key"].get().strip()
            cfg["model"] = self.settings_vars["model"].get().strip()
            cfg["temperature"] = float(self.settings_vars["temperature"].get())
            cfg["max_tokens"] = int(self.settings_vars["max_tokens"].get())
            cfg["timeout"] = int(self.settings_vars["timeout"].get())
            cfg["chunk_rows"] = int(self.settings_vars["chunk_rows"].get())
            cfg["ai_enabled"] = bool(self.ai_enabled_var.get())
            cfg["send_all_rows"] = bool(self.send_all_var.get())
            self.cfg = cfg
            save_config(cfg)
            messagebox.showinfo("ذخیره شد", "تنظیمات با موفقیت ذخیره شد.")
        except Exception as e:
            messagebox.showerror("خطا", f"تنظیمات نامعتبر است:\n{e}")

    def test_api(self):
        self.save_settings()
        if not self.cfg.get("api_key"):
            messagebox.showwarning("API Key", "ابتدا API Key را وارد کنید.")
            return
        try:
            answer = api_chat(self.cfg, [
                {"role": "system", "content": "Reply only with OK."},
                {"role": "user", "content": "Connection test. Reply only with OK."}
            ])
            messagebox.showinfo("نتیجه تست", answer[:1000])
        except Exception as e:
            messagebox.showerror("خطای API", str(e))

    def pick_bank(self):
        p = filedialog.askopenfilename(
            title="انتخاب صورتحساب بانک",
            filetypes=[("Excel Workbook", "*.xlsx *.xlsm *.xltx *.xltm"), ("All files", "*.*")]
        )
        if not p:
            return
        try:
            self.bank_var.set(p)
            self.bank_data = read_excel(p)
            detected = infer_columns(self.bank_data)
            dlg = ColumnMappingDialog(self, self.bank_data, detected, "ساختار صورتحساب بانک")
            self.wait_window(dlg)
            if dlg.result is None:
                self.bank_var.set("")
                self.bank_data = None
                return
            self.bank_mapping = dlg.result
            cols = infer_columns(self.bank_data, self.bank_mapping)
            self.bank_info.set(
                f"بانک: Sheet={self.bank_data['sheet']} | Header Row={self.bank_data['header_row']} | "
                f"Rows={len(self.bank_data['rows'])} | "
                f"Date={cols.get('date')} | Amount={cols.get('amount')} | "
                f"Debit={cols.get('debit')} | Credit={cols.get('credit')} | "
                f"Balance={cols.get('balance')} | ID={cols.get('id')}"
            )
            self.write_log("فایل بانک با موفقیت خوانده شد.")
        except Exception as e:
            messagebox.showerror("خطا در فایل بانک", str(e))

    def pick_paya(self):
        p = filedialog.askopenfilename(
            title="انتخاب گزارش پایا",
            filetypes=[("Excel Workbook", "*.xlsx *.xlsm *.xltx *.xltm"), ("All files", "*.*")]
        )
        if not p:
            return
        try:
            self.paya_var.set(p)
            self.paya_data = read_excel(p)
            detected = infer_columns(self.paya_data)
            dlg = ColumnMappingDialog(self, self.paya_data, detected, "ساختار گزارش پایا")
            self.wait_window(dlg)
            if dlg.result is None:
                self.paya_var.set("")
                self.paya_data = None
                return
            self.paya_mapping = dlg.result
            cols = infer_columns(self.paya_data, self.paya_mapping)
            self.paya_info.set(
                f"پایا: Sheet={self.paya_data['sheet']} | Header Row={self.paya_data['header_row']} | "
                f"Rows={len(self.paya_data['rows'])} | "
                f"Date={cols.get('date')} | Amount={cols.get('amount')} | "
                f"Debit={cols.get('debit')} | Credit={cols.get('credit')} | "
                f"Balance={cols.get('balance')} | ID={cols.get('id')}"
            )
            self.write_log("فایل پایا با موفقیت خوانده شد.")
        except Exception as e:
            messagebox.showerror("خطا در فایل پایا", str(e))

    def clear(self):
        self.bank_data = None
        self.paya_data = None
        self.recon = None
        self.bank_mapping = {}
        self.paya_mapping = {}
        self.bank_var.set("")
        self.paya_var.set("")
        self.bank_info.set("فایل بانک انتخاب نشده است.")
        self.paya_info.set("فایل پایا انتخاب نشده است.")
        self.write_log("پاک شد.")

    def write_log(self, msg):
        def inner():
            self.log.configure(state="normal")
            self.log.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
            self.log.see("end")
            self.log.configure(state="disabled")
        self.after(0, inner)

    def set_progress(self, value):
        self.after(0, lambda: self.progress.configure(value=value))

    def start(self):
        if not self.bank_data or not self.paya_data:
            messagebox.showwarning("فایل‌ها", "هر دو فایل بانک و پایا را انتخاب کنید.")
            return
        self.save_settings()
        self.run_btn.configure(state="disabled")
        self.progress["value"] = 0
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            self.write_log("شروع پردازش...")
            self.set_progress(10)

            self.write_log("در حال اجرای موتور مغایرت‌گیری محلی...")
            recon = reconcile(self.bank_data, self.paya_data, bank_mapping=self.bank_mapping, paya_mapping=self.paya_mapping)
            self.recon = recon
            self.set_progress(35)

            self.write_log(
                f"نتیجه اولیه: تطبیق={len(recon['matched'])} | "
                f"بانک-only={len(recon['bank_only'])} | "
                f"پایا-only={len(recon['paya_only'])} | "
                f"مغایرت مبلغ={sum(1 for x in recon['matched'] if x['status']=='AMOUNT_MISMATCH')}"
            )

            ai_state = {"enabled": False, "chunks": [], "final": ""}
            if self.cfg.get("ai_enabled") and self.cfg.get("api_key") and self.cfg.get("send_all_rows"):
                ai_state = send_all_rows_to_ai(self.cfg, self.bank_data, self.paya_data, self.write_log)
                self.set_progress(75)
                ai_text = final_ai_synthesis(self.cfg, ai_state, recon, self.write_log)
            else:
                self.write_log("ارسال به AI غیرفعال است یا API Key وارد نشده.")
                ai_text = "AI analysis was not executed."

            self.set_progress(88)

            out = OUTPUT_DIR / (
                "Bank_PAYA_Reconciliation_"
                + datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
            )
            self.write_log("در حال ساخت فایل Excel نهایی...")
            export_result(out, self.bank_data, self.paya_data, recon, ai_text)
            self.set_progress(100)
            self.write_log(f"فایل نهایی ساخته شد: {out}")

            self.after(0, lambda: messagebox.showinfo(
                "پایان",
                f"مغایرت‌گیری تمام شد.\n\nفایل خروجی:\n{out}"
            ))
        except Exception as e:
            err = traceback.format_exc()
            self.write_log(err)
            self.after(0, lambda: messagebox.showerror("خطای اجرا", str(e)))
        finally:
            self.after(0, lambda: self.run_btn.configure(state="normal"))

    def open_output(self):
        try:
            os.startfile(str(OUTPUT_DIR))
        except Exception:
            messagebox.showinfo("Output", str(OUTPUT_DIR))


if __name__ == "__main__":
    App().mainloop()
