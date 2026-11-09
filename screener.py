#!/usr/bin/env python3
"""
price-movers - daily NSE upside movers, sourced entirely from the TradingView
scanner feed (https://scanner.tradingview.com/india/scan).

Screen - every condition must hold on the day's close:

    exchange           =  NSE
    change             >= +5.0 %      (upside only, vs previous close)
    average volume 10d >  500,000     (shares)
    market cap         >  Rs 100 cr
    last price         >  Rs 19

Every run ends with exactly one Telegram message: either the updated CSV, or a
description of what went wrong.
"""

import argparse
import csv
import json
import os
import sys
import time
import traceback
from datetime import datetime, time
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

IST = ZoneInfo("Asia/Kolkata")
SCAN_URL = "https://scanner.tradingview.com/india/scan"
SEC_LIST_URL = "https://nsearchives.nseindia.com/content/equities/sec_list.csv"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

# ---------------------------------------------------------------- thresholds
MIN_CHANGE_PCT = 5.0
MIN_AVG_VOL_10D = 500_000
MIN_MARKET_CAP = 1_000_000_000      # Rs 100 crore
MIN_PRICE = 19.0
CRORE = 10_000_000

# Columns requested from the scanner, in the order the API returns them.
COLUMNS = [
    "description",
    "close",
    "change",
    "volume",
    "average_volume_10d_calc",
    "relative_volume_10d_calc",
    "market_cap_basic",
]

# CSV layout - the order the user sees in the file.
FIELDS = [
    "date",
    "symbol",
    "fyers_code",
    "company",
    "close",
    "change_pct",
    "volume",
    "avg_volume_10d",
    "rvol",
    "market_cap_cr",
    "exchange",
    "series",
]

CSV_PATH = Path(__file__).resolve().with_name("price_movers.csv")
XLSX_PATH = Path(__file__).resolve().with_name("price_movers.xlsx")
SHEET_NAME = "movers"

# Fixed-date NSE holidays (certain every year). Variable-date holidays - Holi,
# Diwali, Id, Good Friday etc. - must be added here by hand each January when
# NSE publishes the next year's calendar. See README.
NSE_FIXED_HOLIDAYS = {
    "01-26",  # Republic Day
    "05-01",  # Maharashtra Day
    "08-15",  # Independence Day
    "10-02",  # Gandhi Jayanti
    "12-25",  # Christmas
}


# ------------------------------------------------------------------ fetching
def build_payload():
    """Server-side filter. Bounds are inclusive and slightly loose on purpose;
    the exact comparisons are re-applied locally in `passes()`."""
    return {
        "filter": [
            {"left": "exchange", "operation": "equal", "right": "NSE"},
            {"left": "type", "operation": "equal", "right": "stock"},
            {"left": "change", "operation": "in_range",
             "right": [MIN_CHANGE_PCT, 1_000_000]},
            {"left": "average_volume_10d_calc", "operation": "in_range",
             "right": [MIN_AVG_VOL_10D, 1_000_000_000_000]},
            {"left": "market_cap_basic", "operation": "in_range",
             "right": [MIN_MARKET_CAP, 1_000_000_000_000_000]},
            {"left": "close", "operation": "in_range",
             "right": [MIN_PRICE, 10_000_000]},
        ],
        "options": {"lang": "en"},
        "markets": ["india"],
        "columns": COLUMNS,
        "sort": {"sortBy": "change", "sortOrder": "desc"},
        "range": [0, 5000],
    }


def fetch_scan(attempts=4, timeout=30):
    """POST the screen to TradingView. Returns the decoded JSON body."""
    payload = build_payload()
    headers = {"User-Agent": USER_AGENT, "Content-Type": "application/json"}
    last_error = None

    for attempt in range(1, attempts + 1):
        try:
            r = requests.post(SCAN_URL, json=payload, headers=headers,
                              timeout=timeout)
            if r.status_code == 200:
                return r.json()
            if r.status_code == 429:
                last_error = "HTTP 429 (rate limited)"
            else:
                last_error = f"HTTP {r.status_code}: {r.text[:300]}"
        except Exception as exc:                       # noqa: BLE001
            last_error = f"{type(exc).__name__}: {exc}"
        if attempt < attempts:
            wait = 5 * 2 ** (attempt - 1)
            print(f"  attempt {attempt} failed ({last_error}); retrying in {wait}s",
                  file=sys.stderr)
            time.sleep(wait)

    raise RuntimeError(f"TradingView scan failed after {attempts} attempts - "
                       f"{last_error}")


# ------------------------------------------------------------- NSE series
def fetch_series_map(attempts=3, timeout=45):
    """symbol -> series (EQ / BE / BZ / SM ...) from NSE's securities list.

    This is the current classification, not a per-date one. Every row in the
    file is normalised to it, so a stock that moved EQ -> BE shows BE on all
    of its rows, however old.
    """
    last_error = None
    for attempt in range(1, attempts + 1):
        try:
            r = requests.get(SEC_LIST_URL, headers={"User-Agent": USER_AGENT},
                             timeout=timeout)
            if r.status_code == 200:
                out = {}
                for row in csv.DictReader(r.text.splitlines()):
                    sym = (row.get("Symbol") or "").strip()
                    ser = (row.get("Series") or "").strip().upper()
                    if sym and ser:
                        out[_norm(sym)] = {"symbol": sym.upper(), "series": ser}
                if out:
                    return out
                last_error = "empty securities list"
            else:
                last_error = f"HTTP {r.status_code}"
        except Exception as exc:                       # noqa: BLE001
            last_error = f"{type(exc).__name__}: {exc}"
        if attempt < attempts:
            time.sleep(3 * attempt)
    print(f"WARNING: could not load NSE series list ({last_error}); "
          f"keeping the series already recorded.", file=sys.stderr)
    return {}


def _norm(symbol):
    """TradingView writes M&M as NSE:M_M, so compare symbols with the
    punctuation flattened."""
    return str(symbol).strip().upper().replace("&", "_").replace("-", "_")


def apply_latest_series(rows, series_map):
    """Rewrite series *and* the Fyers code of every row to the symbol's current
    classification. The code reads NSE:<official symbol>-<series>, followed by
    a comma and a space so the column can be concatenated straight into a
    broker request."""
    if not series_map:
        return rows
    for row in rows:
        hit = series_map.get(_norm(row.get("symbol", "")))
        if hit:
            row["series"] = hit["series"]
            row["fyers_code"] = f"NSE:{hit['symbol']}-{hit['series']}, "
        elif row.get("series"):
            row["fyers_code"] = f"NSE:{row['symbol']}-{row['series']}, "
    return rows


# ------------------------------------------------------------------ parsing
def _num(value):
    try:
        if value is None or value == "":
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def parse_scan(body):
    """Turn the scanner response into flat records."""
    records = []
    for entry in body.get("data") or []:
        ticker = str(entry.get("s") or "")
        values = list(entry.get("d") or [])
        if len(values) < len(COLUMNS):
            values += [None] * (len(COLUMNS) - len(values))
        row = dict(zip(COLUMNS, values))

        if ":" in ticker:
            exchange, symbol = ticker.split(":", 1)
        else:
            exchange, symbol = "NSE", ticker
        symbol = symbol.replace("_", "-")

        records.append({
            "symbol": symbol,
            "exchange": exchange,
            "company": (row.get("description") or "").strip(),
            "close": _num(row.get("close")),
            "change_pct": _num(row.get("change")),
            "volume": _num(row.get("volume")),
            "avg_volume_10d": _num(row.get("average_volume_10d_calc")),
            "rvol": _num(row.get("relative_volume_10d_calc")),
            "market_cap": _num(row.get("market_cap_basic")),
        })
    return records


def passes(rec):
    """Exact screen, applied locally so the thresholds are unambiguous."""
    if rec["exchange"].upper() != "NSE":
        return False
    for key in ("close", "change_pct", "avg_volume_10d", "market_cap"):
        if rec[key] is None:
            return False
    return (rec["change_pct"] >= MIN_CHANGE_PCT
            and rec["avg_volume_10d"] > MIN_AVG_VOL_10D
            and rec["market_cap"] > MIN_MARKET_CAP
            and rec["close"] > MIN_PRICE)


def to_csv_rows(records, date_str):
    rows = []
    for rec in records:
        rows.append({
            "date": date_str,
            "symbol": rec["symbol"],
            "company": rec["company"],
            "close": f"{rec['close']:.2f}",
            "change_pct": f"{rec['change_pct']:.2f}",
            "volume": int(rec["volume"]) if rec["volume"] is not None else "",
            "avg_volume_10d": int(rec["avg_volume_10d"]),
            "rvol": f"{rec['rvol']:.2f}" if rec["rvol"] is not None else "",
            "market_cap_cr": f"{rec['market_cap'] / CRORE:.0f}",
            "exchange": "NSE",
            "series": "",
            "fyers_code": "",
        })
    rows.sort(key=lambda r: float(r["change_pct"]), reverse=True)
    return rows


# ------------------------------------------------------------------ the file
# Columns stored as real numbers in the sheet, so Excel can sort and filter.
NUMERIC = {"close", "change_pct", "volume", "avg_volume_10d", "rvol",
           "market_cap_cr"}
INTEGER = {"volume", "avg_volume_10d"}
# Excel number formats: thousands separators wherever a value can reach four
# digits, so prices, volumes and market caps read at a glance.
FORMATS = {
    "close": "#,##0.00",
    "change_pct": "0.00",
    "volume": "#,##0",
    "avg_volume_10d": "#,##0",
    "rvol": "0.00",
    "market_cap_cr": "#,##0",
}
WIDTHS = {"date": 12, "symbol": 14, "fyers_code": 22, "company": 44, "close": 11,
          "change_pct": 12, "volume": 14, "avg_volume_10d": 16, "rvol": 9,
          "market_cap_cr": 15, "exchange": 10, "series": 9}


def _text_to_cell(field, value):
    """Write side: numbers go in as numbers."""
    if value is None or value == "":
        return ""
    if field in INTEGER:
        return int(float(value))
    if field in NUMERIC:
        return float(value)
    return str(value)


def read_rows(path=CSV_PATH):
    """The CSV is the record of truth: line-oriented, cheap in git, diffable."""
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8") as fh:
        return [dict(r) for r in csv.DictReader(fh)]


def write_rows(rows, path=CSV_PATH):
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=FIELDS, extrasaction="ignore",
                                lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def write_workbook(rows, path=XLSX_PATH):
    """A throwaway Excel copy, built at send time for Telegram. Never committed:
    every save rewrites the whole archive, so keeping it in git would cost far
    more history than the CSV it came from."""
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(FIELDS)
    for col in range(1, len(FIELDS) + 1):
        ws.cell(row=1, column=col).font = Font(bold=True)
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = WIDTHS.get(FIELDS[col - 1], 14)

    for row in rows:
        ws.append([_text_to_cell(f, row.get(f, "")) for f in FIELDS])

    for col, field in enumerate(FIELDS, start=1):
        fmt = FORMATS.get(field)
        if fmt:
            for cell in ws.iter_cols(min_col=col, max_col=col, min_row=2):
                for c in cell:
                    c.number_format = fmt

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(FIELDS))}{max(ws.max_row, 1)}"
    wb.save(path)


def merge_rows(existing, date_str, new_rows):
    """Drop any earlier rows for this date, prepend the fresh ones, then sort
    newest date first (and biggest mover first within a date)."""
    kept = [r for r in existing if r.get("date") != date_str]
    merged = new_rows + kept
    merged.sort(key=lambda r: (r.get("date", ""), float(r.get("change_pct") or 0)),
                reverse=True)
    return merged


def looks_like_repeat(existing, new_rows, date_str):
    """Guard against recording the same session twice - on an NSE holiday the
    scanner keeps serving the previous session's numbers. Only a *newer* date
    can be a repeat; re-running a date already in the file is an update."""
    if not existing or not new_rows:
        return False
    last_date = max(r.get("date", "") for r in existing)
    if date_str <= last_date:
        return False
    prev = {(r.get("symbol"), r.get("change_pct"))
            for r in existing if r.get("date") == last_date}
    curr = {(r.get("symbol"), r.get("change_pct")) for r in new_rows}
    return prev == curr


# --------------------------------------------------------------- telegram
def _creds():
    return os.environ.get("TELEGRAM_BOT_TOKEN"), os.environ.get("TELEGRAM_CHAT_ID")


def resolve_chat_id(token, timeout=30):
    """Work out the chat id from the bot's own recent updates when
    TELEGRAM_CHAT_ID is not configured (same idea as gift-nifty-oi's
    resolve_chat.py). Send the bot any message once and it will be found."""
    try:
        r = requests.get(f"https://api.telegram.org/bot{token}/getUpdates",
                         timeout=timeout)
        payload = r.json()
    except Exception as exc:                           # noqa: BLE001
        print(f"Could not look up the chat id: {exc}", file=sys.stderr)
        return None
    for update in reversed(payload.get("result") or []):
        msg = (update.get("message") or update.get("edited_message")
               or update.get("channel_post") or {})
        chat = msg.get("chat") or {}
        if chat.get("id"):
            return str(chat["id"])
    return None


def send_document(token, chat_id, csv_path, caption):
    with csv_path.open("rb") as fh:
        r = requests.post(f"https://api.telegram.org/bot{token}/sendDocument",
                          data={"chat_id": chat_id, "caption": caption,
                                "parse_mode": "HTML"},
                          files={"document": (csv_path.name, fh,
                                "application/vnd.openxmlformats-officedocument"
                                ".spreadsheetml.sheet")},
                          timeout=120)
    if r.status_code != 200:
        raise RuntimeError(f"Telegram sendDocument HTTP {r.status_code}: "
                           f"{r.text[:300]}")


def send_message(token, chat_id, text):
    r = requests.post(f"https://api.telegram.org/bot{token}/sendMessage",
                      data={"chat_id": chat_id, "text": text,
                            "parse_mode": "HTML", "disable_web_page_preview": True},
                      timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"Telegram sendMessage HTTP {r.status_code}: "
                           f"{r.text[:300]}")


def deliver(token, chat_id, csv_path, caption, message):
    """Post the CSV (or a note). If Telegram says the chat is unknown, look the
    chat id up from the bot's own updates and try once more."""
    def _post(cid):
        # The file is the point of the message; the summary rides along as its
        # caption. Only statuses with no file fall back to a plain message.
        if csv_path:
            send_document(token, cid, csv_path, message or caption)
        else:
            send_message(token, cid, message)

    try:
        _post(chat_id)
        return chat_id
    except RuntimeError as exc:
        if "chat not found" not in str(exc):
            raise
        alt = resolve_chat_id(token)
        if not alt:
            raise RuntimeError(
                f"Telegram says chat {chat_id} is unknown, and the bot has no "
                f"recent messages to learn a working chat id from. Send the "
                f"bot a message once (or correct TELEGRAM_CHAT_ID).") from None
        if alt == chat_id:
            raise
        print(f"Chat id {chat_id} not found; retrying with {alt} "
              f"from the bot's updates.")
        _post(alt)
        return alt


def run_url():
    server = os.environ.get("GITHUB_SERVER_URL")
    repo = os.environ.get("GITHUB_REPOSITORY")
    run_id = os.environ.get("GITHUB_RUN_ID")
    if server and repo and run_id:
        return f"{server}/{repo}/actions/runs/{run_id}"
    return None


def report_failure(token, chat_id, date_str, stage, exc):
    """Best-effort Telegram alert. Never masks the original error."""
    detail = f"{type(exc).__name__}: {exc}"[:900]
    lines = [
        "<b>price-movers - run failed</b>",
        f"Date: {date_str}",
        f"Stage: {stage}",
        "",
        f"<code>{detail}</code>",
    ]
    link = run_url()
    if link:
        lines += ["", link]
    try:
        if token and chat_id:
            send_message(token, chat_id, "\n".join(lines))
            print("Failure alert sent to Telegram.", file=sys.stderr)
        else:
            print("No Telegram credentials - cannot send the failure alert.",
                  file=sys.stderr)
    except Exception as exc2:                          # noqa: BLE001
        print(f"Failure alert could not be delivered: {exc2}", file=sys.stderr)


# --------------------------------------------------------------------- main
def is_trading_day(day):
    if day.weekday() >= 5:
        return False
    return day.strftime("%m-%d") not in NSE_FIXED_HOLIDAYS


def run(args, day, date_str):
    """The actual job. Raises on anything unexpected."""
    if not args.force and not is_trading_day(day):
        return {"status": "closed",
                "message": f"<b>Price movers - {date_str}</b>\n"
                           f"No NSE session today (weekend or holiday)."}

    # A delayed trigger can land at any hour. Before the session has closed
    # there is no EOD data for today - the feed still serves yesterday's, which
    # the repeat guard would misread as a holiday. Refuse, quietly.
    now_ist = datetime.now(IST)
    if (not args.force and day == now_ist.date()
            and now_ist.time() < time(15, 45)):
        return {"status": "too_early", "message": ""}

    csv_path = Path(args.csv)
    existing = read_rows(csv_path)

    # Retry safety: scheduled slots repeat through the afternoon, so a later
    # slot must not re-post a day that already landed.
    if args.skip_if_present and any(r.get("date") == date_str for r in existing):
        return {"status": "already", "message": ""}

    if args.fixture:
        with open(args.fixture, encoding="utf-8") as fh:
            body = json.load(fh)
        print(f"{date_str}: using fixture {args.fixture}")
    else:
        body = fetch_scan()
        print(f"{date_str}: scanner returned {len(body.get('data') or [])} rows")

    records = parse_scan(body)
    hits = [r for r in records if passes(r)]
    print(f"{date_str}: {len(hits)} of {len(records)} rows pass the screen")

    new_rows = to_csv_rows(hits, date_str)

    if not args.force and looks_like_repeat(existing, new_rows, date_str):
        return {"status": "repeat",
                "message": f"<b>Price movers - {date_str}</b>\n"
                           f"Data identical to the previous session - treating "
                           f"today as a non-trading day. No rows written."}

    merged = merge_rows(existing, date_str, new_rows)

    # The series column always reflects today's classification, applied to
    # every row in the file - not the series a stock was in on that date.
    series_map = {} if args.no_series else fetch_series_map()
    if series_map:
        apply_latest_series(merged, series_map)
        known = sum(1 for r in merged if r.get("series"))
        print(f"{date_str}: series applied to {known} of {len(merged)} rows "
              f"from {len(series_map)} listed symbols")

    write_rows(merged, csv_path)
    xlsx_path = Path(args.xlsx)
    write_workbook(merged, xlsx_path)
    dates = len({r["date"] for r in merged})
    print(f"{date_str}: wrote {len(new_rows)} rows; file now holds "
          f"{len(merged)} rows across {dates} dates")

    # The file is the message. Only a one-line label rides along with it.
    caption = f"Price movers {date_str} - {len(new_rows)} rows"
    if not new_rows:
        caption = f"Price movers {date_str} - no NSE stock met the screen"

    return {"status": "ok", "caption": caption, "message": "",
            "rows": len(new_rows), "total": len(merged),
            "csv": csv_path, "xlsx": xlsx_path}


def main(argv=None):
    ap = argparse.ArgumentParser(description="NSE upside movers from TradingView")
    ap.add_argument("--date", help="run date, YYYY-MM-DD (default: today IST)")
    ap.add_argument("--csv", default=str(CSV_PATH),
                    help="CSV that records every session")
    ap.add_argument("--xlsx", default=str(XLSX_PATH),
                    help="workbook built for sending (not committed)")
    ap.add_argument("--fixture", help="read a saved scanner response instead of "
                                      "calling TradingView (offline testing)")
    ap.add_argument("--no-telegram", action="store_true",
                    help="skip the Telegram post")
    ap.add_argument("--skip-if-present", action="store_true",
                    help="do nothing if this date is already in the file "
                         "(used by the repeated scheduled slots)")
    ap.add_argument("--no-series", action="store_true",
                    help="skip the NSE series lookup (offline runs)")
    ap.add_argument("--force", action="store_true",
                    help="run even on a weekend / holiday / repeat snapshot")
    args = ap.parse_args(argv)

    day = (datetime.strptime(args.date, "%Y-%m-%d").date()
           if args.date else datetime.now(IST).date())
    date_str = day.isoformat()
    token, chat_id = _creds()
    if token and not chat_id:
        chat_id = resolve_chat_id(token)
        if chat_id:
            print(f"Chat id {chat_id} resolved from the bot's recent updates.")

    stage = "start"
    try:
        stage = "screen"
        result = run(args, day, date_str)

        if result["status"] == "too_early":
            print(f"{date_str}: before the 15:45 IST close - no EOD data yet, "
                  f"nothing to do.")
            return 0

        if result["status"] == "already":
            print(f"{date_str}: already recorded - nothing to do.")
            return 0

        if args.no_telegram:
            print("Telegram post skipped (--no-telegram).")
            return 0
        if not token or not chat_id:
            print("TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID not set - "
                  "nothing posted.", file=sys.stderr)
            return 1

        stage = "telegram"
        if result["status"] == "ok":
            deliver(token, chat_id, result["xlsx"], result["caption"],
                    result["message"])
        else:
            deliver(token, chat_id, None, "", result["message"])
        print("Telegram post sent.")
        return 0

    except Exception as exc:                           # noqa: BLE001
        traceback.print_exc()
        report_failure(token, chat_id, date_str, stage, exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
