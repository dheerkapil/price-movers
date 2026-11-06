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
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests

IST = ZoneInfo("Asia/Kolkata")
SCAN_URL = "https://scanner.tradingview.com/india/scan"
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
    "company",
    "close",
    "change_pct",
    "volume",
    "avg_volume_10d",
    "rvol",
    "market_cap_cr",
    "exchange",
]

CSV_PATH = Path(__file__).resolve().with_name("price_movers.csv")

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
            "market_cap_cr": f"{rec['market_cap'] / CRORE:.2f}",
            "exchange": "NSE",
        })
    rows.sort(key=lambda r: float(r["change_pct"]), reverse=True)
    return rows


# ------------------------------------------------------------------ the file
def read_csv(path=CSV_PATH):
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8") as fh:
        return [dict(r) for r in csv.DictReader(fh)]


def write_csv(rows, path=CSV_PATH):
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=FIELDS, extrasaction="ignore",
                                lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def merge_rows(existing, date_str, new_rows):
    """Drop any earlier rows for this date, prepend the fresh ones, then sort
    newest date first (and biggest mover first within a date)."""
    kept = [r for r in existing if r.get("date") != date_str]
    merged = new_rows + kept
    merged.sort(key=lambda r: (r.get("date", ""), float(r.get("change_pct") or 0)),
                reverse=True)
    return merged


def looks_like_repeat(existing, new_rows):
    """Guard against recording the same session twice - on an NSE holiday the
    scanner keeps serving the previous session's numbers."""
    if not existing or not new_rows:
        return False
    last_date = max(r.get("date", "") for r in existing)
    prev = {(r.get("symbol"), r.get("change_pct"))
            for r in existing if r.get("date") == last_date}
    curr = {(r.get("symbol"), r.get("change_pct")) for r in new_rows}
    return prev == curr


# --------------------------------------------------------------- telegram
def _creds():
    return os.environ.get("TELEGRAM_BOT_TOKEN"), os.environ.get("TELEGRAM_CHAT_ID")


def send_document(token, chat_id, csv_path, caption):
    with csv_path.open("rb") as fh:
        r = requests.post(f"https://api.telegram.org/bot{token}/sendDocument",
                          data={"chat_id": chat_id, "caption": caption},
                          files={"document": (csv_path.name, fh, "text/csv")},
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
    csv_path = Path(args.csv)
    existing = read_csv(csv_path)

    if not args.force and looks_like_repeat(existing, new_rows):
        return {"status": "repeat",
                "message": f"<b>Price movers - {date_str}</b>\n"
                           f"Data identical to the previous session - treating "
                           f"today as a non-trading day. No rows written."}

    merged = merge_rows(existing, date_str, new_rows)
    write_csv(merged, csv_path)
    dates = len({r["date"] for r in merged})
    print(f"{date_str}: wrote {len(new_rows)} rows; file now holds "
          f"{len(merged)} rows across {dates} dates")

    caption = f"Price movers {date_str} - {len(new_rows)} rows"
    if new_rows:
        top = new_rows[:5]
        lines = [f"<b>Price movers - {date_str}</b>",
                 f"{len(new_rows)} NSE stocks up 5% or more", ""]
        for r in top:
            lines.append(f"{r['symbol']}  +{r['change_pct']}%  "
                         f"Rs {r['close']}  rvol {r['rvol']}")
        if len(new_rows) > len(top):
            lines.append(f"...and {len(new_rows) - len(top)} more in the CSV")
    else:
        lines = [f"<b>Price movers - {date_str}</b>",
                 "No NSE stock met the screen today."]

    return {"status": "ok", "caption": caption, "message": "\n".join(lines),
            "rows": len(new_rows), "total": len(merged), "csv": csv_path}


def main(argv=None):
    ap = argparse.ArgumentParser(description="NSE upside movers from TradingView")
    ap.add_argument("--date", help="run date, YYYY-MM-DD (default: today IST)")
    ap.add_argument("--csv", default=str(CSV_PATH), help="CSV file to update")
    ap.add_argument("--fixture", help="read a saved scanner response instead of "
                                      "calling TradingView (offline testing)")
    ap.add_argument("--no-telegram", action="store_true",
                    help="skip the Telegram post")
    ap.add_argument("--force", action="store_true",
                    help="run even on a weekend / holiday / repeat snapshot")
    args = ap.parse_args(argv)

    day = (datetime.strptime(args.date, "%Y-%m-%d").date()
           if args.date else datetime.now(IST).date())
    date_str = day.isoformat()
    token, chat_id = _creds()

    stage = "start"
    try:
        stage = "screen"
        result = run(args, day, date_str)

        if args.no_telegram:
            print("Telegram post skipped (--no-telegram).")
            return 0
        if not token or not chat_id:
            print("TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID not set - "
                  "nothing posted.", file=sys.stderr)
            return 1

        stage = "telegram"
        if result["status"] == "ok":
            send_document(token, chat_id, result["csv"], result["caption"])
            send_message(token, chat_id, result["message"])
        else:
            send_message(token, chat_id, result["message"])
        print("Telegram post sent.")
        return 0

    except Exception as exc:                           # noqa: BLE001
        traceback.print_exc()
        report_failure(token, chat_id, date_str, stage, exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
