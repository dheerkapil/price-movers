# price-movers

Daily list of NSE stocks that moved **+5% or more**, filtered down to names worth
watching — pulled entirely from the TradingView scanner feed, appended to an Excel workbook,
and posted to Telegram every trading day at **16:00 IST**.

## The screen

Every condition must hold on the day's close:

| Condition | Value | TradingView field |
|---|---|---|
| Exchange | NSE only | `exchange` |
| Day change | **≥ +5.0%** (upside only) | `change` |
| Average volume, 10 days | **> 500,000** shares | `average_volume_10d_calc` |
| Market cap | **> ₹100 crore** | `market_cap_basic` |
| Last price | **> ₹19** | `close` |

The filter runs server-side at TradingView, and every threshold is re-checked
locally before a row is written, so the numbers in the CSV are exactly the ones
above.

## The output

`price_movers.csv` lives in the repo and grows by one date block per session. It is
the record: plain text, cheap to append in git, and readable as a diff.

The Excel workbook (`price_movers.xlsx`) is built fresh at send time from that CSV
and is **not** committed — every save rewrites the whole archive, so storing it
daily would cost far more history than the CSV it came from. It exists only as the
Telegram attachment.
Newest date first, biggest mover first within a date:

| Column | Meaning |
|---|---|
| `date` | session date (YYYY-MM-DD, IST) |
| `symbol` | NSE trading symbol |
| `fyers_code` | broker code — `NSE:<symbol>-<series>, ` (trailing comma + space), e.g. `NSE:RELIANCE-EQ, ` |
| `company` | company name |
| `close` | closing price, ₹ |
| `change_pct` | % change vs previous close |
| `volume` | shares traded that day |
| `avg_volume_10d` | 10-day average volume, shares |
| `rvol` | relative volume — today's volume ÷ 10-day average |
| `market_cap_cr` | market cap, ₹ crore |
| `exchange` | always `NSE` |
| `series` | current NSE series — EQ, BE, BZ, SM — refreshed on every run |

## The series column

`series` comes from NSE's securities list (`sec_list.csv`), fetched on every run.
It is **not** the series the stock traded in on that date: every row for a symbol
is rewritten to the symbol's *current* classification. If a stock was BZ ten days
ago, EQ five days ago and BE today, all three of its rows show BE after today's
run. The file always reads as the position today, applied backwards.

Each `fyers_code` ends with a comma and a space, so the column can be
concatenated straight into a broker request without a join step. Because the
value then contains a comma, the CSV writer quotes that one field —
`"NSE:RELIANCE-EQ, "` — which is ordinary CSV and parses back exactly.

`fyers_code` is built from the same lookup: `NSE:`, the official NSE symbol and
the current series. It uses NSE's own spelling, so a stock TradingView writes as
`M_M` comes out as `NSE:M&M-EQ`. If a symbol is missing from NSE's list there is
no code to build, and the cell is left empty rather than guessed.

If the NSE list cannot be fetched, the run continues with the series already
recorded and logs a warning rather than losing the day's data.

## Telegram

Every run ends with **exactly one** Telegram message:

- the updated Excel workbook, or
- a failure alert naming the stage that broke and the error, with a link to the
  workflow run, or
- a note that there was no NSE session (weekend, holiday) or that the data was
  identical to the previous session.

## Setup

The workflow needs two repository secrets:

| Secret | Value |
|---|---|
| `TELEGRAM_BOT_TOKEN` | token of the `Price_alerts` bot |
| `TELEGRAM_CHAT_ID` | optional — chat id the CSV should be posted to |

Only the bot token is required. If `TELEGRAM_CHAT_ID` is not set, the script
asks Telegram for the bot's recent updates and uses the chat that last messaged
it, so sending the bot a single message once is enough.

Add them under **Settings → Secrets and variables → Actions → New repository
secret**.

## Schedule

`.github/workflows/daily.yml` runs at `30 10 * * 1-5` — 10:30 UTC, which is
16:00 IST, Monday to Friday. GitHub's scheduler can start a run a few minutes
late under load, and it does not know about NSE holidays; the script handles
those itself. `workflow_dispatch` runs it on demand.

## Running it locally

```bash
pip install -r requirements.txt
export TELEGRAM_BOT_TOKEN=...   # optional
export TELEGRAM_CHAT_ID=...     # optional

python screener.py                       # today, live from TradingView
python screener.py --no-telegram         # skip the post
python screener.py --date 2026-11-06     # backfill a specific date's screen
python screener.py --fixture tests/sample_scan.json --no-telegram   # offline
```

Useful flags: `--no-telegram`, `--fixture FILE`, `--force` (ignore the
weekend/holiday/repeat-snapshot guards).

## Known limits

- **No history.** The scanner returns a live snapshot and takes no date
  parameter, so this cannot reconstruct past screens. Rows only accumulate from
  the day the workflow first runs.
- **Market cap is live-only** for the same reason; it is TradingView's current
  figure for the stock.
- **Holidays.** Fixed-date NSE holidays are hardcoded in `screener.py`. The
  variable ones (Holi, Diwali, Id, Good Friday and so on) need adding each
  January when NSE publishes the next calendar — otherwise the repeat-snapshot
  guard is what stops a holiday from being recorded as a session.
- **One source.** All values, including rvol and market cap, come from
  TradingView. If the scanner is rate-limited or changes its field names, the
  run fails loudly rather than writing partial data.

Not investment advice — a screening tool.
