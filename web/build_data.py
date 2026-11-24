"""Rebuild web/data.js from ../ie_data.xlsx (Shiller Data sheet) + the latest S&P 500 price from Yahoo Finance.

Usage (from the invest_tops folder):
    python3 convert_data.py ~/Downloads/"ie_data (N).xls"   # .xls -> ie_data.xlsx  (when Shiller posts a new file)
    python3 web/build_data.py                               # ie_data.xlsx + Yahoo -> web/data.js
    python3 web/build_data.py --no-live                     # Shiller only (offline)
then deploy:  cd web && vercel deploy --prod

Live-price step: pulls the latest S&P 500 (^GSPC) price from Yahoo Finance (no key needed);
  during market hours that is today's price, otherwise the last close.
  - Same month as Shiller's last row -> replaces that row's price with the close.
  - The month right after Shiller's last row -> appends one new row.
  - Anything older/further out -> skipped with a warning (refresh the Shiller file).
Real total return is rescaled with Shiller's own formula; dividends and CPI carry forward.
If Yahoo can't be reached the build still runs on Shiller alone and says so.
"""
import json
import sys
import urllib.request
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
XLSX = HERE.parent / "ie_data.xlsx"
YAHOO_URL = "https://query1.finance.yahoo.com/v8/finance/chart/%5EGSPC?range=5d&interval=1d"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# columns A..K: date, price, div, earnings, CPI, date frac, GS10, real price, real div, real TR price, real earnings
DATE, PRICE, DIV, EARN, CPI, FRAC, GS10, RPRICE, RDIV, RTR, REARN = range(11)


def ym(code):
    """Shiller date code (2026.1 = Oct, 2026.11 = Nov) -> (year, month)."""
    y = int(code)
    return y, round((code - y) * 100)


def code_of(y, m):
    return round(y + m / 100, 2)


def load_shiller():
    wb = openpyxl.load_workbook(str(XLSX), data_only=True)
    rows = []
    for row in wb["Data"].iter_rows(min_row=9, values_only=True):
        if row[0] is None:
            continue
        rows.append([v if isinstance(v, (int, float)) else None for v in row[:11]])
    return rows


def latest_yahoo_price():
    """Latest S&P 500 (^GSPC) price from Yahoo Finance: today's price while the market is open, else the last close."""
    req = urllib.request.Request(YAHOO_URL, headers={"User-Agent": "Mozilla/5.0"})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                meta = json.load(resp)["chart"]["result"][0]["meta"]
            break
        except OSError:
            if attempt == 2:
                raise
    price, ts = meta.get("regularMarketPrice"), meta.get("regularMarketTime")
    if not price or not ts:
        raise ValueError("Yahoo returned no price")
    when = datetime.fromtimestamp(ts, ZoneInfo("America/New_York"))
    return when.date(), float(price)


def last_known(rows, col):
    return next(r[col] for r in reversed(rows) if r[col] is not None)


def splice(rows, close_date, close):
    """Put the latest price into the last month. Returns a note string, or None if skipped."""
    last = rows[-1]
    ly, lm = ym(last[DATE])
    nxt = (ly + (lm == 12), lm % 12 + 1)
    month = (close_date.year, close_date.month)

    if month == (ly, lm):  # replace the price in Shiller's (partial) last month
        prev = rows[-2]
        old = last[PRICE]
        last[PRICE] = close
        old_rp = last[RPRICE]
        last[RPRICE] = close * old_rp / old  # same CPI base as Shiller's row
        # Shiller: realTR_t = realTR_{t-1} * (realP_t + realD_t/12) / realP_{t-1}.
        # Only the price term changes, so shift his own figure (keeps his dividend estimate exactly).
        last[RTR] += prev[RTR] * (last[RPRICE] - old_rp) / prev[RPRICE]
        return f"replaced {MONTHS[lm-1]} {ly} price {old:,.2f} -> {close:,.2f}"

    if month == nxt:  # append the next month, carrying CPI and dividends forward
        cpi = last[CPI]
        rdiv = last_known(rows, RDIV)
        rp = close  # CPI carried forward = same base as Shiller's last row
        new = [None] * 11
        new[DATE] = code_of(*nxt)
        new[PRICE] = close
        new[CPI] = cpi
        new[FRAC] = nxt[0] + (nxt[1] - 0.5) / 12
        new[GS10] = last[GS10]
        new[RPRICE] = rp
        new[RTR] = last[RTR] * (rp + rdiv / 12) / last[RPRICE]
        rows.append(new)
        return f"appended {MONTHS[nxt[1]-1]} {nxt[0]} at {close:,.2f}"

    print(f"WARNING: latest price is {close_date}, Shiller ends {MONTHS[lm-1]} {ly}. "
          "Too far apart to splice; refresh ie_data.xlsx. Using Shiller only.")
    return None


def main():
    rows = load_shiller()
    as_of = None
    if "--no-live" not in sys.argv:
        try:
            d, v = latest_yahoo_price()
            note = splice(rows, d, v)
            if note:
                as_of = f"{MONTHS[d.month-1]} {d.day}, {d.year}"
                print(f"Yahoo S&P 500 {d} = {v:,.2f}: {note}")
        except Exception as e:  # network down, Yahoo format change, etc.
            print(f"WARNING: Yahoo fetch failed ({e}). Using Shiller only.")

    (HERE / "data.js").write_text(
        "// Generated by build_data.py from ie_data.xlsx (+ latest Yahoo S&P 500 price). Do not edit by hand.\n"
        "window.SHILLER = " + json.dumps(rows, separators=(",", ":")) + ";\n"
        "window.PRICE_AS_OF = " + json.dumps(as_of) + ";\n"
    )
    print(f"wrote {len(rows)} rows, last date {rows[-1][0]}, price as of {as_of or 'Shiller month'}")


if __name__ == "__main__":
    main()
