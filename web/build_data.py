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
YAHOO_DAILY = "https://query1.finance.yahoo.com/v8/finance/chart/%5EGSPC?period1=-1325635200&period2=9999999999&interval=1d"
DAILY_CACHE = HERE.parent / "sp500_daily.csv"  # offline fallback, refreshed every successful build

# Paul's bear-market rule (reproduces his 10 since 1984 from daily closes):
#   a bear starts once the S&P 500 closes 18% below its running high;
#   it ends at a new high, or once the index has rallied 25% off the bear's low.
BEAR_DROP, BEAR_RALLY, HALF = 0.18, 0.25, 0.45  # HALF: drops of 45%+ count as "cut roughly in half"
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


def daily_closes():
    """Every S&P 500 daily close since Dec 1927 from Yahoo, as [(YYYY-MM-DD, close)]. Falls back to the local cache."""
    try:
        req = urllib.request.Request(YAHOO_DAILY, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=60) as resp:
            res = json.load(resp)["chart"]["result"][0]
        ny = ZoneInfo("America/New_York")
        out = [(datetime.fromtimestamp(t, ny).strftime("%Y-%m-%d"), round(c, 2))
               for t, c in zip(res["timestamp"], res["indicators"]["quote"][0]["close"]) if c]
        DAILY_CACHE.write_text("date,close\n" + "".join(f"{d},{c}\n" for d, c in out))
        return out
    except Exception as e:
        print(f"WARNING: Yahoo daily history failed ({e}); using {DAILY_CACHE.name}")
        lines = DAILY_CACHE.read_text().splitlines()[1:]
        return [(l.split(",")[0], float(l.split(",")[1])) for l in lines]


def find_bears(series):
    """series = [(label, price)] in time order. Returns [(peak_label, peak, trough_label, trough, ended)]."""
    out, peak, trough, inbear = [], series[0], None, False
    for d, c in series:
        if not inbear:
            if c > peak[1]:
                peak = (d, c)
            elif c <= peak[1] * (1 - BEAR_DROP):
                inbear, trough = True, (d, c)
        else:
            if c < trough[1]:
                trough = (d, c)
            elif c >= peak[1] or c >= trough[1] * (1 + BEAR_RALLY):
                out.append((peak[0], peak[1], trough[0], trough[1], True))
                inbear, peak = False, (d, c)
    if inbear:
        out.append((peak[0], peak[1], trough[0], trough[1], False))
    # One long slide, not several: if the next bear sinks below the last low before the market
    # ever regains the earlier high, it is the same bear (1929-32 is one -86% bear, not four).
    merged = []
    for b in out:
        if merged and b[3] < merged[-1][3] and b[1] < merged[-1][1]:
            prev = merged.pop()
            b = (prev[0], prev[1], b[2], b[3], b[4])
        merged.append(b)
    return merged


def bear_tables(rows, daily):
    """For every start month: the bear markets an investor starting that month lived through.
    Pre-1928 months use Shiller's monthly prices until daily data begins (Jan 1928)."""
    monthly = [(f"{ym(r[DATE])[0]:04d}-{ym(r[DATE])[1]:02d}-15", r[PRICE]) for r in rows
               if r[PRICE] and ym(r[DATE])[0] < 1928]
    table, index, by_month = [], {}, {}
    for r in rows:
        y, m = ym(r[DATE])
        start = f"{y:04d}-{m:02d}"
        series = [x for x in monthly if x[0] >= start] + [x for x in daily if x[0] >= max(start, "1928-01")]
        ids = []
        for b in find_bears(series):
            key = b[:4]
            if key not in index:
                index[key] = len(table)
                table.append([b[0], b[2], round((b[3] / b[1] - 1) * 100, 1), b[4]])
            ids.append(index[key])
        by_month[f"{r[DATE]:.2f}"] = ids
    return table, by_month


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

    bears, bear_by_month = bear_tables(rows, daily_closes())
    (HERE / "data.js").write_text(
        "// Generated by build_data.py from ie_data.xlsx (+ latest Yahoo S&P 500 price). Do not edit by hand.\n"
        "window.SHILLER = " + json.dumps(rows, separators=(",", ":")) + ";\n"
        "window.PRICE_AS_OF = " + json.dumps(as_of) + ";\n"
        f"// Bear markets (drop {BEAR_DROP:.0%}+, ends at new high or {BEAR_RALLY:.0%} rally): [peak date, low date, % drop, ended]\n"
        "window.BEARS = " + json.dumps(bears, separators=(",", ":")) + ";\n"
        "window.BEAR_BY_MONTH = " + json.dumps(bear_by_month, separators=(",", ":")) + ";\n"
        f"window.BEAR_RULE = {json.dumps({'drop': BEAR_DROP, 'rally': BEAR_RALLY, 'half': HALF})};\n"
    )
    print(f"wrote {len(rows)} rows, last date {rows[-1][0]}, price as of {as_of or 'Shiller month'}")


if __name__ == "__main__":
    main()
