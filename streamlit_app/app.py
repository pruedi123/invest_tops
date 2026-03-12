import streamlit as st
import openpyxl
import os
import json
from pathlib import Path

# ── Page config ──
st.set_page_config(
    page_title="What If You Bought At The Very Top?",
    page_icon="📈",
    layout="wide",
)

# ── Custom CSS matching the original design ──
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,600;0,700;1,400;1,500&family=DM+Mono:wght@300;400;500&display=swap');

:root {
    --bg: #f5f2ed;
    --gold: #7a5c08;
    --green: #155c38;
    --teal: #0e5c5c;
    --blue: #1a3f7a;
    --red: #8b1a1a;
    --text: #2a2a28;
    --muted: #6b6860;
    --card-bg: #fffef9;
    --border: #d5d0c6;
}

.stApp { background-color: var(--bg); }
.stApp header { background-color: var(--bg); }

.hero-title {
    font-family: 'Cormorant Garamond', serif;
    font-size: clamp(36px, 6vw, 56px);
    font-weight: 700;
    line-height: 1.1;
    color: var(--text);
    text-align: center;
    margin-bottom: 16px;
}
.hero-title em { color: var(--gold); font-style: italic; }

.hero-subtitle {
    font-family: 'Cormorant Garamond', serif;
    font-size: 18px;
    line-height: 1.6;
    color: var(--muted);
    text-align: center;
    font-style: italic;
    max-width: 720px;
    margin: 0 auto 24px;
}

.eyebrow {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: var(--teal);
    text-align: center;
    margin-bottom: 12px;
}

.section-label {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: var(--teal);
    text-align: center;
    margin: 40px 0 20px;
}

.method-box {
    background: var(--card-bg);
    border: 2px solid var(--teal);
    border-radius: 10px;
    padding: 28px;
    margin: 0 auto 32px;
}
.method-box h3 {
    font-family: 'Cormorant Garamond', serif;
    font-size: 22px;
    color: var(--teal);
    margin-bottom: 12px;
}
.method-box p {
    font-family: 'DM Mono', monospace;
    font-size: 13px;
    line-height: 1.7;
    color: var(--muted);
}
.formula-box {
    background: var(--bg);
    border-radius: 6px;
    padding: 16px;
    font-family: 'DM Mono', monospace;
    font-size: 12px;
    line-height: 1.8;
    color: var(--muted);
}

.cpi-tile {
    background: var(--card-bg);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 14px 10px;
    text-align: center;
}
.cpi-date {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: var(--muted);
    margin-bottom: 4px;
}
.cpi-amt {
    font-family: 'DM Mono', monospace;
    font-size: 17px;
    font-weight: 500;
    color: var(--teal);
}
.cpi-eq {
    font-family: 'DM Mono', monospace;
    font-size: 10px;
    color: var(--muted);
    margin-top: 2px;
}

.fear-bar {
    background: var(--card-bg);
    border: 2px solid var(--red);
    border-radius: 10px;
    padding: 24px 28px;
    margin: 0 auto 32px;
}
.fear-header {
    font-family: 'DM Mono', monospace;
    font-size: 13px;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: var(--red);
    font-weight: 700;
    margin-bottom: 8px;
}
.fear-text {
    font-family: 'Cormorant Garamond', serif;
    font-size: 16px;
    line-height: 1.6;
    color: var(--muted);
}
.fear-badge {
    border: 2px solid var(--red);
    border-radius: 8px;
    background: #fff;
    padding: 14px 18px;
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    line-height: 1.5;
    color: var(--red);
    text-align: center;
    font-weight: 500;
    margin-top: 16px;
}

.scenario-card {
    background: var(--card-bg);
    border: 1px solid var(--border);
    border-radius: 10px;
    margin-bottom: 12px;
    overflow: hidden;
}
.card-header-row {
    display: flex;
    align-items: center;
    padding: 18px 24px;
    gap: 16px;
}
.card-num {
    font-family: 'Cormorant Garamond', serif;
    font-size: 28px;
    font-weight: 700;
    color: var(--border);
    width: 36px;
    flex-shrink: 0;
}
.card-title {
    font-family: 'Cormorant Garamond', serif;
    font-size: 18px;
    font-weight: 600;
    color: var(--text);
}
.card-peak {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    color: var(--muted);
}
.card-val-label {
    font-family: 'DM Mono', monospace;
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: var(--muted);
}
.card-invested-val {
    font-family: 'DM Mono', monospace;
    font-size: 14px;
    color: var(--teal);
    font-weight: 500;
}
.card-result-val {
    font-family: 'DM Mono', monospace;
    font-size: 20px;
    color: var(--gold);
    font-weight: 500;
}

.conversion-note {
    background: rgba(14,92,92,.06);
    border-radius: 6px;
    padding: 14px 18px;
    font-family: 'DM Mono', monospace;
    font-size: 14px;
    color: var(--teal);
    line-height: 1.7;
    margin: 14px 0;
}

.narrative {
    font-family: 'Cormorant Garamond', serif;
    font-size: 18px;
    line-height: 1.7;
    color: var(--muted);
    font-style: italic;
    margin: 14px 0;
}

.pain-box {
    background: rgba(139,26,26,.04);
    border-left: 3px solid var(--red);
    padding: 14px 18px;
    font-family: 'DM Mono', monospace;
    font-size: 14px;
    color: var(--red);
    line-height: 1.7;
    border-radius: 0 6px 6px 0;
    margin: 14px 0;
}

.era-quote {
    background: var(--card-bg);
    border-left: 3px solid var(--gold);
    padding: 16px 20px;
    margin: 14px 0;
    border-radius: 0 8px 8px 0;
}
.era-quote blockquote {
    font-family: 'Cormorant Garamond', serif;
    font-size: 19px;
    line-height: 1.5;
    color: var(--text);
    font-style: italic;
    margin: 0 0 8px 0;
    padding: 0;
}
.era-quote .quote-attr {
    font-family: 'DM Mono', monospace;
    font-size: 13px;
    color: var(--muted);
    letter-spacing: 0.5px;
}

.metric-card {
    background: var(--bg);
    border-radius: 8px;
    padding: 14px;
    text-align: center;
}
.metric-label {
    font-family: 'DM Mono', monospace;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: var(--muted);
    margin-bottom: 4px;
}
.metric-val {
    font-family: 'DM Mono', monospace;
    font-size: 20px;
    font-weight: 500;
    color: var(--text);
}
.metric-val.gold { color: var(--gold); font-size: 24px; }
.metric-val.teal { color: var(--teal); }
.metric-val.green { color: var(--green); }

.bar-container { margin: 16px 0; }
.bar-row {
    display: flex;
    align-items: center;
    margin-bottom: 6px;
    gap: 10px;
}
.bar-label-text {
    font-family: 'DM Mono', monospace;
    font-size: 13px;
    color: var(--muted);
    width: 110px;
    text-align: right;
    flex-shrink: 0;
}
.bar-track {
    flex: 1;
    height: 22px;
    background: var(--bg);
    border-radius: 4px;
    overflow: hidden;
}
.bar-fill {
    height: 100%;
    border-radius: 4px;
    transition: width 1s ease;
}
.bar-fill-gold { background: var(--gold); }
.bar-fill-green { background: var(--green); }
.bar-fill-teal { background: var(--teal); }
.bar-fill-muted { background: var(--border); }
.bar-val-text {
    font-family: 'DM Mono', monospace;
    font-size: 13px;
    color: var(--text);
    width: 110px;
    font-weight: 500;
}

.closing-quote {
    font-family: 'Cormorant Garamond', serif;
    font-size: 22px;
    line-height: 1.5;
    color: var(--text);
    text-align: center;
    font-weight: 600;
    margin-bottom: 16px;
}
.closing-tagline {
    font-family: 'Cormorant Garamond', serif;
    font-size: 18px;
    color: var(--gold);
    text-align: center;
    font-style: italic;
}

.footnote-text {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    color: var(--muted);
    text-align: center;
    line-height: 1.7;
    margin-top: 40px;
}

/* Single date analysis */
.analysis-table {
    font-family: 'DM Mono', monospace;
    font-size: 13px;
    width: 100%;
    border-collapse: collapse;
}
.analysis-table th {
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: var(--muted);
    border-bottom: 2px solid var(--border);
    padding: 8px 10px;
    text-align: left;
}
.analysis-table td {
    padding: 8px 10px;
    border-bottom: 1px solid var(--border);
    font-family: 'DM Mono', monospace;
}

/* Summary table */
.summary-table {
    width: 100%;
    border-collapse: collapse;
    font-family: 'DM Mono', monospace;
    font-size: 12px;
}
.summary-table th {
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: var(--muted);
    border-bottom: 2px solid var(--border);
    padding: 8px 10px;
    text-align: left;
}
.summary-table td {
    padding: 8px 10px;
    border-bottom: 1px solid var(--border);
}
.summary-table .col-gold { color: var(--gold); font-weight: 500; }
.summary-table .col-teal { color: var(--teal); }

/* Hide streamlit defaults */
.stDeployButton { display: none; }
div[data-testid="stDecoration"] { display: none; }
</style>
""", unsafe_allow_html=True)


# ── Data loading ──
DATA_DIR = Path(__file__).resolve().parent.parent
EXCEL_PATH = DATA_DIR / "ie_data.xlsx"


@st.cache_data
def load_shiller():
    wb = openpyxl.load_workbook(str(EXCEL_PATH), data_only=True)
    ws = wb["Data"]
    rows = []
    for row in ws.iter_rows(min_row=9, values_only=True):
        if row[0] is None:
            continue
        rows.append(row)
    return rows


def find_row(rows, target):
    best, best_diff = None, 9999
    for r in rows:
        diff = abs(float(r[0]) - target)
        if diff < best_diff:
            best_diff = diff
            best = r
    return best


def find_last_complete(rows):
    for r in reversed(rows):
        if r[0] and r[1] and r[4] and r[9]:
            return r


def find_last_with_dividends(rows):
    for r in reversed(rows):
        if r[0] and r[1] and r[2] and r[4] and r[7] and r[8] and r[9]:
            return r


MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]


def date_str(d):
    year = int(d)
    mon = max(1, min(12, round((d - year) * 100)))
    return f"{MONTHS[mon-1]} {year}"


def analyze(start_date_float, investment, rows):
    start = find_row(rows, start_date_float)
    end = find_last_complete(rows)
    div_end = find_last_with_dividends(rows) or end

    cpi_s = float(start[4]); cpi_e = float(end[4])
    rtr_s = float(start[9]); rtr_e = float(end[9])
    rp_s = float(start[7]); rp_e = float(end[7])

    scaled = investment * (cpi_s / cpi_e)
    cpi_factor = cpi_e / cpi_s
    real_tr_f = rtr_e / rtr_s
    nom_tr_f = real_tr_f * cpi_factor
    real_end = scaled * real_tr_f
    nom_end = scaled * nom_tr_f
    price_r_f = rp_e / rp_s
    price_only = scaled * price_r_f

    div_nom_f = div_real_f = div_start_r = div_end_r = None
    if start[2] and div_end[2] and start[8] and div_end[8]:
        dn_s = float(start[2]); dn_e = float(div_end[2])
        dr_s = float(start[8]); dr_e = float(div_end[8])
        div_nom_f = dn_e / dn_s
        div_real_f = dr_e / dr_s
        div_start_r = dr_s
        div_end_r = dr_e

    sy = int(start[0]); sm = round((float(start[0]) - sy) * 100)
    ey = int(end[0]); em = round((float(end[0]) - ey) * 100)
    years = (ey - sy) + (em - sm) / 12

    return {
        "start_str": date_str(float(start[0])),
        "end_str": date_str(float(end[0])),
        "start_date": float(start[0]),
        "investment": investment,
        "scaled": round(scaled, 0),
        "cpi_start": round(cpi_s, 3),
        "cpi_end": round(cpi_e, 3),
        "cpi_factor": round(cpi_factor, 2),
        "real_tr_f": round(real_tr_f, 2),
        "nom_tr_f": round(nom_tr_f, 1),
        "real_end": round(real_end, 0),
        "nom_end": round(nom_end, 0),
        "price_r_f": round(price_r_f, 2),
        "price_only": round(price_only, 0),
        "div_nom_f": round(div_nom_f, 1) if div_nom_f else None,
        "div_real_f": round(div_real_f, 1) if div_real_f else None,
        "div_start_r": round(div_start_r, 2) if div_start_r else None,
        "div_end_r": round(div_end_r, 2) if div_end_r else None,
        "years": int(years),
    }


def fmt(n):
    if n is None:
        return "—"
    if abs(n) >= 1e9:
        return f"${n/1e9:.1f}B"
    if abs(n) >= 1e6:
        return f"${n/1e6:.1f}M"
    if abs(n) >= 1e3:
        return f"${round(n):,}"
    return f"${n:.0f}"


def fmt_x(n):
    if n is None:
        return "—"
    if n >= 100:
        return f"{n:.0f}x"
    if n >= 10:
        return f"{n:.1f}x"
    return f"{n:.2f}x"


# ── Scenario definitions ──
SCENARIOS = [
    {"num": 1, "label": "The Great Crash", "short": "Great Crash", "date": 1929.09, "crash": "−89%", "peak_desc": "Peak before −89% collapse",
     "narrative": 'Just days before the crash, Irving Fisher declared that stocks had reached "a permanently high plateau." The Dow didn\'t reclaim its 1929 high until 1954 — 25 years later. And yet, with dividends reinvested, the patient investor went on to extraordinary wealth.',
     "pain": "Market fell 89% peak-to-trough. By 1932, the investment had shrunk to ~20¢ on the dollar. Most investors sold in panic and locked in permanent losses.",
     "quote": "I now see nothing to give ground to hope — nothing of man.",
     "quote_attr": "Calvin Coolidge, former U.S. President — 1933, near the market bottom"},
    {"num": 2, "label": "Depression-Era Rally Peak", "short": "Depression Rally", "date": 1937.02, "crash": "−60%", "peak_desc": "Peak before −60% plunge",
     "narrative": "By early 1937, GDP had returned to 1929 levels. The Dow had recovered 80% from the lows. At his second inaugural on January 20, FDR declared: \"Our progress out of the depression is obvious.\" Then the Fed doubled reserve requirements, FDR cut spending to balance the budget, and stocks crashed 60%. Investors who had just survived the Great Crash were crushed again.",
     "pain": "Two consecutive crashes within a decade. Industrial production dropped 37%. Unemployment surged from 14% back to 19%. Congressman Maury Maverick said: \"We have pulled all the rabbits out of the hat, and there are no more rabbits.\" Even Treasury Secretary Morgenthau privately admitted: \"We have tried spending money... it does not work.\" It took until 1945 to break even on price alone.",
     "quote": "There is no real and fundamental basis upon which to build enduring prosperity. We are now helplessly floundering.",
     "quote_attr": "Commercial &amp; Financial Chronicle — Apr 2, 1938, the exact month of the market bottom"},
    {"num": 3, "label": "Go-Go Era Peak", "short": "Go-Go Era", "date": 1968.11, "crash": "−36%", "peak_desc": "Peak before the stagflation decade",
     "narrative": 'The "Nifty Fifty" era — glamour stocks at any price. Then oil shocks, Vietnam, Nixon\'s resignation, and 15% inflation. The 1970s were the worst decade for real returns in market history.',
     "pain": "Real inflation-adjusted losses persisted for over a decade. By 1982, purchasing power was still negative vs. the 1968 entry. A brutal 14-year grind.",
     "quote": "We live in an investment world, populated not by those who must be logically persuaded to believe, but by the hopeful, credulous and greedy, grasping for an excuse to believe.",
     "quote_attr": "Warren Buffett, Partnership Letter — Jan 1968"},
    {"num": 4, "label": "Pre-Oil Shock Peak", "short": "Oil Shock", "date": 1973.01, "crash": "−48%", "peak_desc": "Peak before −48% bear market",
     "narrative": "In January 1973, Barron's annual Roundtable ran the headline \"Not a Bear Among Them\" — every single panelist was bullish. The Nifty Fifty traded at an average P/E of 42x; Polaroid at 95x, McDonald's at 86x. They were called \"one-decision stocks: buy and never sell.\" Then OPEC's oil embargo sent inflation to 12%, stocks crashed 48%, and the Nifty Fifty darlings were, as Forbes later put it, \"taken out and shot one by one.\" Avon fell from $140 to $18.50.",
     "pain": "A vicious double-whammy: portfolio down nearly half while prices of everything else doubled. BusinessWeek wrote: \"The stream of equity capital to US industry has run dry.\" Forbes ran a headline asking \"Dow Below 400?\" — speculating it could fall another 30% from already devastating levels.",
     "quote": "For me, it was like the Great Depression. Everything we owned went down. It seemed as if the world was coming to an end.",
     "quote_attr": "Chuck Royce, Pennsylvania Mutual Fund — 1974, near the market bottom"},
    {"num": 5, "label": "Death of Equities", "short": "Death of Equities", "date": 1979.08, "crash": "N/A", "peak_desc": "Not a peak — but nothing made you want to own stocks",
     "narrative": "This wasn't a market top. It was something worse: the moment America gave up on stocks entirely. Seven million shareholders had defected. The money went elsewhere — gold surged from $200 to $850, diamonds boomed, and pension funds piled into hard assets. Bestselling author Howard Ruff warned of \"runaway inflation\" and told everyone to buy gold and silver \"forever.\" Barron's marveled that gold bug James Dines' prediction that bullion would cross the Dow \"begins to look like one of the most fantastic investment calls on record.\" The Dow sat at 875. Within three years, gold had crashed 60% and the greatest bull market of the 20th century began.",
     "pain": "No crash followed — but that's the point. Stocks had been punishing investors for over a decade. Real returns since 1968 were negative. BusinessWeek wrote: \"Only the elderly who have not understood the changes in the nation's financial markets are sticking with stocks.\" The pain was already priced in — and then some.",
     "quote": "The death of equities is a near-permanent condition — reversible someday, but not soon. The old attitude of buying solid stocks as a cornerstone for one's life savings and retirement has simply disappeared.",
     "quote_attr": "BusinessWeek — Aug 13, 1979. The S&amp;P 500 returned 17.6% annualized over the next 20 years."},
    {"num": 6, "label": "Black Monday Peak", "short": "Black Monday", "date": 1987.08, "crash": "−34%", "peak_desc": "Peak before −34% crash in 77 days",
     "narrative": "In September 1987, Robert Prechter — the most famous market guru of the decade — predicted the Dow would reach 3,600 from 2,660. Six weeks later, on October 19, the Dow fell 22.6% in a single day — the largest one-day drop in history. The Wall Street Journal's front page the next morning drew direct comparisons to 1929. The market fully recovered within 2 years.",
     "pain": None,
     "quote": "Does 1987 Equal 1929?",
     "quote_attr": "New York Times, front page — Oct 20, 1987"},
    {"num": 7, "label": "Dot-Com Bubble Peak", "short": "Dot-Com", "date": 2000.01, "crash": "−49%", "peak_desc": "Peak before −49% crash over 2.5 years",
     "narrative": "In September 1999, James Glassman and Kevin Hassett published Dow 36,000 — arguing stocks were undervalued by 350%. Fortune later called it \"the most spectacularly wrong investing book ever.\" The NASDAQ fell 78%. Individual tech stocks lost 90–99%. Valuation metrics were said to be obsolete. They weren't. The hardest entry point on this list — yet still profitable over 25 years.",
     "pain": "S&P 500 didn't reclaim Jan 2000 levels until 2007 — then immediately crashed again. On a real basis, still underwater in 2012. A 12-year price drought. In July 2002, investors pulled $50 billion from equity mutual funds in a single month — an all-time record.",
     "quote": "Stocks stink and will continue to do so until they're priced appropriately, probably somewhere around Dow 5,000.",
     "quote_attr": "Bill Gross, PIMCO ($270B under management) — Sep 2002, one month before the exact bottom"},
    {"num": 8, "label": "Housing Bubble Peak", "short": "Housing Bubble", "date": 2007.10, "crash": "−57%", "peak_desc": "Peak before −57% financial crisis",
     "narrative": "In March 2007, Fed Chair Ben Bernanke told Congress the subprime crisis \"seems likely to be contained.\" Citigroup CEO Chuck Prince said \"as long as the music is playing, you've got to get up and dance.\" Lehman Brothers collapsed. AIG needed a $182 billion bailout. The banking system nearly failed. Many serious analysts believed capitalism itself was at risk.",
     "pain": "−57% decline. The S&P 500 didn't reclaim Oct 2007 highs until March 2013 — 5.5 years later. Jim Cramer told viewers \"Bear Stearns is not in trouble\" at $62/share. Five days later it sold for $2.",
     "quote": "It is highly likely it goes to 600 or below.",
     "quote_attr": "Nouriel Roubini, Mar 9, 2009 — the exact day of the market bottom"},
    {"num": 9, "label": "Pre-Pandemic Peak", "short": "COVID", "date": 2020.02, "crash": "−34%", "peak_desc": "Peak before −34% COVID crash in 33 days",
     "narrative": "On February 25, six days after the peak, Larry Kudlow told investors the virus was \"pretty close to airtight\" contained and to \"seriously consider buying the dip.\" Fifteen days before the peak, the President declared \"our economy is the best it has ever been.\" Then came the fastest 30%+ decline in market history — 34% in 33 days. Lockdowns, mass unemployment, and an economy in freefall.",
     "pain": "34% decline in 33 days — the fastest crash on this list. David Stockman declared \"Wall Street is toast.\" Guggenheim's Scott Minerd predicted the S&P would fall to 1,200. Nouriel Roubini warned of a \"Greater Depression.\" Jeffrey Gundlach was actively shorting, betting the lows would be retaken. The market fully recovered within 5 months — the fastest recovery in history.",
     "quote": "Hell is coming.",
     "quote_attr": "Bill Ackman, CNBC — Mar 18, 2020, five days before the bottom"},
]

DATE_ALIASES = {
    "End of WW2": 1945.08,
    "Great Depression": 1933.03,
    "Black Monday": 1987.10,
    "Black Monday Peak": 1987.08,
    "Dot-Com Peak": 2000.01,
    "Dot-Com Crash": 2002.10,
    "9/11": 2001.09,
    "GFC / Financial Crisis": 2009.03,
    "COVID Low": 2020.03,
    "COVID Peak": 2020.02,
    "Great Crash / 1929 Peak": 1929.09,
    "Oil Shock": 1973.01,
    "Stagflation": 1968.11,
}


def render_bar(label, value, max_val, color_class):
    pct = (value / max_val * 100) if max_val > 0 else 0
    return f"""
    <div class="bar-row">
        <div class="bar-label-text">{label}</div>
        <div class="bar-track"><div class="bar-fill {color_class}" style="width:{pct:.1f}%"></div></div>
        <div class="bar-val-text">{fmt(value)}</div>
    </div>"""


def render_scenario_card(scenario, result, investment):
    scaled = result["scaled"]
    real_end = result["real_end"]
    nom_end = result["nom_end"]
    price_only = result["price_only"]
    s = scenario

    pain_html = ""
    if s["pain"]:
        pain_html = f'<div class="pain-box"><strong>The Pain</strong><br>{s["pain"]}</div>'

    max_val = max(real_end, nom_end, price_only, scaled)

    bars = render_bar("Nominal TR", nom_end, max_val, "bar-fill-gold")
    bars += render_bar("Real TR ★", real_end, max_val, "bar-fill-green")
    bars += render_bar("Price Only", price_only, max_val, "bar-fill-teal")
    bars += render_bar("Invested", scaled, max_val, "bar-fill-muted")

    return f"""
    <div class="scenario-card">
        <div class="card-header-row">
            <div class="card-num">{s['num']}</div>
            <div style="flex:1;min-width:0">
                <div class="card-title">{s['label']} · {result['start_str']}</div>
                <div class="card-peak">{s['peak_desc']}</div>
            </div>
            <div style="text-align:right;margin-right:8px">
                <div class="card-val-label">Invested</div>
                <div class="card-invested-val">{fmt(scaled)}</div>
            </div>
            <div style="text-align:right">
                <div class="card-val-label">★ Real Value</div>
                <div class="card-result-val">{fmt(real_end)}</div>
            </div>
        </div>
    </div>
    """


# ── Load data ──
rows = load_shiller()
end_row = find_last_complete(rows)
data_end_str = date_str(float(end_row[0]))

# ── Sidebar: data info + refresh ──
with st.sidebar:
    st.markdown(f"**Data through:** {data_end_str}")
    st.markdown(f"**Rows loaded:** {len(rows):,}")
    if st.button("Reload data", help="Clear cache and re-read ie_data.xlsx"):
        load_shiller.clear()
        st.rerun()

# ── Navigation ──
tab1, tab2 = st.tabs(["📈 Bought at the Top", "🔍 Single Date Query"])


# ═══════════════════════════════════════════
#  TAB 1: BOUGHT AT THE TOP
# ═══════════════════════════════════════════
with tab1:
    # Hero
    st.markdown('<div class="eyebrow">S&P 500 · Shiller CAPE Data · Total Return With Dividends Reinvested</div>', unsafe_allow_html=True)
    st.markdown('<div class="hero-title">What If You Bought At The <em>Very Top?</em></div>', unsafe_allow_html=True)
    st.markdown('<div style="text-align:center;font-family:\'Cormorant Garamond\',serif;font-size:16px;color:var(--muted);margin-bottom:24px">By <strong style="color:var(--text)">Paul Ruedi</strong></div>', unsafe_allow_html=True)

    # Investment input — preset buttons use callbacks to set state BEFORE the widget renders
    def set_preset(val):
        st.session_state.investment_top = val

    col_spacer1, col_input, col_spacer2 = st.columns([1, 2, 1])
    with col_input:
        preset_cols = st.columns(5)
        presets = [10_000, 50_000, 100_000, 500_000, 1_000_000]
        preset_labels = ["$10K", "$50K", "$100K", "$500K", "$1M"]
        for i, (col, val, label) in enumerate(zip(preset_cols, presets, preset_labels)):
            with col:
                st.button(label, key=f"preset_{val}", use_container_width=True,
                          on_click=set_preset, args=(val,))
        investment = st.number_input(
            "Investment amount (today's dollars)",
            min_value=1000,
            max_value=100_000_000,
            value=100_000,
            step=10_000,
            format="%d",
            key="investment_top",
        )

    st.markdown(
        f'<div class="hero-subtitle">Every generation believes the market is overvalued and ready to crash. '
        f'Here is what actually happened to someone with the equivalent of <strong>{fmt(investment)}</strong> in today\'s money '
        f'who bought at the absolute peak of every major bubble — and simply held.</div>',
        unsafe_allow_html=True,
    )

    # Compute all 8 scenarios
    results = []
    for s in SCENARIOS:
        r = analyze(s["date"], investment, rows)
        results.append(r)

    # Methodology box
    s5 = results[4]  # Aug 1987
    s5_scaled = investment * (s5["cpi_start"] / s5["cpi_end"])
    st.markdown(f"""
    <div class="method-box">
        <h3>📐 How We Made This Apples-to-Apples</h3>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:24px">
            <div>
                <p>A flat {fmt(investment)} at each peak is misleading. That amount in 1929 would be a fortune; in 2020 it's a modest sum.
                To compare honestly, we scale every investment using CPI so each one represents the <strong>same real purchasing power</strong> as {fmt(investment)} today.</p>
                <p style="margin-top:10px">Every dollar figure on this page is computed live from factor multipliers — change the amount above and every number updates instantly.</p>
            </div>
            <div class="formula-box">
                <strong style="color:var(--teal)">Example — Aug 1987:</strong><br>
                CPI then: {s5['cpi_start']} · CPI now: {s5['cpi_end']}<br>
                {fmt(investment)} × ({s5['cpi_start']} ÷ {s5['cpi_end']:.1f})<br>
                = <strong style="color:var(--teal)">{fmt(s5_scaled)}</strong> invested in 1987 dollars
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # CPI Strip
    st.markdown('<div class="section-label">The Equivalent Dollar Amount Invested At Each Peak</div>', unsafe_allow_html=True)
    cpi_cols = st.columns(4)
    for i, (s, r) in enumerate(zip(SCENARIOS, results)):
        with cpi_cols[i % 4]:
            st.markdown(f"""
            <div class="cpi-tile">
                <div class="cpi-date">{r['start_str']}</div>
                <div class="cpi-amt">{fmt(r['scaled'])}</div>
                <div class="cpi-eq">≡ {fmt(investment)} today</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Fear Bar
    st.markdown("""
    <div class="fear-bar">
        <div style="display:flex;align-items:flex-start;gap:20px;flex-wrap:wrap">
            <div style="font-size:28px;flex-shrink:0">⚠️</div>
            <div style="flex:1;min-width:0">
                <div class="fear-header">THE NARRATIVE AT EVERY SINGLE PEAK</div>
                <div class="fear-text">"Valuations are stretched." · "The market can't go higher." · "A crash is inevitable." · People have said this at every single peak — 1929, 1968, 2000, 2007, the pandemic crash of 2020, and today.</div>
            </div>
            <div class="fear-badge">EVERY SINGLE TIME, THEY WERE RIGHT ABOUT THE CRASH — AND WRONG ABOUT WHAT FOLLOWED.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Scenario Cards
    st.markdown('<div class="section-label">Nine Moments of Maximum Pessimism · Nine Investors · One Result</div>', unsafe_allow_html=True)

    for s, r in zip(SCENARIOS, results):
        scaled = r["scaled"]
        real_end = r["real_end"]
        nom_end = r["nom_end"]
        price_only = r["price_only"]
        max_val = max(real_end, nom_end, price_only, scaled)

        # Card header (always visible)
        st.markdown(render_scenario_card(s, r, investment), unsafe_allow_html=True)

        # Expandable details
        with st.expander(f"Details: {s['label']} · {r['start_str']}"):
            st.markdown(
                f'<div class="conversion-note">Invested <strong>{fmt(scaled)}</strong> '
                f'(CPI {r["cpi_start"]} → {r["cpi_end"]}) — equivalent purchasing power of '
                f'<strong>{fmt(investment)}</strong> today.</div>',
                unsafe_allow_html=True,
            )

            st.markdown(f'<div class="narrative">{s["narrative"]}</div>', unsafe_allow_html=True)

            if s.get("quote"):
                st.markdown(
                    f'<div class="era-quote">'
                    f'<blockquote>"{s["quote"]}"</blockquote>'
                    f'<div class="quote-attr">— {s["quote_attr"]}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

            if s["pain"]:
                st.markdown(f'<div class="pain-box"><strong>The Pain</strong><br>{s["pain"]}</div>', unsafe_allow_html=True)

            # Metrics grid
            m_cols = st.columns(3)
            metrics = [
                ("Invested", fmt(scaled), "teal"),
                ("★ Real Ending Value", fmt(real_end), "gold"),
                ("Nominal Ending Value", fmt(nom_end), ""),
                ("Price Only (Real)", fmt(price_only), ""),
                ("Dividend Growth (Real)", fmt_x(r["div_real_f"]), ""),
                ("Holding Period", f'{r["years"]} years', ""),
            ]
            for j, (label, val, color) in enumerate(metrics):
                with m_cols[j % 3]:
                    css_class = f"metric-val {color}" if color else "metric-val"
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">{label}</div>
                        <div class="{css_class}">{val}</div>
                    </div>
                    """, unsafe_allow_html=True)

            # Bar chart
            bars_html = '<div class="bar-container">'
            bars_html += render_bar("Nominal TR", nom_end, max_val, "bar-fill-gold")
            bars_html += render_bar("Real TR ★", real_end, max_val, "bar-fill-green")
            bars_html += render_bar("Price Only", price_only, max_val, "bar-fill-teal")
            bars_html += render_bar("Invested", scaled, max_val, "bar-fill-muted")
            bars_html += '</div>'
            st.markdown(bars_html, unsafe_allow_html=True)

    # Closing
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        f'<div class="closing-quote">Every investor with the equivalent of {fmt(investment)} in today\'s money '
        f'who bought at the peaks, at the worst moments, even when the experts declared stocks dead — and simply held — made money. '
        f'In every single case.</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="closing-tagline">"The risk was never buying at the top. The risk was not buying at all."<br><span style="font-size:14px;color:var(--muted);font-style:normal">— Paul Ruedi</span></div>',
        unsafe_allow_html=True,
    )

    # Summary table
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-label">Summary Comparison</div>', unsafe_allow_html=True)

    table_html = '<table class="summary-table"><thead><tr>'
    table_html += "<th>Peak / Crisis</th><th>Crash After</th><th>Invested</th>"
    table_html += "<th>★ Real Value</th><th>Real TR Factor</th><th>Real Div Factor</th><th>Years</th>"
    table_html += "</tr></thead><tbody>"
    for s, r in zip(SCENARIOS, results):
        table_html += (
            f'<tr><td>{r["start_str"]} · {s["short"]}</td>'
            f'<td>{s["crash"]}</td>'
            f'<td class="col-teal">{fmt(r["scaled"])}</td>'
            f'<td class="col-gold">{fmt(r["real_end"])}</td>'
            f'<td>{fmt_x(r["real_tr_f"])}</td>'
            f'<td>{fmt_x(r["div_real_f"])}</td>'
            f'<td>{r["years"]}</td></tr>'
        )
    table_html += "</tbody></table>"
    st.markdown(table_html, unsafe_allow_html=True)

    # Footnote
    st.markdown(
        f'<div class="footnote-text">Created by Paul Ruedi<br>'
        f'Source: Robert Shiller, <em>Irrational Exuberance</em> dataset · '
        f'irrationalexuberance.com · S&P 500 monthly data with dividends reinvested · '
        f'CPI-adjusted to constant purchasing power · Data through {data_end_str}</div>',
        unsafe_allow_html=True,
    )


# ═══════════════════════════════════════════
#  TAB 2: SINGLE DATE QUERY
# ═══════════════════════════════════════════
with tab2:
    st.markdown('<div class="hero-title" style="font-size:36px">Single Date Analysis</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="hero-subtitle">Pick any historical date or event and see what happened to an S&P 500 investment held to today.</div>',
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)

    with col1:
        input_mode = st.radio("Choose input method", ["Historical event", "Custom date"], horizontal=True)

        if input_mode == "Historical event":
            event = st.selectbox("Select an event", list(DATE_ALIASES.keys()))
            target_date = DATE_ALIASES[event]
        else:
            year = st.number_input("Year", min_value=1871, max_value=2026, value=2000)
            month = st.selectbox("Month", MONTHS, index=0)
            month_num = MONTHS.index(month) + 1
            target_date = year + month_num / 100

    with col2:
        single_investment = st.number_input(
            "Investment amount ($)",
            min_value=100,
            max_value=100_000_000,
            value=10_000,
            step=1_000,
            format="%d",
            key="investment_single",
        )

    if st.button("Analyze", type="primary", use_container_width=True):
        r = analyze(target_date, single_investment, rows)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            f'<div class="eyebrow">{r["start_str"]} → {r["end_str"]} · {fmt(single_investment)} invested (today\'s dollars)</div>',
            unsafe_allow_html=True,
        )

        # Metric cards
        st.markdown('<br>', unsafe_allow_html=True)
        mc = st.columns(3)
        card_data = [
            ("CPI-Adjusted Amount Invested", fmt(r["scaled"]), "teal"),
            ("★ Real Ending Value", fmt(r["real_end"]), "gold"),
            ("Nominal Ending Value", fmt(r["nom_end"]), ""),
            ("Price Only (Real)", fmt(r["price_only"]), ""),
            ("Inflation (CPI)", f'{r["cpi_factor"]}x', ""),
            ("Holding Period", f'{r["years"]} years', ""),
        ]
        for j, (label, val, color) in enumerate(card_data):
            with mc[j % 3]:
                css = f"metric-val {color}" if color else "metric-val"
                st.markdown(f"""
                <div class="metric-card" style="margin-bottom:10px">
                    <div class="metric-label">{label}</div>
                    <div class="{css}">{val}</div>
                </div>
                """, unsafe_allow_html=True)

        # Return breakdown table
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"""
        <table class="analysis-table">
            <thead>
                <tr><th></th><th>Nominal</th><th>Real (Infl-Adj)</th></tr>
            </thead>
            <tbody>
                <tr>
                    <td>🏷️ Inflation (CPI)</td>
                    <td>{r['cpi_factor']}x</td>
                    <td>—</td>
                </tr>
                <tr>
                    <td>💸 Dividend growth</td>
                    <td>{fmt_x(r['div_nom_f'])}</td>
                    <td>{fmt_x(r['div_real_f'])}</td>
                </tr>
                <tr>
                    <td>📈 Price only</td>
                    <td>—</td>
                    <td>{fmt_x(r['price_r_f'])}</td>
                </tr>
                <tr>
                    <td>💰 Total return</td>
                    <td>{fmt_x(r['nom_tr_f'])}</td>
                    <td>{fmt_x(r['real_tr_f'])}</td>
                </tr>
                <tr>
                    <td></td>
                    <td>{fmt(r['scaled'])} → {fmt(r['nom_end'])}</td>
                    <td>{fmt(r['scaled'])} → {fmt(r['real_end'])}</td>
                </tr>
            </tbody>
        </table>
        <div style="font-family:'DM Mono',monospace;font-size:11px;color:var(--muted);margin-top:8px">
            Real = CPI stripped out. Total return = dividends reinvested. Holding period: {r['years']} years.
        </div>
        """, unsafe_allow_html=True)

        # Bar chart
        max_val = max(r["real_end"], r["nom_end"], r["price_only"], r["scaled"])
        bars_html = '<div class="bar-container" style="margin-top:24px">'
        bars_html += render_bar("Nominal TR", r["nom_end"], max_val, "bar-fill-gold")
        bars_html += render_bar("Real TR ★", r["real_end"], max_val, "bar-fill-green")
        bars_html += render_bar("Price Only", r["price_only"], max_val, "bar-fill-teal")
        bars_html += render_bar("Invested", r["scaled"], max_val, "bar-fill-muted")
        bars_html += '</div>'
        st.markdown(bars_html, unsafe_allow_html=True)

        # Footnote
        st.markdown(
            f'<div class="footnote-text" style="margin-top:32px">Source: Robert Shiller, <em>Irrational Exuberance</em> dataset · '
            f'irrationalexuberance.com · Data through {data_end_str}</div>',
            unsafe_allow_html=True,
        )
