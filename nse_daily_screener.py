"""
NSE Nifty 250 Daily DMA Screener - RELIABLE VERSION
=====================================================
Uses Yahoo Finance (yfinance) as primary data source.
- No NSE cookie/session issues
- Works 100% reliably in GitHub Actions
- Gets real closing prices + calculates DMA 50/100/200

Install: pip install yfinance pandas openpyxl requests
"""

import os, json, time, warnings
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ── Config ────────────────────────────────────────────────────────────────────
DOCS_DIR               = "./docs"
HIGH_VOL_MULTIPLIER    = 1.5
HIGH_VALUE_CRORE       = 5
DMA200_OVEREXTENDED_PCT= 10
HISTORY_DAYS           = 220   # need 200 for DMA200 + buffer

# ── Nifty 250 symbols (Yahoo Finance format = NSE symbol + ".NS") ─────────────
NIFTY250_NSE = [
"COALINDIA","HDFCBANK","ATGL","ICICIBANK","ADANIPOWER","BSE","RELIANCE","BHARTIARTL","SUZLON","ADANIENSOL","ADANIENT","POWERINDIA","CUMMINSIND",
"ADANIGREEN","MCX","VEDL","AXISBANK","ETERNAL","EXIDEIND","HINDALCO","TVSMOTOR","CGPOWER","TATASTEEL","IDEA","ONGC","WIPRO","SBIN","ITC","INFY",
"SAIL","SWIGGY","ENRIN","MOTHERSON","GVT&D","NATIONALUM","SIEMENS","JSWENERGY","LT","TMPV","INDIGO","WAAREEENER","TCS","NTPC","AIAENG","BAJFINANCE",
"PREMIERENE","LICI","OFSS","BEL","ABB","BHEL","M&M","DIXON","MARUTI","ADANIPORTS","RVNL","KOTAKBANK","ASHOKLEY","CANBK","HINDPETRO","SHRIRAMFIN","BAJAJ-AUTO",
"APARINDS","GROWW","POWERGRID","HINDZINC","ULTRACEMCO","THERMAX","BANKBARODA","NAUKRI","TECHM","JSWSTEEL","LENSKART","TATAPOWER","MAXHEALTH","TMCV","SUNPHARMA",
"PAYTM","APOLLOHOSP","COFORGE","HEROMOTOCO","LODHA","ZYDUSLIFE","ASIANPAINT","TATACOMM","VBL","ICICIAMC","HAL","NMDC","JIOFIN","HYUNDAI","INDUSTOWER",
"POLYCAB","EICHERMOT","NLCINDIA","TIINDIA","AMBUJACEM","TITAN","SOLARINDS","TATACONSUM","YESBANK","TRENT","BPCL","LAURUSLABS","BDL","BANKINDIA","KEI",
"COCHINSHIP","CONCOR","RADICO","NHPC","NESTLEIND","IOC","APLAPOLLO","MOTILALOFS","HAVELLS","HDFCAMC","JINDALSTEL","NYKAA","PFC","HINDUNILVR","CIPLA","MAZDOCK",
"LUPIN","GLENMARK","UNIONBANK","DRREDDY","GODREJCP","PNB","FLUOROCHEM","MFSL","IDFCFIRSTB","HCLTECH","AUROPHARMA","DLF","FEDERALBNK","BIOCON","LLOYDSME",
"RECLTD","BHARATFORG","VOLTAS","DIVISLAB","PERSISTENT","GMRAIRPORT","BRITANNIA","MANKIND","APOLLOTYRE","INDUSINDBK","ALKEM","GRASIM","GODREJPROP","KPITTECH",
"IRCTC","TORNTPHARM","DMART","GAIL","UNOMINDA","INDHOTEL","CHOLAFIN","PRESTIGE","INDIANB","AUBANK","ABCAPITAL","MARICO","HDFCLIFE","LTM","TORNTPOWER","ASTRAL",
"TATAELXSI","POLICYBZR","KALYANKJIL","PHOENIXLTD","VMM","SBILIFE","COROMANDEL","GODFRYPHLP","IRFC","LICHSGFIN","IREDA","BAJAJFINSV","SRF","NAM-INDIA","UPL",
"MAHABANK","ITCHOTELS","UNITDSPR","DABUR","OIL","MUTHOOTFIN","JUBLFOOD","PAGEIND","OBEROIRLTY","FORTIS","LGEINDIA","MPHASIS","PATANJALI","LTF","BOSCHLTD",
"BAJAJHFL","AWL","AIIL","PIDILITIND","AJANTPHARM","NTPCGREEN","ICICIPRULI","ICICIGI","JSWINFRA","HONAUT","COLPAL","GLAXO","ESCORTS","BLUESTARCO","MRF","M&MFIN",
"TATACAP","SJVN","JSL","PETRONET","BAJAJHLDNG","360ONE","SBICARD","ACC","SUPREMEIND","PIIND","HUDCO","SUNDARMFIN","DALBHARAT","LINDEINDIA","TATAINVEST",
"JKCEMENT","GICRE","UBL","SHREECEM","MEDANTA","LTTS","ENDURANCE","ANTHEM","ABBOTINDIA","CRISIL","BERGEPAINT","KPRMILL","BALKRISIND","IPCALAB","SCHAEFFLER",
"NIACL","3MINDIA","HDBFS","BHARTIHEXA","HEXT","GODREJIND",
]

# Deduplicate
seen = set()
NIFTY250_NSE = [x for x in NIFTY250_NSE if not (x in seen or seen.add(x))]

# Convert to Yahoo Finance tickers (append .NS)
def to_yf_ticker(sym):
    # Special cases
    special = {"M&M": "M&M.NS", "BAJAJ-AUTO": "BAJAJ-AUTO.NS"}
    return special.get(sym, f"{sym}.NS")

# ── Step 1: Download historical data via yfinance ─────────────────────────────
def fetch_all_data():
    try:
        import yfinance as yf
    except ImportError:
        raise ImportError("Run: pip install yfinance")

    print("\n[1/3] Downloading data from Yahoo Finance…")

    end_date   = datetime.today()
    start_date = end_date - timedelta(days=int(HISTORY_DAYS * 1.6))  # extra buffer for weekends/holidays

    tickers = [to_yf_ticker(s) for s in NIFTY250_NSE]

    print(f"  Fetching {len(tickers)} stocks × {HISTORY_DAYS} days…")

    # Download in batches of 50 to avoid timeouts
    all_close  = {}
    all_volume = {}
    all_today  = {}

    BATCH = 50
    for i in range(0, len(tickers), BATCH):
        batch = tickers[i:i+BATCH]
        batch_syms = NIFTY250_NSE[i:i+BATCH]
        print(f"  Batch {i//BATCH + 1}/{(len(tickers)-1)//BATCH + 1}: {batch_syms[0]}…{batch_syms[-1]}")

        try:
            raw = yf.download(
                batch,
                start=start_date.strftime("%Y-%m-%d"),
                end=end_date.strftime("%Y-%m-%d"),
                auto_adjust=True,
                progress=False,
                threads=True,
                group_by="ticker",
            )

            for sym, ticker in zip(batch_syms, batch):
                try:
                    if len(batch) == 1:
                        close_series  = raw["Close"]
                        volume_series = raw["Volume"]
                        high_series   = raw["High"]
                        low_series    = raw["Low"]
                        open_series   = raw["Open"]
                    else:
                        if ticker not in raw.columns.get_level_values(0):
                            continue
                        close_series  = raw[ticker]["Close"]
                        volume_series = raw[ticker]["Volume"]
                        high_series   = raw[ticker]["High"]
                        low_series    = raw[ticker]["Low"]
                        open_series   = raw[ticker]["Open"]

                    close_series  = close_series.dropna()
                    volume_series = volume_series.dropna()

                    if len(close_series) < 10:
                        print(f"    ⚠ {sym}: not enough data ({len(close_series)} rows)")
                        continue

                    all_close[sym]  = close_series
                    all_volume[sym] = volume_series

                    # Today's (last available) data
                    last_date = close_series.index[-1]
                    all_today[sym] = {
                        "ltp":       round(float(close_series.iloc[-1]), 2),
                        "prevClose": round(float(close_series.iloc[-2]), 2) if len(close_series) >= 2 else 0,
                        "high":      round(float(high_series.loc[last_date]),   2) if last_date in high_series.index   else 0,
                        "low":       round(float(low_series.loc[last_date]),    2) if last_date in low_series.index    else 0,
                        "open":      round(float(open_series.loc[last_date]),   2) if last_date in open_series.index   else 0,
                        "volume":    int(float(volume_series.loc[last_date]))       if last_date in volume_series.index else 0,
                        "date":      last_date.strftime("%Y-%m-%d"),
                    }
                except Exception as e:
                    print(f"    ⚠ {sym}: {e}")
                    continue

        except Exception as e:
            print(f"  ✗ Batch error: {e}")
            continue

        time.sleep(1)  # be polite

    print(f"  ✓ Got data for {len(all_today)} / {len(NIFTY250_NSE)} stocks")
    return all_close, all_volume, all_today

# ── Step 2: Calculate DMAs and screen ─────────────────────────────────────────
def screen(all_close, all_volume, all_today):
    print("\n[2/3] Calculating DMAs and screening…")
    records = []

    for sym in NIFTY250_NSE:
        if sym not in all_today:
            continue

        td  = all_today[sym]
        ltp = td["ltp"]
        prv = td["prevClose"]
        vol = td["volume"]
        hgh = td["high"]
        low = td["low"]

        chg     = round(ltp - prv, 2)
        chg_pct = round((chg / prv * 100) if prv else 0, 2)

        # Close price history for DMA
        hist = all_close.get(sym, pd.Series(dtype=float)).dropna()
        # Remove today from history for DMA calc (use only historical closes)
        hist_for_dma = hist.iloc[:-1] if len(hist) > 1 else hist

        dma50   = round(float(hist_for_dma.tail(50).mean()),  2) if len(hist_for_dma) >= 50  else None
        dma100  = round(float(hist_for_dma.tail(100).mean()), 2) if len(hist_for_dma) >= 100 else None
        dma200  = round(float(hist_for_dma.tail(200).mean()), 2) if len(hist_for_dma) >= 200 else None

        # Average volume (20-day)
        vol_hist  = all_volume.get(sym, pd.Series(dtype=float)).dropna()
        avg_vol   = float(vol_hist.tail(20).mean()) if len(vol_hist) >= 5 else 0

        # Value in crores (approx: ltp × volume / 1Cr)
        val_cr    = round(ltp * vol / 1e7, 2)

        a50   = bool(ltp > dma50)  if dma50  else False
        a100  = bool(ltp > dma100) if dma100 else False
        a200  = bool(ltp > dma200) if dma200 else False
        pct200= round((ltp - dma200) / dma200 * 100, 2) if dma200 else None
        hi_vol= bool(vol > avg_vol * HIGH_VOL_MULTIPLIER) if avg_vol else False
        hi_val= bool(val_cr >= HIGH_VALUE_CRORE)
        over  = bool(pct200 > DMA200_OVEREXTENDED_PCT) if pct200 is not None else False

        if   a50 and a100 and a200 and hi_vol and not over: sig = "BUY"
        elif a50 and a100 and a200 and over:                sig = "OVEREXTENDED"
        elif a50 and a100 and a200:                         sig = "HOLD"
        else:                                               sig = "WATCH"

        records.append({
            "symbol":        sym,
            "ltp":           ltp,
            "prevClose":     prv,
            "change":        chg,
            "changePct":     chg_pct,
            "high":          hgh,
            "low":           low,
            "volume":        vol,
            "avgVolume":     int(avg_vol),
            "highVolume":    hi_vol,
            "valueCr":       val_cr,
            "highValue":     hi_val,
            "dma50":         dma50,
            "aboveDMA50":    a50,
            "dma100":        dma100,
            "aboveDMA100":   a100,
            "dma200":        dma200,
            "aboveDMA200":   a200,
            "pctAboveDMA200":pct200,
            "dma200Alert":   "OVEREXTENDED" if over else "IN_RANGE",
            "signal":        sig,
            "dataDate":      td["date"],
        })

    df = pd.DataFrame(records)
    if df.empty:
        print("  ✗ No records! Check network/yfinance.")
        return df

    df = df.sort_values("pctAboveDMA200", ascending=False,
                        key=lambda x: pd.to_numeric(x, errors="coerce"))

    buy  = (df.signal=='BUY').sum()
    hold = (df.signal=='HOLD').sum()
    ovr  = (df.signal=='OVEREXTENDED').sum()
    wtch = (df.signal=='WATCH').sum()
    print(f"  ✓ {len(df)} stocks screened")
    print(f"     BUY:{buy}  HOLD:{hold}  OVEREXTENDED:{ovr}  WATCH:{wtch}")

    # Print sample to verify correctness
    sample = df[df.symbol.isin(["RELIANCE","TCS","HDFCBANK"])][["symbol","ltp","dma50","dma200","signal"]]
    print(f"\n  Sample check:\n{sample.to_string(index=False)}\n")

    return df

# ── Step 3a: Export JSON for dashboard ────────────────────────────────────────
def export_json(df):
    os.makedirs(DOCS_DIR, exist_ok=True)
    payload = {
        "generated_at": datetime.now().strftime("%d %b %Y %H:%M IST"),
        "trading_day":  df["dataDate"].iloc[0] if not df.empty else "N/A",
        "stocks": df.where(pd.notnull(df), None).to_dict(orient="records"),
    }
    path = os.path.join(DOCS_DIR, "screener_data.json")
    with open(path, "w") as f:
        json.dump(payload, f, separators=(',', ':'))
    size = os.path.getsize(path) / 1024
    print(f"  ✓ JSON → {path}  ({size:.1f} KB)")

# ── Step 3b: Export colour-coded Excel ────────────────────────────────────────
CLR = dict(header="0d1e3b", buy="C6EFCE", hold="FFEB9C",
           over="FFC7CE",   watch="F2F2F2", yes="E2EFDA", no="FCE4D6")

def export_excel(df):
    rows = []
    for _, s in df.iterrows():
        rows.append({
            "Symbol":         s.symbol,
            "Date":           s.dataDate,
            "LTP (₹)":        s.ltp,
            "Prev Close (₹)": s.prevClose,
            "Change (₹)":     s.change,
            "Change (%)":     s.changePct,
            "High (₹)":       s.high,
            "Low (₹)":        s.low,
            "Volume":         s.volume,
            "Avg Vol (20D)":  s.avgVolume,
            "High Volume":    "YES" if s.highVolume else "NO",
            "Value (₹ Cr)":   s.valueCr,
            "High Value":     "YES" if s.highValue  else "NO",
            "DMA 50 (₹)":     s.dma50,
            "Above DMA 50":   "YES" if s.aboveDMA50  else "NO",
            "DMA 100 (₹)":    s.dma100,
            "Above DMA 100":  "YES" if s.aboveDMA100 else "NO",
            "DMA 200 (₹)":    s.dma200,
            "Above DMA 200":  "YES" if s.aboveDMA200 else "NO",
            "% vs DMA 200":   s.pctAboveDMA200,
            "DMA 200 Alert":  "⚠ OVEREXTENDED (>10%)" if s.dma200Alert=="OVEREXTENDED" else "✓ Within Range",
            "Signal":         s.signal,
        })

    out_df = pd.DataFrame(rows)
    os.makedirs(DOCS_DIR, exist_ok=True)
    arc = os.path.join(DOCS_DIR, "archive")
    os.makedirs(arc, exist_ok=True)

    date_str = datetime.today().strftime("%Y-%m-%d")
    dated    = os.path.join(arc,      f"NSE_DMA_Screener_{date_str}.xlsx")
    latest   = os.path.join(DOCS_DIR, "latest.xlsx")

    def write(path):
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            out_df.to_excel(w, sheet_name="All Nifty 250", index=False)
            out_df[out_df.Signal.isin(["BUY","HOLD","OVEREXTENDED"])].to_excel(
                w, sheet_name="Screened", index=False)
            out_df[out_df.Signal == "BUY"].to_excel(
                w, sheet_name="BUY Signals", index=False)
        wb = load_workbook(path)
        for ws in wb.worksheets:
            _fmt_sheet(ws)
        wb.save(path)

    write(dated)
    import shutil
    shutil.copy(dated, latest)
    print(f"  ✓ Excel → {latest}")
    print(f"  ✓ Archive → {dated}")

def _fmt_sheet(ws):
    thin = Side(style="thin", color="D0D0D0")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=CLR["header"])
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    for col in ws.columns:
        mw = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(mw + 4, 24)
    hdrs = {c.value: c.column for c in ws[1]}
    SC   = {"BUY": CLR["buy"], "HOLD": CLR["hold"],
            "OVEREXTENDED": CLR["over"], "WATCH": CLR["watch"]}
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border    = bdr
            cell.alignment = Alignment(horizontal="right")
        for col_name in ["Above DMA 50", "Above DMA 100", "Above DMA 200"]:
            ci = hdrs.get(col_name)
            if ci:
                c = row[ci-1]
                c.fill = PatternFill("solid",
                    fgColor=(CLR["yes"] if c.value == "YES" else CLR["no"]))
                c.font = Font(bold=True)
        ci = hdrs.get("Signal")
        if ci:
            c = row[ci-1]
            clr = SC.get(str(c.value), "")
            if clr:
                c.fill = PatternFill("solid", fgColor=clr)
                c.font = Font(bold=True)
        ci = hdrs.get("DMA 200 Alert")
        if ci:
            c = row[ci-1]
            c.fill = PatternFill("solid",
                fgColor=(CLR["over"] if "OVER" in str(c.value) else CLR["buy"]))
            c.font = Font(bold=True)
        ci = hdrs.get("Change (%)")
        if ci:
            c = row[ci-1]
            try:
                v = float(c.value)
                c.font = Font(color=("375623" if v >= 0 else "9C0006"), bold=True)
            except:
                pass
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

# ── Main ──────────────────────────────────────────────────────────────────────
def run():
    print("=" * 60)
    print("  NSE Nifty 250 Daily DMA Screener (yfinance)")
    print(f"  {datetime.now().strftime('%d %b %Y %H:%M UTC')}")
    print("=" * 60)

    all_close, all_volume, all_today = fetch_all_data()

    if not all_today:
        print("\n✗ FATAL: No data fetched. Exiting.")
        raise SystemExit(1)

    result_df = screen(all_close, all_volume, all_today)

    if result_df.empty:
        print("\n✗ FATAL: Screening produced no results. Exiting.")
        raise SystemExit(1)

    print("\n[3/3] Exporting…")
    export_json(result_df)
    export_excel(result_df)

    print("\n" + "=" * 60)
    print("  ✅ ALL DONE!")
    print("=" * 60)

if __name__ == "__main__":
    run()
