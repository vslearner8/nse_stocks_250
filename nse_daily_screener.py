"""
NSE Nifty 250 Daily DMA Screener
Downloads NSE CM-UDiFF Bhavcopy, calculates DMA 50/100/200,
screens stocks, outputs Excel + JSON for the live dashboard.
"""

import os, io, zipfile, json, time, requests, pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Config ────────────────────────────────────────────────────────────────────
DOCS_DIR    = "./docs"
HISTORY_DIR = "./nse_history"
HIGH_VOL_MULTIPLIER     = 1.5
HIGH_VALUE_CRORE        = 5
DMA200_OVEREXTENDED_PCT = 10

NIFTY250 = [
    "RELIANCE","TCS","HDFCBANK","INFY","ICICIBANK","HINDUNILVR","SBIN","BHARTIARTL",
    "ITC","KOTAKBANK","LT","AXISBANK","ASIANPAINT","MARUTI","BAJFINANCE","TITAN",
    "NESTLEIND","WIPRO","ULTRACEMCO","TECHM","SUNPHARMA","POWERGRID","NTPC","ONGC",
    "TATAMOTORS","HCLTECH","ADANIENT","ADANIPORTS","BAJAJFINSV","DIVISLAB",
    "DRREDDY","EICHERMOT","GRASIM","HEROMOTOCO","HINDALCO","INDUSINDBK","JSWSTEEL",
    "M&M","CIPLA","COALINDIA","APOLLOHOSP","DABUR","GODREJCP","MARICO","PIDILITIND",
    "SIEMENS","TATACONSUM","TORNTPHARM","VEDL","ZOMATO","IRCTC","HDFCLIFE",
    "SBILIFE","ICICIPRULI","BAJAJ-AUTO","BOSCHLTD","BRITANNIA","COLPAL","CONCOR",
    "CUMMINSIND","HAVELLS","HINDPETRO","IOC","MCDOWELL-N","MOTHERSON","MPHASIS",
    "OFSS","PETRONET","PIIND","SAIL","TATAPOWER","TATASTEEL","TVSMOTOR","UBL",
    "VOLTAS","ZYDUSLIFE","ACC","ADANIGREEN","ADANIPOWER","AMBUJACEM","AUBANK",
    "BALKRISIND","BANDHANBNK","BANKBARODA","BEL","BPCL","CANFINHOME","CANBK",
    "CHOLAFIN","DEEPAKNTR","DIXON","DLF","ESCORTS","FEDERALBNK","GAIL",
    "GODREJPROP","GRINDWELL","HAL","HDFCAMC","ICICIGI","IGL","INDHOTEL","INDIGO",
    "INDUSTOWER","IPCALAB","JUBLFOOD","KAJARIACER","LALPATHLAB","LICHSGFIN",
    "LUPIN","MANAPPURAM","MUTHOOTFIN","NATCOPHARM","NAUKRI","NBCC","NCC",
    "NHPC","NMDC","OBEROIRLTY","PAGEIND","PERSISTENT","PHOENIXLTD","POLYCAB",
    "RAMCOCEM","RECLTD","RELAXO","SBICARD","SCHAEFFLER","SRF","STARHEALTH",
    "SUNTV","SUPREMEIND","SYNGENE","TATACOMM","TORNTPOWER","TRENT","UNIONBANK",
    "UPL","VGUARD","YESBANK","ZEEL","SUNDARMFIN","LATENTVIEW","ANGELONE",
    "KAYNES","MAPMYINDIA","METRO","RRKABEL","UTIAMC","RITES","DEEPAKNTR",
    "NAVINFLUOR","GNFC","GMRINFRA","LINDEINDIA","MFSL","PIIND","RAYMOND",
    "TATACONSUM","TORNTPHARM","VEDL","WHIRLPOOL","ABCAPITAL","CHOLAFIN",
]

# Remove duplicates while preserving order
seen = set()
NIFTY250 = [x for x in NIFTY250 if not (x in seen or seen.add(x))]

# ── NSE URL templates ─────────────────────────────────────────────────────────
# New CM-UDiFF format (post July 2024)
BHAVCOPY_NEW = (
    "https://nsearchives.nseindia.com/content/cm/"
    "BhavCopy_NSE_CM_0_0_0_{date}_F_0000.csv.zip"
)
# Legacy fallback
BHAVCOPY_OLD = (
    "https://archives.nseindia.com/content/historical/EQUITIES/"
    "{year}/{mon}/cm{date2}bhav.csv.zip"
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Encoding": "gzip, deflate, br",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
    "Referer": "https://www.nseindia.com/",
    "sec-ch-ua": '"Chromium";v="124"',
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "document",
    "sec-fetch-mode": "navigate",
    "sec-fetch-site": "same-origin",
}

# ── Helpers ───────────────────────────────────────────────────────────────────
def last_trading_day(ref=None):
    d = ref or datetime.today()
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d

def get_nse_session():
    """Create a session with valid NSE cookies."""
    session = requests.Session()
    session.headers.update(HEADERS)
    try:
        # Hit home page first to get cookies
        session.get("https://www.nseindia.com", timeout=15)
        time.sleep(2)
        # Hit market data page (sets additional cookies)
        session.get("https://www.nseindia.com/market-data/all-reports", timeout=15)
        time.sleep(1)
    except Exception as e:
        print(f"  ⚠ Session setup warning: {e}")
    return session

def download_bhavcopy(date_obj, session=None):
    date_str  = date_obj.strftime("%Y%m%d")
    cache     = os.path.join(HISTORY_DIR, f"bhav_{date_str}.csv")
    if os.path.exists(cache):
        return pd.read_csv(cache)

    os.makedirs(HISTORY_DIR, exist_ok=True)

    if session is None:
        session = get_nse_session()

    # Try new URL format first
    url = BHAVCOPY_NEW.format(date=date_str)
    print(f"  ↓ {url}")
    try:
        resp = session.get(url, timeout=30)
        if resp.status_code == 200 and len(resp.content) > 500:
            with zipfile.ZipFile(io.BytesIO(resp.content)) as z:
                csv_name = [n for n in z.namelist() if n.endswith(".csv")][0]
                df = pd.read_csv(z.open(csv_name))
            df.to_csv(cache, index=False)
            return df
        print(f"  ↙ New URL returned {resp.status_code}, trying legacy…")
    except Exception as e:
        print(f"  ↙ New URL error: {e}, trying legacy…")

    # Legacy fallback URL
    try:
        mon   = date_obj.strftime("%b").upper()
        year  = date_obj.strftime("%Y")
        date2 = date_obj.strftime("%d%b%Y").upper()
        url2  = BHAVCOPY_OLD.format(year=year, mon=mon, date2=date2)
        print(f"  ↓ {url2}")
        resp2 = session.get(url2, timeout=30)
        if resp2.status_code == 200:
            with zipfile.ZipFile(io.BytesIO(resp2.content)) as z:
                csv_name = [n for n in z.namelist() if n.endswith(".csv")][0]
                df = pd.read_csv(z.open(csv_name))
            df.to_csv(cache, index=False)
            return df
    except Exception as e2:
        print(f"  ✗ Legacy URL error: {e2}")

    raise Exception(f"Failed to download bhavcopy for {date_str}")

def normalise(df):
    """Normalise column names across old and new bhavcopy formats."""
    df.columns = [c.strip().upper().replace(" ", "") for c in df.columns]
    rename = {
        # New format → standard
        "TCKRSYMB":    "SYMBOL",
        "CLSPRIC":     "CLOSE",
        "OPNPRIC":     "OPEN",
        "HGHPRIC":     "HIGH",
        "LWPRIC":      "LOW",
        "PRVSCLSGPRIC":"PREVCLOSE",
        "TTLTRADGVOL": "VOLUME",
        "TTLTRFVAL":   "VALUE",
        "SCTYSRS":     "SERIES",
        # Old format already uses SYMBOL, CLOSE etc — no rename needed
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
    return df

def build_history(today, days=210):
    print(f"\n[2/4] Building {days}-day price history…")
    session = get_nse_session()
    closes  = {}
    date    = today
    count   = 0
    fails   = 0
    while count < days and fails < 20:
        try:
            df  = download_bhavcopy(date, session)
            df  = normalise(df)
            eq  = df[df["SERIES"] == "EQ"].copy() if "SERIES" in df.columns else df.copy()
            if "SYMBOL" in eq.columns and "CLOSE" in eq.columns:
                eq["CLOSE"] = pd.to_numeric(eq["CLOSE"], errors="coerce")
                closes[date.strftime("%Y%m%d")] = eq.set_index("SYMBOL")["CLOSE"]
                count += 1
                fails = 0
            else:
                print(f"    ⚠ Missing cols in {date.strftime('%Y%m%d')}: {list(df.columns[:8])}")
                fails += 1
        except Exception as e:
            print(f"    skip {date.strftime('%Y%m%d')}: {e}")
            fails += 1
        date -= timedelta(days=1)
        while date.weekday() >= 5:
            date -= timedelta(days=1)
        time.sleep(0.3)  # be polite to NSE servers

    price_df = pd.DataFrame(closes).T.sort_index()
    price_df = price_df.apply(pd.to_numeric, errors="coerce")
    print(f"  ✓ {len(price_df)} days × {len(price_df.columns)} symbols")
    return price_df, session

def screen(today, price_df, session):
    print("\n[3/4] Screening…")
    bhav = download_bhavcopy(today, session)
    bhav = normalise(bhav)

    eq = bhav[bhav.get("SERIES", bhav.iloc[:, 0]) == "EQ"].copy() \
         if "SERIES" in bhav.columns else bhav.copy()
    eq = eq.set_index("SYMBOL")

    for col in ["CLOSE","PREVCLOSE","HIGH","LOW","VOLUME","VALUE"]:
        if col in eq.columns:
            eq[col] = pd.to_numeric(eq[col], errors="coerce")

    records = []
    for sym in NIFTY250:
        if sym not in eq.index:
            continue
        r   = eq.loc[sym]
        ltp = float(r.get("CLOSE",     0) or 0)
        prv = float(r.get("PREVCLOSE", 0) or 0)
        vol = float(r.get("VOLUME",    0) or 0)
        val = float(r.get("VALUE",     0) or 0)
        hgh = float(r.get("HIGH",      0) or 0)
        low = float(r.get("LOW",       0) or 0)

        chg     = round(ltp - prv, 2)
        chg_pct = round((chg / prv * 100) if prv else 0, 2)

        hist    = price_df[sym].dropna() if sym in price_df.columns else pd.Series(dtype=float)
        dma50   = round(hist.tail(50).mean(),  2) if len(hist) >= 50  else None
        dma100  = round(hist.tail(100).mean(), 2) if len(hist) >= 100 else None
        dma200  = round(hist.tail(200).mean(), 2) if len(hist) >= 200 else None
        avg_vol = hist.tail(20).mean() if len(hist) >= 20 else 0

        a50    = bool(ltp > dma50)  if dma50  else False
        a100   = bool(ltp > dma100) if dma100 else False
        a200   = bool(ltp > dma200) if dma200 else False
        pct200 = round((ltp - dma200) / dma200 * 100, 2) if dma200 else None
        hi_vol = bool(vol > avg_vol * HIGH_VOL_MULTIPLIER) if avg_vol else False
        hi_val = bool((val / 1e7) >= HIGH_VALUE_CRORE)
        over   = bool(pct200 > DMA200_OVEREXTENDED_PCT) if pct200 is not None else False

        if   a50 and a100 and a200 and hi_vol and not over: sig = "BUY"
        elif a50 and a100 and a200 and over:                sig = "OVEREXTENDED"
        elif a50 and a100 and a200:                         sig = "HOLD"
        else:                                               sig = "WATCH"

        records.append({
            "symbol": sym, "ltp": round(ltp,2), "prevClose": round(prv,2),
            "change": chg, "changePct": chg_pct,
            "high": round(hgh,2), "low": round(low,2),
            "volume": int(vol), "avgVolume": int(avg_vol) if avg_vol else 0,
            "highVolume": hi_vol,
            "valueCr": round(val/1e7, 2), "highValue": hi_val,
            "dma50": dma50,   "aboveDMA50":  a50,
            "dma100": dma100, "aboveDMA100": a100,
            "dma200": dma200, "aboveDMA200": a200,
            "pctAboveDMA200": pct200,
            "dma200Alert": "OVEREXTENDED" if over else "IN_RANGE",
            "signal": sig,
        })

    df = pd.DataFrame(records)
    df = df.sort_values("pctAboveDMA200", ascending=False,
                        key=lambda x: pd.to_numeric(x, errors="coerce"))
    buy  = (df.signal=='BUY').sum()
    hold = (df.signal=='HOLD').sum()
    ovr  = (df.signal=='OVEREXTENDED').sum()
    wtch = (df.signal=='WATCH').sum()
    print(f"  ✓ {len(df)} stocks | BUY:{buy} HOLD:{hold} OVER:{ovr} WATCH:{wtch}")
    return df

# ── Export JSON ───────────────────────────────────────────────────────────────
def export_json(df, today):
    payload = {
        "generated_at": datetime.now().strftime("%d %b %Y %H:%M IST"),
        "trading_day":  today.strftime("%d %b %Y (%A)"),
        "stocks": df.where(pd.notnull(df), None).to_dict(orient="records"),
    }
    os.makedirs(DOCS_DIR, exist_ok=True)
    path = os.path.join(DOCS_DIR, "screener_data.json")
    with open(path, "w") as f:
        json.dump(payload, f, separators=(',', ':'))
    print(f"  ✓ JSON → {path}")

# ── Export Excel ──────────────────────────────────────────────────────────────
CLR = dict(header="0d1e3b", buy="C6EFCE", hold="FFEB9C",
           over="FFC7CE",   watch="F2F2F2", yes="E2EFDA", no="FCE4D6")

def export_excel(df, today):
    rows = []
    for _, s in df.iterrows():
        rows.append({
            "Symbol":        s.symbol,
            "LTP (₹)":       s.ltp,
            "Prev Close (₹)":s.prevClose,
            "Change (₹)":    s.change,
            "Change (%)":    s.changePct,
            "High (₹)":      s.high,
            "Low (₹)":       s.low,
            "Volume":        s.volume,
            "Avg Vol (20D)": s.avgVolume,
            "High Volume":   "YES" if s.highVolume else "NO",
            "Value (₹ Cr)":  s.valueCr,
            "High Value":    "YES" if s.highValue else "NO",
            "DMA 50 (₹)":    s.dma50,
            "Above DMA 50":  "YES" if s.aboveDMA50  else "NO",
            "DMA 100 (₹)":   s.dma100,
            "Above DMA 100": "YES" if s.aboveDMA100 else "NO",
            "DMA 200 (₹)":   s.dma200,
            "Above DMA 200": "YES" if s.aboveDMA200 else "NO",
            "% vs DMA 200":  s.pctAboveDMA200,
            "DMA 200 Alert": "⚠ OVEREXTENDED (>10%)" if s.dma200Alert=="OVEREXTENDED" else "✓ Within Range",
            "Signal":        s.signal,
        })
    out_df = pd.DataFrame(rows)
    os.makedirs(DOCS_DIR, exist_ok=True)
    arc    = os.path.join(DOCS_DIR, "archive")
    os.makedirs(arc, exist_ok=True)
    dated  = os.path.join(arc, f"NSE_DMA_Screener_{today.strftime('%Y-%m-%d')}.xlsx")
    latest = os.path.join(DOCS_DIR, "latest.xlsx")

    def write(path):
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            out_df.to_excel(w, sheet_name="All Nifty 250", index=False)
            out_df[out_df.Signal.isin(["BUY","HOLD","OVEREXTENDED"])].to_excel(
                w, sheet_name="Screened", index=False)
            out_df[out_df.Signal=="BUY"].to_excel(w, sheet_name="BUY Signals", index=False)
        wb = load_workbook(path)
        for ws in wb.worksheets: _fmt(ws)
        wb.save(path)

    write(dated)
    import shutil; shutil.copy(dated, latest)
    print(f"  ✓ Excel → {latest}")

def _fmt(ws):
    thin = Side(style="thin", color="D0D0D0")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=CLR["header"])
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[1].height = 28
    for col in ws.columns:
        mw = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(mw+4, 22)
    hdrs = {c.value: c.column for c in ws[1]}
    SC   = {"BUY":CLR["buy"],"HOLD":CLR["hold"],"OVEREXTENDED":CLR["over"],"WATCH":CLR["watch"]}
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = bdr
            cell.alignment = Alignment(horizontal="right")
        for col_name in ["Above DMA 50","Above DMA 100","Above DMA 200"]:
            ci = hdrs.get(col_name)
            if ci:
                c = row[ci-1]
                c.fill = PatternFill("solid", fgColor=(CLR["yes"] if c.value=="YES" else CLR["no"]))
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
            c.fill = PatternFill("solid", fgColor=(CLR["over"] if "OVER" in str(c.value) else CLR["buy"]))
            c.font = Font(bold=True)
        ci = hdrs.get("Change (%)")
        if ci:
            c = row[ci-1]
            try:
                v = float(c.value)
                c.font = Font(color=("375623" if v >= 0 else "9C0006"), bold=True)
            except: pass
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

# ── Main ──────────────────────────────────────────────────────────────────────
def run():
    print("="*60)
    print("  NSE Nifty 250 Daily DMA Screener")
    print(f"  {datetime.now().strftime('%d %b %Y %H:%M')}")
    print("="*60)
    today = last_trading_day()
    print(f"\n[1/4] Trading day: {today.strftime('%d %b %Y (%A)')}")
    price_df, session = build_history(today, days=210)
    result_df = screen(today, price_df, session)
    print("\n[4/4] Exporting…")
    export_json(result_df, today)
    export_excel(result_df, today)
    print("\n" + "="*60)
    print("  ✅ Done!")
    print("="*60)

if __name__ == "__main__":
    run()
