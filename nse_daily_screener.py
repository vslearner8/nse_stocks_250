"""
NSE Nifty 250 Daily DMA Screener
=================================
Downloads NSE CM-UDiFF Bhavcopy, calculates DMA 50/100/200,
screens stocks, and outputs:
  1. Excel report       → docs/latest.xlsx
  2. JSON data file     → docs/screener_data.json   (feeds the dashboard)
  3. Dated Excel copy   → docs/archive/NSE_DMA_Screener_YYYY-MM-DD.xlsx

Run:  python nse_daily_screener.py
Deps: pip install requests pandas openpyxl
"""

import os, io, zipfile, json, requests, pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Config ────────────────────────────────────────────────────────────────────
DOCS_DIR    = "./docs"           # GitHub Pages root (everything served from here)
HISTORY_DIR = "./nse_history"    # Cached daily bhavcopy CSVs
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
    "SIEMENS","TATACONSUM","TORNTPHARM","VEDL","ZOMATO","PAYTM","NYKAA","DELHIVERY",
    "POLICYBZR","IRCTC","HDFCLIFE","SBILIFE","ICICIPRULI","BAJAJ-AUTO","BOSCHLTD",
    "BRITANNIA","COLPAL","CONCOR","CUMMINSIND","GLAXO","HAVELLS","HINDPETRO","IOC",
    "MCDOWELL-N","MOTHERSON","MPHASIS","OFSS","PETRONET","PIIND","RAYMOND","SAIL",
    "TATAPOWER","TATASTEEL","TVSMOTOR","UBL","VOLTAS","WHIRLPOOL","ZYDUSLIFE",
    "ABCAPITAL","ACC","ADANIGREEN","ADANIPOWER","AMBUJACEM","AUBANK","BALKRISIND",
    "BANDHANBNK","BANKBARODA","BEL","BPCL","CANFINHOME","CANBK","CHOLAFIN","CUB",
    "DEEPAKNTR","DIXON","DLF","ESCORTS","FEDERALBNK","GAIL","GMRINFRA","GNFC",
    "GODREJPROP","GRINDWELL","HAL","HDFCAMC","ICICIGI","IGL","INDHOTEL","INDIGO",
    "INDUSTOWER","IPCALAB","JUBLFOOD","KAJARIACER","LALPATHLAB","LICHSGFIN",
    "LINDEINDIA","LUPIN","MANAPPURAM","MFSL","MUTHOOTFIN","NATCOPHARM","NAUKRI",
    "NAVINFLUOR","NBCC","NCC","NHPC","NMDC","OBEROIRLTY","PAGEIND","PEL",
    "PERSISTENT","PHOENIXLTD","POLYCAB","RAMCOCEM","RECLTD","RELAXO","RITES",
    "SBICARD","SCHAEFFLER","SRF","STARHEALTH","SUNTV","SUPREMEIND","SYNGENE",
    "TATACOMM","TORNTPOWER","TRENT","UNIONBANK","UPL","VGUARD","WESTLIFE",
    "YESBANK","ZEEL","SUNDARMFIN","HAPPSTMNDS","LATENTVIEW","ANGELONE","FINEORG",
    "CAMPUS","MEDANTA","RADICO","AAVAS","APTUS","BIKAJI","DREAMFOLKS","GLOBALHEALT",
    "KAYNES","MAPMYINDIA","METRO","RRKABEL","UTIAMC",
]

BHAVCOPY_URL = (
    "https://nsearchives.nseindia.com/content/cm/"
    "BhavCopy_NSE_CM_0_0_0_{date}_F_0000.csv.zip"
)
HEADERS = {
    "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept-Encoding":"gzip, deflate",
    "Accept":"*/*",
    "Connection":"keep-alive",
    "Referer":"https://www.nseindia.com/",
}

# ── Helpers ───────────────────────────────────────────────────────────────────
def last_trading_day(ref=None):
    d = ref or datetime.today()
    while d.weekday() >= 5: d -= timedelta(days=1)
    return d

def download_bhavcopy(date_obj):
    date_str   = date_obj.strftime("%Y%m%d")
    cache_path = os.path.join(HISTORY_DIR, f"bhav_{date_str}.csv")
    if os.path.exists(cache_path):
        return pd.read_csv(cache_path)
    url = BHAVCOPY_URL.format(date=date_str)
    print(f"  ↓ {url}")
    session = requests.Session()
    session.get("https://www.nseindia.com", headers=HEADERS, timeout=10)
    resp = session.get(url, headers=HEADERS, timeout=30)
    if resp.status_code != 200:
        raise Exception(f"HTTP {resp.status_code} for {date_str}")
    with zipfile.ZipFile(io.BytesIO(resp.content)) as z:
        csv_name = [n for n in z.namelist() if n.endswith(".csv")][0]
        df = pd.read_csv(z.open(csv_name))
    os.makedirs(HISTORY_DIR, exist_ok=True)
    df.to_csv(cache_path, index=False)
    return df

def build_history(today, days=210):
    print(f"\n[2/4] Building {days}-day price history…")
    closes = {}
    date   = today
    count  = 0
    while count < days:
        try:
            df = download_bhavcopy(date)
            df.columns = [c.strip().upper() for c in df.columns]
            ser_col  = next((c for c in ["SCTYSRS","SERIES"]      if c in df.columns), None)
            sym_col  = next((c for c in ["TCKRSYMB","SYMBOL"]     if c in df.columns), None)
            cls_col  = next((c for c in ["CLSPRIC","CLOSE"]       if c in df.columns), None)
            if not all([ser_col, sym_col, cls_col]): raise ValueError("Missing columns")
            eq = df[df[ser_col] == "EQ"].set_index(sym_col)[cls_col]
            closes[date.strftime("%Y%m%d")] = eq
            count += 1
        except Exception as e:
            print(f"    skip {date.strftime('%Y%m%d')}: {e}")
        date -= timedelta(days=1)
        while date.weekday() >= 5: date -= timedelta(days=1)
    price_df = pd.DataFrame(closes).T.sort_index()
    price_df = price_df.apply(pd.to_numeric, errors="coerce")
    print(f"  ✓ {len(price_df)} days × {len(price_df.columns)} symbols")
    return price_df

def screen(today, price_df, bhav_today):
    print("\n[3/4] Screening…")
    bhav_today.columns = [c.strip().upper() for c in bhav_today.columns]
    ser_col  = next((c for c in ["SCTYSRS","SERIES"]        if c in bhav_today.columns), None)
    sym_col  = next((c for c in ["TCKRSYMB","SYMBOL"]       if c in bhav_today.columns), None)
    cls_col  = next((c for c in ["CLSPRIC","CLOSE"]         if c in bhav_today.columns), None)
    prv_col  = next((c for c in ["PRVSCLSGPRIC","PREVCLOSE"] if c in bhav_today.columns), None)
    hgh_col  = next((c for c in ["HGHPRIC","HIGH"]          if c in bhav_today.columns), None)
    low_col  = next((c for c in ["LWPRIC","LOW"]            if c in bhav_today.columns), None)
    vol_col  = next((c for c in ["TTLTRADGVOL","TOTTRDQTY"] if c in bhav_today.columns), None)
    val_col  = next((c for c in ["TTLTRFVAL","TOTTRDVAL"]   if c in bhav_today.columns), None)

    eq  = bhav_today[bhav_today[ser_col] == "EQ"].set_index(sym_col)
    records = []
    for sym in NIFTY250:
        if sym not in eq.index: continue
        r   = eq.loc[sym]
        ltp = float(r.get(cls_col, 0) or 0)
        prv = float(r.get(prv_col, 0) or 0)
        vol = float(r.get(vol_col, 0) or 0)
        val = float(r.get(val_col, 0) or 0)
        hgh = float(r.get(hgh_col, 0) or 0)
        low = float(r.get(low_col, 0) or 0)

        chg     = round(ltp - prv, 2)
        chg_pct = round((chg / prv * 100) if prv else 0, 2)

        hist    = price_df[sym].dropna() if sym in price_df.columns else pd.Series(dtype=float)
        dma50   = round(hist.tail(50).mean(),  2) if len(hist) >= 50  else None
        dma100  = round(hist.tail(100).mean(), 2) if len(hist) >= 100 else None
        dma200  = round(hist.tail(200).mean(), 2) if len(hist) >= 200 else None
        avg_vol = hist.tail(20).mean() if len(hist) >= 20 else 0

        a50  = bool(ltp > dma50)  if dma50  else False
        a100 = bool(ltp > dma100) if dma100 else False
        a200 = bool(ltp > dma200) if dma200 else False
        pct200 = round((ltp - dma200) / dma200 * 100, 2) if dma200 else None
        hi_vol = bool(vol > avg_vol * HIGH_VOL_MULTIPLIER) if avg_vol else False
        hi_val = bool((val / 1e7) >= HIGH_VALUE_CRORE)
        over   = bool(pct200 > DMA200_OVEREXTENDED_PCT) if pct200 is not None else False

        if a50 and a100 and a200 and hi_vol and not over: sig = "BUY"
        elif a50 and a100 and a200 and over:              sig = "OVEREXTENDED"
        elif a50 and a100 and a200:                       sig = "HOLD"
        else:                                             sig = "WATCH"

        records.append({
            "symbol": sym, "ltp": round(ltp,2), "prevClose": round(prv,2),
            "change": chg, "changePct": chg_pct,
            "high": round(hgh,2), "low": round(low,2),
            "volume": int(vol), "avgVolume": int(avg_vol) if avg_vol else 0,
            "highVolume": hi_vol,
            "valueCr": round(val/1e7, 2), "highValue": hi_val,
            "dma50": dma50, "aboveDMA50": a50,
            "dma100": dma100, "aboveDMA100": a100,
            "dma200": dma200, "aboveDMA200": a200,
            "pctAboveDMA200": pct200,
            "dma200Alert": "OVEREXTENDED" if over else "IN_RANGE",
            "signal": sig,
        })

    df = pd.DataFrame(records)
    df = df.sort_values("pctAboveDMA200", ascending=False,
                        key=lambda x: pd.to_numeric(x, errors="coerce"))
    print(f"  ✓ {len(df)} stocks  |  BUY:{(df.signal=='BUY').sum()}  HOLD:{(df.signal=='HOLD').sum()}  OVER:{(df.signal=='OVEREXTENDED').sum()}  WATCH:{(df.signal=='WATCH').sum()}")
    return df

# ── Export JSON (feeds the live dashboard) ────────────────────────────────────
def export_json(df, today):
    payload = {
        "generated_at": datetime.now().strftime("%d %b %Y %H:%M IST"),
        "trading_day":  today.strftime("%d %b %Y (%A)"),
        "stocks": df.where(pd.notnull(df), None).to_dict(orient="records"),
    }
    path = os.path.join(DOCS_DIR, "screener_data.json")
    with open(path, "w") as f:
        json.dump(payload, f, separators=(',', ':'))
    print(f"  ✓ JSON → {path}")

# ── Export Excel (downloadable from dashboard) ────────────────────────────────
CLR = dict(
    header="0d1e3b", buy="C6EFCE", hold="FFEB9C",
    over="FFC7CE", watch="F2F2F2", yes="E2EFDA", no="FCE4D6",
)

def export_excel(df, today):
    rows = []
    for _, s in df.iterrows():
        rows.append({
            "Symbol":          s.symbol,
            "LTP (₹)":         s.ltp,
            "Prev Close (₹)":  s.prevClose,
            "Change (₹)":      s.change,
            "Change (%)":      s.changePct,
            "High (₹)":        s.high,
            "Low (₹)":         s.low,
            "Volume":          s.volume,
            "Avg Vol (20D)":   s.avgVolume,
            "High Volume":     "YES" if s.highVolume else "NO",
            "Value (₹ Cr)":    s.valueCr,
            "High Value":      "YES" if s.highValue else "NO",
            "DMA 50 (₹)":      s.dma50,
            "Above DMA 50":    "YES" if s.aboveDMA50  else "NO",
            "DMA 100 (₹)":     s.dma100,
            "Above DMA 100":   "YES" if s.aboveDMA100 else "NO",
            "DMA 200 (₹)":     s.dma200,
            "Above DMA 200":   "YES" if s.aboveDMA200 else "NO",
            "% vs DMA 200":    s.pctAboveDMA200,
            "DMA 200 Alert":   "⚠ OVEREXTENDED (>10%)" if s.dma200Alert == "OVEREXTENDED" else "✓ Within Range",
            "Signal":          s.signal,
        })
    out_df = pd.DataFrame(rows)

    os.makedirs(DOCS_DIR, exist_ok=True)
    archive_dir = os.path.join(DOCS_DIR, "archive")
    os.makedirs(archive_dir, exist_ok=True)

    dated_path  = os.path.join(archive_dir, f"NSE_DMA_Screener_{today.strftime('%Y-%m-%d')}.xlsx")
    latest_path = os.path.join(DOCS_DIR, "latest.xlsx")

    def write_wb(path):
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            out_df.to_excel(w, sheet_name="All Nifty 250", index=False)
            out_df[out_df.Signal.isin(["BUY","HOLD","OVEREXTENDED"])].to_excel(
                w, sheet_name="Screened", index=False)
            out_df[out_df.Signal == "BUY"].to_excel(
                w, sheet_name="BUY Signals", index=False)
        wb = load_workbook(path)
        for ws in wb.worksheets: _fmt(ws)
        wb.save(path)

    write_wb(dated_path)
    import shutil; shutil.copy(dated_path, latest_path)
    print(f"  ✓ Excel → {latest_path}")
    print(f"  ✓ Archive → {dated_path}")

def _fmt(ws):
    thin = Side(style="thin", color="D0D0D0")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr  = PatternFill("solid", fgColor=CLR["header"])
    hdrf = Font(bold=True, color="FFFFFF", size=10)
    for cell in ws[1]:
        cell.fill = hdr; cell.font = hdrf
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[1].height = 28
    for col in ws.columns:
        mw = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(mw+4, 22)
    hdrs = {c.value: c.column for c in ws[1]}
    SIG_CLR = {"BUY":CLR["buy"],"HOLD":CLR["hold"],"OVEREXTENDED":CLR["over"],"WATCH":CLR["watch"]}
    for row in ws.iter_rows(min_row=2):
        for cell in row: cell.border = bdr; cell.alignment = Alignment(horizontal="right")
        for col_name, yes_c, no_c in [
            ("Above DMA 50", CLR["yes"], CLR["no"]),
            ("Above DMA 100", CLR["yes"], CLR["no"]),
            ("Above DMA 200", CLR["yes"], CLR["no"]),
        ]:
            ci = hdrs.get(col_name)
            if ci:
                c = row[ci-1]
                c.fill = PatternFill("solid", fgColor=(yes_c if c.value=="YES" else no_c))
                c.font = Font(bold=True)
        ci = hdrs.get("Signal")
        if ci:
            c = row[ci-1]; clr = SIG_CLR.get(str(c.value),"")
            if clr: c.fill = PatternFill("solid", fgColor=clr); c.font = Font(bold=True)
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
    today      = last_trading_day()
    print(f"\n[1/4] Trading day: {today.strftime('%d %b %Y (%A)')}")
    bhav       = download_bhavcopy(today)
    price_df   = build_history(today, days=210)
    result_df  = screen(today, price_df, bhav)
    print("\n[4/4] Exporting…")
    os.makedirs(DOCS_DIR, exist_ok=True)
    export_json(result_df, today)
    export_excel(result_df, today)
    print("\n" + "="*60)
    print("  ✅ Done!")
    print("="*60)

if __name__ == "__main__":
    run()
