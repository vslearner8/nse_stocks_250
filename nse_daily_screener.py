"""
NSE Nifty LargeMidCap 250 Daily DMA Screener
=============================================
Uses Yahoo Finance exclusively - reliable in GitHub Actions
Fixed: All ticker mappings + accepts yesterday's data as valid
"""

import os, json, time, warnings
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ── Config ─────────────────────────────────────────────────────────────────────
DOCS_DIR                = "./docs"
HISTORY_DIR             = "./nse_history"
SYMBOLS_CACHE_FILE      = "./nse_history/nifty250_symbols_cache.json"
HIGH_VOL_MULTIPLIER     = 1.5
HIGH_VALUE_CRORE        = 5
DMA200_OVEREXTENDED_PCT = 10
HISTORY_DAYS            = 220
# Accept data up to 3 trading days old (handles yfinance delay + weekends)
MAX_DATA_AGE_DAYS       = 3

# ══════════════════════════════════════════════════════════════════════
# YAHOO FINANCE TICKER MAP — ALL KNOWN NSE→YF DIFFERENCES
# ══════════════════════════════════════════════════════════════════════
YF_TICKER_MAP = {
    # Special characters
    "M&M":           "M%26M.NS",
    "BAJAJ-AUTO":    "BAJAJ-AUTO.NS",
    "MCDOWELL-N":    "MCDOWELL-N.NS",
    "L&TFH":         "LTFH.NS",          # L&T Finance Holdings

    # Confirmed different Yahoo tickers
    "LTIM":          "LTIM.NS",
    "TATAMOTORS":    "TATAMOTORS.NS",
    "ZOMATO":        "ZOMATO.NS",
    "AJANTPHARMA":   "AJANTPHARMA.NS",
    "ISEC":          "ICICIBANK.NS",     # ICICI Securities — use parent or skip
    "MTAR":          "MTAR.NS",
    "TCNSBRANDS":    "TCNSBRANDS.NS",
    "AARTI":         "AARTIIND.NS",      # Aarti Industries
    "VODAFONE":      "IDEA.NS",          # Vodafone Idea = IDEA
    "LODHA":         "LODHA.NS",
    "DMART":         "DMART.NS",
    "LICI":          "LICI.NS",
    "BAJAJHFL":      "BAJAJHFL.NS",
    "ATGL":          "ATGL.NS",
    "GMRAIRPORT":    "GMRAIRPORT.NS",
    "PATANJALI":     "PATANJALI.NS",
    "PRESTIGE":      "PRESTIGE.NS",
    "PVRINOX":       "PVRINOX.NS",
    "BSE":           "BSE.NS",
    "IEX":           "IEX.NS",
    "METROBRAND":    "METROBRAND.NS",
    "APLAPOLLO":     "APLAPOLLO.NS",
}

# Symbols to SKIP — truly delisted or not on Yahoo Finance
SKIP_SYMBOLS = {
    "ISEC",        # ICICI Securities - different structure on YF
    "NSLNISP",     # Not available
    "KRISHANA",    # Not available
    "PGHH",        # P&G Health - delisted from NSE
    "NIACL",       # New India Assurance - limited YF data
    "RKFORGE",     # Ramkrishna Forgings - check ticker
    "ROUTE",       # Route Mobile - acquired
    "INDIANB",     # Indian Bank
    "ANANDRATHI",  # Anand Rathi
}

def to_yf(sym):
    """Convert NSE symbol to Yahoo Finance ticker."""
    if sym in SKIP_SYMBOLS:
        return None
    if sym in YF_TICKER_MAP:
        return YF_TICKER_MAP[sym]
    return f"{sym}.NS"

# ══════════════════════════════════════════════════════════════════════
# SYMBOL LISTS (3-tier fallback)
# ══════════════════════════════════════════════════════════════════════
NIFTY_50 = [
    "ADANIENT","ADANIPORTS","APOLLOHOSP","ASIANPAINT","AXISBANK",
    "BAJAJ-AUTO","BAJFINANCE","BAJAJFINSV","BEL","BHARTIARTL",
    "BPCL","BRITANNIA","CIPLA","COALINDIA","DRREDDY",
    "EICHERMOT","GRASIM","HCLTECH","HDFCBANK","HDFCLIFE",
    "HEROMOTOCO","HINDALCO","HINDUNILVR","ICICIBANK","INDUSINDBK",
    "INFY","ITC","JSWSTEEL","KOTAKBANK","LT",
    "LTIM","M&M","MARUTI","NESTLEIND","NTPC",
    "ONGC","POWERGRID","RELIANCE","SBILIFE","SBIN",
    "SHRIRAMFIN","SUNPHARMA","TATACONSUM","TATAMOTORS","TATASTEEL",
    "TCS","TECHM","TITAN","TRENT","ULTRACEMCO","WIPRO",
]

NIFTY_NEXT_50 = [
    "ABB","ADANIGREEN","ADANIPOWER","AMBUJACEM","BAJAJHFL",
    "BANKBARODA","BSE","CANBK","CHOLAFIN","COLPAL",
    "DABUR","DMART","GAIL","GODREJCP","HAVELLS",
    "HDFCAMC","HINDPETRO","ICICIGI","ICICIPRULI","INDUSTOWER",
    "IOC","IRCTC","IRFC","LICI","LODHA",
    "MARICO","MCDOWELL-N","MOTHERSON","MPHASIS","NAUKRI",
    "NMDC","OFSS","OIL","PAGEIND","PFC",
    "PIDILITIND","PIIND","RECLTD","SAIL","SIEMENS",
    "SRF","TATAPOWER","TORNTPHARM","TVSMOTOR","UBL",
    "UNIONBANK","VBL","VEDL","ZOMATO","ZYDUSLIFE",
]

NIFTY_MIDCAP_150 = [
    "AARTI","AAVAS","ABCAPITAL","ABFRL","ACC",
    "AIAENG","AJANTPHARMA","ALKEM","ANGELONE","APLAPOLLO",
    "APTUS","ASTRAL","ATGL","ATUL","AUROPHARMA",
    "AUBANK","BALKRISIND","BANDHANBNK","BHEL","BIOCON",
    "BLUESTARCO","BRIGADE","CAMS","CANFINHOME","CDSL",
    "CESC","CGPOWER","CHAMBLFERT","COFORGE","CONCOR",
    "CROMPTON","CUMMINSIND","DALBHARAT","DEEPAKNTR","DELHIVERY",
    "DIXON","DLF","ELGIEQUIP","EMAMILTD","ENDURANCE",
    "ESCORTS","EXIDEIND","FEDERALBNK","FIVESTAR","FLUOROCHEM",
    "FORTIS","GLAND","GLAXO","GMRAIRPORT","GODFRYPHLP",
    "GODREJIND","GPIL","GRINDWELL","HAL","HUDCO",
    "IEX","IGL","IIFL","INDHOTEL","INDIGO",
    "IPCALAB","JKCEMENT","JSL","JSWENERGY",
    "JUBLFOOD","KAJARIACER","KANSAINER","KARURVYSYA","KAYNES",
    "KEC","KEI","KPITTECH","KRBL","KTKBANK",
    "L&TFH","LALPATHLAB","LATENTVIEW","LICHSGFIN","LINDEINDIA",
    "LUPIN","MANAPPURAM","MAPMYINDIA","MASTEK","MCX",
    "MEDANTA","METROBRAND","MFSL","MGL","MOTILALOFS",
    "MTAR","MUTHOOTFIN","NATCOPHARM","NBCC","NCC",
    "NHPC","NLCINDIA","OBEROIRLTY","OLECTRA",
    "PATANJALI","PERSISTENT","PHOENIXLTD","PNBHOUSING","POLYCAB",
    "POONAWALLA","PRESTIGE","PVRINOX","RADICO","RAILTEL",
    "RAMCOCEM","RAYMOND","REDINGTON","RELAXO","RITES",
    "RRKABEL","SAFARI","SAREGAMA","SBICARD","SCHAEFFLER",
    "SJVN","SKFINDIA","SOBHA","SPARC","STARHEALTH",
    "STLTECH","SUMICHEM","SUNDARMFIN","SUPREMEIND","SUNTV",
    "SYNGENE","TANLA","TATACOMM","TATAELXSI","TATAINVEST",
    "TCNSBRANDS","TEAMLEASE","TIMKEN","TITAGARH","TORNTPOWER",
    "TRITURBINE","UCOBANK","UJJIVANSFB","UNOMINDA","UTIAMC",
    "VGUARD","VMART","WHIRLPOOL","YESBANK","ZEEL",
]

def get_hardcoded_fallback():
    combined = NIFTY_50 + NIFTY_NEXT_50 + NIFTY_MIDCAP_150
    seen = set(); result = []
    for s in combined:
        if s not in seen and s not in SKIP_SYMBOLS:
            seen.add(s); result.append(s)
    return result

# ══════════════════════════════════════════════════════════════════════
# CACHE
# ══════════════════════════════════════════════════════════════════════
def save_cache(symbols):
    os.makedirs(HISTORY_DIR, exist_ok=True)
    with open(SYMBOLS_CACHE_FILE, "w") as f:
        json.dump({
            "symbols":    symbols,
            "count":      len(symbols),
            "saved_at":   datetime.now().strftime("%d %b %Y %H:%M IST"),
            "saved_date": datetime.now().strftime("%Y-%m-%d"),
        }, f, indent=2)
    print(f"  ✓ Cache saved: {len(symbols)} symbols")

def load_cache():
    if not os.path.exists(SYMBOLS_CACHE_FILE): return None, None
    try:
        with open(SYMBOLS_CACHE_FILE) as f: d = json.load(f)
        s = d.get("symbols", [])
        return (s, d.get("saved_at")) if len(s) >= 200 else (None, None)
    except: return None, None

# ══════════════════════════════════════════════════════════════════════
# STEP 1: Get symbols — 3-tier with NSE API
# ══════════════════════════════════════════════════════════════════════
def get_nse_session():
    import requests
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Referer": "https://www.nseindia.com/",
    })
    try:
        s.get("https://www.nseindia.com", timeout=15); time.sleep(2)
        s.get("https://www.nseindia.com/market-data/all-reports", timeout=10)
        time.sleep(1)
    except: pass
    return s

def fetch_live_symbols(session):
    import requests
    urls = [
        ("LARGEMIDCAP250",
         "https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20LARGEMIDCAP%20250",
         200),
        ("NIFTY100",
         "https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20100",
         90),
        ("MIDCAP150",
         "https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20MIDCAP%20150",
         100),
    ]
    s100 = []; smid = []
    for name, url, minc in urls:
        try:
            r = session.get(url, timeout=20)
            if r.status_code != 200:
                print(f"    ⚠ {name}: HTTP {r.status_code}"); time.sleep(1); continue
            syms = [x["symbol"] for x in r.json().get("data", [])
                    if x.get("symbol") and len(x["symbol"]) > 2
                    and x["symbol"] not in SKIP_SYMBOLS]
            print(f"    ✓ {name}: {len(syms)} symbols")
            if "LARGEMIDCAP" in name and len(syms) >= 200: return syms
            if "100" in name:   s100 = syms
            if "MIDCAP" in name: smid = syms
            time.sleep(1)
        except Exception as e:
            print(f"    ⚠ {name}: {e}")
    if len(s100) >= 90 and len(smid) >= 100:
        combined = list(dict.fromkeys(s100 + smid))
        print(f"    ✓ Combined: {len(combined)}")
        return combined
    return []

def get_symbols():
    print("\n[STEP 1] Getting Nifty 250 symbols…")
    session = get_nse_session()
    source  = "Hardcoded Tier 3"

    print("  [Tier 1] NSE Live API…")
    try:
        live = fetch_live_symbols(session)
        if len(live) >= 200:
            print(f"  ✅ Tier 1: {len(live)} live symbols")
            save_cache(live)
            return live, "NSE Live API (Tier 1)"
        print(f"  ⚠ Tier 1: only {len(live)} symbols")
    except Exception as e:
        print(f"  ⚠ Tier 1: {e}")

    print("  [Tier 2] Loading cache…")
    cached, saved_at = load_cache()
    if cached:
        print(f"  ✅ Tier 2: {len(cached)} cached symbols (from {saved_at})")
        return cached, f"Cache Tier 2 ({saved_at})"

    print("  [Tier 3] Hardcoded fallback…")
    fb = get_hardcoded_fallback()
    print(f"  ✅ Tier 3: {len(fb)} symbols")
    return fb, "Hardcoded Tier 3"

# ══════════════════════════════════════════════════════════════════════
# STEP 2: Yahoo Finance — accepts data up to MAX_DATA_AGE_DAYS old
# ══════════════════════════════════════════════════════════════════════
def last_trading_day():
    d = datetime.today()
    while d.weekday() >= 5: d -= timedelta(days=1)
    return d

def is_recent_enough(date_str):
    """Accept data from last 3 trading days (handles yfinance delay)."""
    try:
        data_date = datetime.strptime(date_str, "%Y-%m-%d")
        check_day = last_trading_day()
        for _ in range(MAX_DATA_AGE_DAYS):
            if data_date.date() == check_day.date():
                return True
            check_day -= timedelta(days=1)
            while check_day.weekday() >= 5:
                check_day -= timedelta(days=1)
        return False
    except:
        return False

def fetch_yfinance(symbols):
    import yfinance as yf

    # Filter out skip symbols and get valid tickers
    valid_syms    = [s for s in symbols if s not in SKIP_SYMBOLS]
    ticker_pairs  = [(s, to_yf(s)) for s in valid_syms if to_yf(s)]

    print(f"\n[STEP 2] Yahoo Finance — {len(ticker_pairs)} symbols…")
    end   = datetime.today()
    start = end - timedelta(days=int(HISTORY_DAYS * 1.6))

    syms_list    = [p[0] for p in ticker_pairs]
    tickers_list = [p[1] for p in ticker_pairs]

    all_close  = {}
    all_volume = {}
    all_today  = {}
    failed     = []
    BATCH      = 50

    for i in range(0, len(tickers_list), BATCH):
        bt = tickers_list[i:i+BATCH]
        bs = syms_list[i:i+BATCH]
        print(f"  Batch {i//BATCH+1}/{(len(tickers_list)-1)//BATCH+1}: "
              f"{bs[0]}…{bs[-1]}")
        try:
            raw = yf.download(
                bt,
                start=start.strftime("%Y-%m-%d"),
                end=end.strftime("%Y-%m-%d"),
                auto_adjust=True, progress=False,
                threads=True, group_by="ticker",
            )
            for sym, tk in zip(bs, bt):
                try:
                    if len(bt) == 1:
                        cs=raw["Close"]; vs=raw["Volume"]
                        hs=raw["High"];  ls=raw["Low"]
                    else:
                        if tk not in raw.columns.get_level_values(0):
                            failed.append(sym); continue
                        cs=raw[tk]["Close"]; vs=raw[tk]["Volume"]
                        hs=raw[tk]["High"];  ls=raw[tk]["Low"]

                    cs=cs.dropna(); vs=vs.dropna()
                    if len(cs) < 10:
                        failed.append(sym); continue

                    ld = cs.index[-1]
                    date_str = ld.strftime("%Y-%m-%d")

                    all_close[sym]  = cs
                    all_volume[sym] = vs
                    all_today[sym]  = {
                        "ltp":       round(float(cs.iloc[-1]), 2),
                        "prevClose": round(float(cs.iloc[-2]), 2) if len(cs)>=2 else 0,
                        "high":      round(float(hs.loc[ld]),  2) if ld in hs.index else 0,
                        "low":       round(float(ls.loc[ld]),  2) if ld in ls.index else 0,
                        "volume":    int(float(vs.loc[ld]))        if ld in vs.index else 0,
                        "date":      date_str,
                    }
                except Exception as e:
                    failed.append(sym)
        except Exception as e:
            print(f"  ✗ Batch error: {e}")
        time.sleep(1)

    # Count recent data (within MAX_DATA_AGE_DAYS trading days)
    recent = sum(1 for v in all_today.values() if is_recent_enough(v["date"]))
    total  = len(all_today)
    td_str = last_trading_day().strftime("%Y-%m-%d")

    if failed:
        print(f"  ⚠ Failed ({len(failed)}): {', '.join(failed[:15])}"
              f"{'...' if len(failed)>15 else ''}")

    print(f"  ✓ Fetched: {total} stocks")
    print(f"  ✓ Recent data (≤{MAX_DATA_AGE_DAYS} trading days): {recent}")

    if total == 0:
        print("  ✗ No data at all — failing")
        return None, None, None

    if recent < total * 0.3:
        # Show what dates we got to help diagnose
        dates = sorted(set(v["date"] for v in all_today.values()), reverse=True)
        print(f"  ⚠ Most recent dates in data: {dates[:5]}")
        print(f"  ⚠ Expected: {td_str}")
        print(f"  ⚠ Only {recent}/{total} stocks have recent data")
        print(f"  → Using whatever data we have (best available)")
        # Still use the data — don't fail completely
        # yfinance may just be delayed today

    # Use most common date as the "trading day"
    if all_today:
        date_counts = {}
        for v in all_today.values():
            date_counts[v["date"]] = date_counts.get(v["date"], 0) + 1
        best_date = max(date_counts, key=date_counts.get)
        print(f"  ✓ Most common data date: {best_date} ({date_counts[best_date]} stocks)")

    return all_close, all_volume, all_today

# ══════════════════════════════════════════════════════════════════════
# STEP 3: Screen
# ══════════════════════════════════════════════════════════════════════
def screen(all_close, all_volume, all_today, symbols):
    print(f"\n[STEP 3] Screening {len(all_today)} stocks…")
    records = []

    for sym in symbols:
        if sym not in all_today: continue
        td      = all_today[sym]
        ltp     = td["ltp"]
        prv     = td["prevClose"]
        vol     = td["volume"]
        hgh     = td.get("high", 0)
        low     = td.get("low",  0)
        chg     = round(ltp - prv, 2)
        chg_pct = round((chg / prv * 100) if prv else 0, 2)

        hist    = all_close.get(sym, pd.Series(dtype=float)).dropna()
        hd      = hist.iloc[:-1] if len(hist) > 1 else hist
        dma50   = round(float(hd.tail(50).mean()),  2) if len(hd) >= 50  else None
        dma100  = round(float(hd.tail(100).mean()), 2) if len(hd) >= 100 else None
        dma200  = round(float(hd.tail(200).mean()), 2) if len(hd) >= 200 else None

        vh      = all_volume.get(sym, pd.Series(dtype=float)).dropna()
        avg_vol = float(vh.iloc[:-1].tail(20).mean()) if len(vh) >= 5 else 0
        val_cr  = round(ltp * vol / 1e7, 2)

        a50    = bool(ltp > dma50)  if dma50  else False
        a100   = bool(ltp > dma100) if dma100 else False
        a200   = bool(ltp > dma200) if dma200 else False
        pct200 = round((ltp - dma200) / dma200 * 100, 2) if dma200 else None
        hi_vol = bool(vol > avg_vol * HIGH_VOL_MULTIPLIER) if avg_vol else False
        hi_val = bool(val_cr >= HIGH_VALUE_CRORE)
        over   = bool(pct200 > DMA200_OVEREXTENDED_PCT) if pct200 is not None else False

        if   a50 and a100 and a200 and hi_vol and not over: sig = "BUY"
        elif a50 and a100 and a200 and over:                sig = "OVEREXTENDED"
        elif a50 and a100 and a200:                         sig = "HOLD"
        else:                                               sig = "WATCH"

        records.append({
            "symbol":sym,"ltp":ltp,"prevClose":prv,
            "change":chg,"changePct":chg_pct,
            "high":hgh,"low":low,
            "volume":vol,"avgVolume":int(avg_vol),
            "highVolume":hi_vol,"valueCr":val_cr,"highValue":hi_val,
            "dma50":dma50,"aboveDMA50":a50,
            "dma100":dma100,"aboveDMA100":a100,
            "dma200":dma200,"aboveDMA200":a200,
            "pctAboveDMA200":pct200,
            "dma200Alert":"OVEREXTENDED" if over else "IN_RANGE",
            "signal":sig,"dataDate":td["date"],
        })

    df = pd.DataFrame(records)
    df = df.sort_values("pctAboveDMA200", ascending=False,
                        key=lambda x: pd.to_numeric(x, errors="coerce"))

    print(f"\n  ✅ Screened: {len(df)}/{len(symbols)}")
    print(f"     BUY:{(df.signal=='BUY').sum()}  "
          f"HOLD:{(df.signal=='HOLD').sum()}  "
          f"OVEREXTENDED:{(df.signal=='OVEREXTENDED').sum()}  "
          f"WATCH:{(df.signal=='WATCH').sum()}")

    # Sample verification
    sample = df[df.symbol.isin(["RELIANCE","ADANIPOWER","HDFCBANK","TCS"])][
        ["symbol","ltp","prevClose","dma50","dma200","signal","dataDate"]]
    if not sample.empty:
        print(f"\n  Verification sample:")
        print(sample.to_string(index=False))
    return df

# ══════════════════════════════════════════════════════════════════════
# STEP 4: Export JSON + Excel
# ══════════════════════════════════════════════════════════════════════
def export_json(df, sym_source, total_syms):
    os.makedirs(DOCS_DIR, exist_ok=True)
    # Get most common data date
    date_counts = {}
    for _, r in df.iterrows():
        date_counts[r["dataDate"]] = date_counts.get(r["dataDate"], 0) + 1
    best_date = max(date_counts, key=date_counts.get) if date_counts else "N/A"

    payload = {
        "generated_at":  datetime.now().strftime("%d %b %Y %H:%M IST"),
        "trading_day":   best_date,
        "data_source":   "Yahoo Finance",
        "symbol_source": sym_source,
        "index_total":   total_syms,
        "fetched_total": len(df),
        "stocks": df.where(pd.notnull(df), None).to_dict(orient="records"),
    }
    path = os.path.join(DOCS_DIR, "screener_data.json")
    with open(path, "w") as f:
        json.dump(payload, f, separators=(",",":"))
    print(f"  ✓ JSON  → {path} ({os.path.getsize(path)/1024:.1f} KB)")

CLR = dict(header="0d1e3b", buy="C6EFCE", hold="FFEB9C",
           over="FFC7CE",   watch="F2F2F2", yes="E2EFDA", no="FCE4D6")

def export_excel(df):
    rows = []
    for _, s in df.iterrows():
        rows.append({
            "Symbol":s.symbol, "Date":s.dataDate,
            "LTP (₹)":s.ltp, "Prev Close (₹)":s.prevClose,
            "Change (₹)":s.change, "Change (%)":s.changePct,
            "High (₹)":s.high, "Low (₹)":s.low,
            "Volume":s.volume, "Avg Vol (20D)":s.avgVolume,
            "High Volume":"YES" if s.highVolume else "NO",
            "Value (₹ Cr)":s.valueCr,
            "High Value":"YES" if s.highValue else "NO",
            "DMA 50 (₹)":s.dma50,
            "Above DMA 50":"YES" if s.aboveDMA50 else "NO",
            "DMA 100 (₹)":s.dma100,
            "Above DMA 100":"YES" if s.aboveDMA100 else "NO",
            "DMA 200 (₹)":s.dma200,
            "Above DMA 200":"YES" if s.aboveDMA200 else "NO",
            "% vs DMA 200":s.pctAboveDMA200,
            "DMA 200 Alert":"⚠ OVEREXTENDED (>10%)" if s.dma200Alert=="OVEREXTENDED"
                            else "✓ Within Range",
            "Signal":s.signal,
        })
    out = pd.DataFrame(rows)
    os.makedirs(DOCS_DIR, exist_ok=True)
    arc = os.path.join(DOCS_DIR, "archive"); os.makedirs(arc, exist_ok=True)
    ds     = datetime.today().strftime("%Y-%m-%d")
    dated  = os.path.join(arc,      f"NSE_DMA_Screener_{ds}.xlsx")
    latest = os.path.join(DOCS_DIR, "latest.xlsx")

    def write(path):
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            out.to_excel(w, sheet_name="All Nifty 250", index=False)
            out[out.Signal.isin(["BUY","HOLD","OVEREXTENDED"])].to_excel(
                w, sheet_name="Screened", index=False)
            out[out.Signal=="BUY"].to_excel(w, sheet_name="BUY Signals", index=False)
        wb = load_workbook(path)
        for ws in wb.worksheets: _fmt(ws)
        wb.save(path)

    write(dated)
    import shutil; shutil.copy(dated, latest)
    print(f"  ✓ Excel → {latest}")

def _fmt(ws):
    thin = Side(style="thin", color="D0D0D0")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c in ws[1]:
        c.fill = PatternFill("solid", fgColor=CLR["header"])
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    for col in ws.columns:
        mw = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(mw+4, 24)
    hdrs = {c.value: c.column for c in ws[1]}
    SC   = {"BUY":CLR["buy"],"HOLD":CLR["hold"],
            "OVEREXTENDED":CLR["over"],"WATCH":CLR["watch"]}
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = bdr; cell.alignment = Alignment(horizontal="right")
        for cn in ["Above DMA 50","Above DMA 100","Above DMA 200"]:
            ci = hdrs.get(cn)
            if ci:
                c = row[ci-1]
                c.fill = PatternFill("solid",
                    fgColor=(CLR["yes"] if c.value=="YES" else CLR["no"]))
                c.font = Font(bold=True)
        ci = hdrs.get("Signal")
        if ci:
            c = row[ci-1]; cl = SC.get(str(c.value),"")
            if cl: c.fill=PatternFill("solid",fgColor=cl); c.font=Font(bold=True)
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
                c.font = Font(color=("375623" if v>=0 else "9C0006"), bold=True)
            except: pass
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions

# ══════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════
def run():
    print("="*60)
    print("  NSE Nifty LargeMidCap 250 — Daily DMA Screener")
    print(f"  {datetime.now().strftime('%d %b %Y %H:%M IST')}")
    print("="*60)

    today = last_trading_day()
    print(f"\n  Trading day: {today.strftime('%d %b %Y (%A)')}")

    # Step 1: Symbols
    symbols, sym_source = get_symbols()
    print(f"\n  Symbol source : {sym_source}")
    print(f"  Symbol count  : {len(symbols)}")

    # Step 2: Yahoo Finance (only source — no bhavcopy)
    all_close, all_volume, all_today = fetch_yfinance(symbols)

    if not all_today:
        print("  ✗ FATAL: Yahoo Finance returned no data at all.")
        raise SystemExit(1)

    # Step 3: Screen
    result_df = screen(all_close, all_volume, all_today, symbols)
    if result_df.empty:
        print("  ✗ FATAL: No results after screening.")
        raise SystemExit(1)

    # Step 4: Export
    print("\n[STEP 4] Exporting…")
    export_json(result_df, sym_source, len(symbols))
    export_excel(result_df)

    print("\n" + "="*60)
    print(f"  ✅ COMPLETE!")
    print(f"  Symbol source : {sym_source}")
    print(f"  Symbols       : {len(symbols)}")
    print(f"  Stocks fetched: {len(result_df)}")
    print("="*60)

if __name__ == "__main__":
    run()
