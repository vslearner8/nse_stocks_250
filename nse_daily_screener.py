import os
import io
import zipfile
import requests
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
 
# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────
 
OUTPUT_DIR = "./nse_output"          # Where Excel files are saved
HISTORY_DIR = "./nse_history"        # Where daily CSVs are cached
DMA_PERIODS = [50, 100, 200]
HIGH_VOL_MULTIPLIER = 1.5            # Volume must be 1.5× avg to qualify
HIGH_VALUE_CRORE = 5                 # Value must be ≥ ₹5 Cr to qualify (in crores)
DMA200_OVEREXTENDED_PCT = 10         # Flag if price > 10% above DMA200
 
# Nifty 250 symbols (full list — update if index composition changes)
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
    "ABCAPITAL","ABIRLANUVO","ACC","ADANIGREEN","ADANIPOWER","AMBUJACEM","AUBANK",
    "BALKRISIND","BANDHANBNK","BANKBARODA","BEL","BPCL","CANFINHOME","CANBK",
    "CHOLAFIN","CUB","DEEPAKNTR","DIXON","DLF","ESCORTS","FEDERALBNK","GAIL",
    "GMRINFRA","GNFC","GODREJPROP","GRINDWELL","HAL","HDFCAMC","ICICIGI","IGL",
    "INDHOTEL","INDIGO","INDUSTOWER","IPCALAB","JUBLFOOD","KAJARIACER","LALPATHLAB",
    "LICHSGFIN","LINDEINDIA","LUPIN","MANAPPURAM","MFSL","MINDTREE","MUTHOOTFIN",
    "NATCOPHARM","NAUKRI","NAVINFLUOR","NBCC","NCC","NHPC","NMDC","OBEROIRLTY",
    "PAGEIND","PEL","PERSISTENT","PHOENIXLTD","POLYCAB","PVR","RAMCOCEM","RECLTD",
    "RELAXO","RITES","SBICARD","SCHAEFFLER","SRF","STAR","STARHEALTH","SUNTV",
    "SUPREMEIND","SYNGENE","TATACOMM","TORNTPOWER","TRENT","UNIONBANK","UPL",
    "VGUARD","WESTLIFE","YESBANK","ZEEL","SUNDARMFIN","HAPPSTMNDS","LATENTVIEW",
    "ANGELONE","FINEORG","CAMPUS","MEDANTA","RADICO","AAVAS","APTUS","BIKAJI",
    "DREAMFOLKS","GLOBALHEALT","KAYNES","MAPMYINDIA","METRO","RRKABEL","UTIAMC",
]
 
# NSE bhavcopy URL template (new CM-UDiFF format, post July 2024)
BHAVCOPY_URL = (
    "https://nsearchives.nseindia.com/content/cm/"
    "BhavCopy_NSE_CM_0_0_0_{date}_F_0000.csv.zip"
)
 
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept-Encoding": "gzip, deflate",
    "Accept": "*/*",
    "Connection": "keep-alive",
    "Referer": "https://www.nseindia.com/",
}
 
 
# ─────────────────────────────────────────────────────────────────────────────
# STEP 1: Download today's Bhavcopy
# ─────────────────────────────────────────────────────────────────────────────
 
def get_last_trading_day(ref_date=None):
    """Returns the most recent weekday (Mon–Fri) date string in YYYYMMDD format."""
    d = ref_date or datetime.today()
    while d.weekday() >= 5:  # Saturday=5, Sunday=6
        d -= timedelta(days=1)
    return d
 
 
def download_bhavcopy(date_obj):
    """Download and parse NSE CM Bhavcopy for the given date. Returns a DataFrame."""
    date_str = date_obj.strftime("%Y%m%d")
    cache_path = os.path.join(HISTORY_DIR, f"bhav_{date_str}.csv")
 
    # Use cached file if available
    if os.path.exists(cache_path):
        print(f"  ✓ Using cached bhavcopy for {date_str}")
        return pd.read_csv(cache_path)
 
    url = BHAVCOPY_URL.format(date=date_str)
    print(f"  ↓ Downloading: {url}")
 
    # NSE requires a session cookie — first hit the home page
    session = requests.Session()
    session.get("https://www.nseindia.com", headers=HEADERS, timeout=10)
 
    resp = session.get(url, headers=HEADERS, timeout=30)
    if resp.status_code != 200:
        raise Exception(f"Failed to download bhavcopy for {date_str}: HTTP {resp.status_code}")
 
    with zipfile.ZipFile(io.BytesIO(resp.content)) as z:
        csv_name = [n for n in z.namelist() if n.endswith(".csv")][0]
        df = pd.read_csv(z.open(csv_name))
 
    os.makedirs(HISTORY_DIR, exist_ok=True)
    df.to_csv(cache_path, index=False)
    print(f"  ✓ Saved to cache: {cache_path}")
    return df
 
 
# ─────────────────────────────────────────────────────────────────────────────
# STEP 2: Collect 200 trading days of history for DMA calculation
# ─────────────────────────────────────────────────────────────────────────────
 
def build_price_history(today, days_needed=210):
    """
    Collect closing prices for the last `days_needed` trading days.
    Returns a DataFrame: columns = symbols, index = dates (sorted ascending).
    """
    print(f"\n[2/4] Building price history ({days_needed} trading days)…")
    all_closes = {}
    date = today
    collected = 0
 
    while collected < days_needed:
        date_str = date.strftime("%Y%m%d")
        try:
            df = download_bhavcopy(date)
            # Normalise column names (UDiFF format uses uppercase)
            df.columns = [c.strip().upper() for c in df.columns]
 
            # UDiFF columns: TradDt, BizDt, Sgmt, Src, FinInstrmTp, FinInstrmId,
            #                ISIN, TckrSymb, SctySrs, XpryDt, FininstrmActlXpryDt,
            #                StrkPric, OptnTp, FinInstrmNm, OpnPric, HghPric, LwPric,
            #                ClsPric, LastPric, PrvsClsgPric, UndrlygPric, SttlmPric,
            #                OpnIntrst, ChngInOpnIntrst, TtlTradgVol, TtlTrfVal,
            #                TtlNbOfTxsExctd, SsnId, NewBrdLotQty, Rmks
            eq = df[df.get("SCTYSRS", df.get("SERIES", "")) == "EQ"].copy()
            sym_col = "TCKRSYMB" if "TCKRSYMB" in eq.columns else "SYMBOL"
            close_col = "CLSPRIC" if "CLSPRIC" in eq.columns else "CLOSE"
 
            closes = eq.set_index(sym_col)[close_col]
            all_closes[date_str] = closes
            collected += 1
        except Exception as e:
            print(f"    ⚠ Skipped {date_str}: {e}")
 
        date -= timedelta(days=1)
        while date.weekday() >= 5:
            date -= timedelta(days=1)
 
    price_df = pd.DataFrame(all_closes).T.sort_index()
    price_df = price_df.apply(pd.to_numeric, errors="coerce")
    print(f"  ✓ History built: {len(price_df)} days × {len(price_df.columns)} symbols")
    return price_df
 
 
# ─────────────────────────────────────────────────────────────────────────────
# STEP 3: Screen stocks
# ─────────────────────────────────────────────────────────────────────────────
 
def screen_stocks(today, price_df, today_bhav):
    """Apply all screening rules and return a final DataFrame."""
    print("\n[3/4] Applying DMA screening rules…")
 
    today_bhav.columns = [c.strip().upper() for c in today_bhav.columns]
    eq = today_bhav[today_bhav.get("SCTYSRS", today_bhav.get("SERIES", "")) == "EQ"].copy()
 
    sym_col  = "TCKRSYMB" if "TCKRSYMB" in eq.columns else "SYMBOL"
    close_col = "CLSPRIC" if "CLSPRIC" in eq.columns else "CLOSE"
    open_col  = "OPNPRIC" if "OPNPRIC" in eq.columns else "OPEN"
    high_col  = "HGHPRIC" if "HGHPRIC" in eq.columns else "HIGH"
    low_col   = "LWPRIC"  if "LWPRIC"  in eq.columns else "LOW"
    prev_col  = "PRVSCLSGPRIC" if "PRVSCLSGPRIC" in eq.columns else "PREVCLOSE"
    vol_col   = "TTLTRADGVOL" if "TTLTRADGVOL" in eq.columns else "TOTTRDQTY"
    val_col   = "TTLTRFVAL"   if "TTLTRFVAL"   in eq.columns else "TOTTRDVAL"
 
    eq = eq.set_index(sym_col)
    nifty = [s for s in NIFTY250 if s in eq.index]
 
    records = []
    for sym in nifty:
        row = eq.loc[sym]
        ltp       = float(row.get(close_col, 0) or 0)
        prev      = float(row.get(prev_col,  0) or 0)
        volume    = float(row.get(vol_col,   0) or 0)
        value     = float(row.get(val_col,   0) or 0)  # in ₹
        high      = float(row.get(high_col,  0) or 0)
        low       = float(row.get(low_col,   0) or 0)
 
        change    = ltp - prev
        change_pct = (change / prev * 100) if prev else 0
 
        # DMA calculation from history
        hist = price_df[sym].dropna() if sym in price_df.columns else pd.Series(dtype=float)
        dma50  = hist.tail(50).mean()  if len(hist) >= 50  else None
        dma100 = hist.tail(100).mean() if len(hist) >= 100 else None
        dma200 = hist.tail(200).mean() if len(hist) >= 200 else None
 
        above50  = ltp > dma50  if dma50  else False
        above100 = ltp > dma100 if dma100 else False
        above200 = ltp > dma200 if dma200 else False
        pct_above_dma200 = ((ltp - dma200) / dma200 * 100) if dma200 else None
 
        # Average volume = last 20-day average
        avg_vol = price_df[sym].tail(20).mean() if sym in price_df.columns else 0
        high_vol   = volume > avg_vol * HIGH_VOL_MULTIPLIER if avg_vol else False
        high_value = (value / 1e7) >= HIGH_VALUE_CRORE        # convert to crores
 
        all_above  = above50 and above100 and above200
        over_ext   = pct_above_dma200 > DMA200_OVEREXTENDED_PCT if pct_above_dma200 is not None else False
 
        # Signal
        if all_above and high_vol and not over_ext:
            signal = "BUY"
        elif all_above and not high_vol:
            signal = "HOLD"
        elif all_above and over_ext:
            signal = "OVEREXTENDED"
        else:
            signal = "WATCH"
 
        records.append({
            "Symbol":          sym,
            "LTP (₹)":         round(ltp, 2),
            "Prev Close (₹)":  round(prev, 2),
            "Change (₹)":      round(change, 2),
            "Change (%)":      round(change_pct, 2),
            "High (₹)":        round(high, 2),
            "Low (₹)":         round(low, 2),
            "Volume":          int(volume),
            "Avg Vol (20D)":   int(avg_vol) if avg_vol else 0,
            "High Volume":     "YES" if high_vol else "NO",
            "Value (₹ Cr)":    round(value / 1e7, 2),
            "High Value":      "YES" if high_value else "NO",
            "DMA 50 (₹)":      round(dma50, 2)  if dma50  else "N/A",
            "Above DMA 50
