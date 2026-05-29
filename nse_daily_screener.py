"""
NSE Nifty 250 Daily DMA Screener
=================================
Automatically downloads NSE CM-UDiFF Bhavcopy after market close,
calculates DMA 50/100/200, screens stocks, and generates Excel report.

Requirements:
    pip install requests pandas openpyxl

Schedule this script to run daily at 4:00 PM IST on market days.
  - Linux/Mac:  crontab -e  →  0 16 * * 1-5 /usr/bin/python3 /path/to/nse_daily_screener.py
  - Windows:    Task Scheduler → Daily 4:00 PM, Mon–Fri
  - GitHub Actions: See .github/workflows/nse_screener.yml (provided separately)
"""

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
            "Above DMA 50":    "YES" if above50  else "NO",
            "DMA 100 (₹)":     round(dma100, 2) if dma100 else "N/A",
            "Above DMA 100":   "YES" if above100 else "NO",
            "DMA 200 (₹)":     round(dma200, 2) if dma200 else "N/A",
            "Above DMA 200":   "YES" if above200 else "NO",
            "% vs DMA 200":    round(pct_above_dma200, 2) if pct_above_dma200 is not None else "N/A",
            "DMA 200 Alert":   "⚠ OVEREXTENDED (>10%)" if over_ext else "✓ Within Range",
            "Signal":          signal,
        })

    df = pd.DataFrame(records)
    df = df.sort_values("% vs DMA 200", ascending=False, key=lambda x: pd.to_numeric(x, errors="coerce"))
    print(f"  ✓ Screened {len(df)} Nifty 250 stocks")
    print(f"    BUY: {(df['Signal']=='BUY').sum()}  |  HOLD: {(df['Signal']=='HOLD').sum()}  |  OVEREXTENDED: {(df['Signal']=='OVEREXTENDED').sum()}  |  WATCH: {(df['Signal']=='WATCH').sum()}")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# STEP 4: Export to Excel with conditional formatting
# ─────────────────────────────────────────────────────────────────────────────

# Colors
CLR_HEADER   = "1E3A5F"
CLR_BUY      = "C6EFCE"
CLR_HOLD     = "FFEB9C"
CLR_OVEREXT  = "FFC7CE"
CLR_WATCH    = "F2F2F2"
CLR_YES_DMA  = "E2EFDA"
CLR_NO_DMA   = "FCE4D6"
CLR_OVER10   = "FFC7CE"
CLR_WITHIN   = "C6EFCE"


def export_excel(df, today):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    date_str = today.strftime("%Y-%m-%d")
    filename = os.path.join(OUTPUT_DIR, f"NSE_DMA_Screener_{date_str}.xlsx")

    # ── Sheet 1: Full list ──────────────────────────────────────────────────
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="All Nifty 250", index=False)
        df[df["Signal"].isin(["BUY","HOLD","OVEREXTENDED"])].to_excel(
            writer, sheet_name="Screened (DMA+HiVol)", index=False)
        df[df["Signal"] == "BUY"].to_excel(writer, sheet_name="BUY Signals", index=False)

        for sheet_name in writer.sheets:
            _format_sheet(writer.sheets[sheet_name])

    print(f"\n  ✅ Excel saved: {filename}")
    return filename


def _format_sheet(ws):
    # Header row
    header_fill = PatternFill("solid", fgColor=CLR_HEADER)
    header_font = Font(bold=True, color="FFFFFF", size=10)
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    ws.row_dimensions[1].height = 30

    # Column widths
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 22)

    # Find key columns by header name
    headers = {cell.value: cell.column for cell in ws[1]}
    signal_col  = headers.get("Signal")
    alert_col   = headers.get("DMA 200 Alert")
    above50_col = headers.get("Above DMA 50")
    above100_col= headers.get("Above DMA 100")
    above200_col= headers.get("Above DMA 200")
    chg_col     = headers.get("Change (%)")
    highvol_col = headers.get("High Volume")

    SIGNAL_COLORS = {
        "BUY":          CLR_BUY,
        "HOLD":         CLR_HOLD,
        "OVEREXTENDED": CLR_OVEREXT,
        "WATCH":        CLR_WATCH,
    }

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="right")

        # Signal row highlight
        if signal_col:
            sig_cell = row[signal_col - 1]
            sig_val = str(sig_cell.value or "")
            color = SIGNAL_COLORS.get(sig_val)
            if color:
                sig_cell.fill = PatternFill("solid", fgColor=color)
                sig_cell.font = Font(bold=True)

        # Above DMA columns
        for col_idx, yes_color, no_color in [
            (above50_col,  CLR_YES_DMA, CLR_NO_DMA),
            (above100_col, CLR_YES_DMA, CLR_NO_DMA),
            (above200_col, CLR_YES_DMA, CLR_NO_DMA),
        ]:
            if col_idx:
                c = row[col_idx - 1]
                c.fill = PatternFill("solid", fgColor=(yes_color if c.value == "YES" else no_color))
                c.font = Font(bold=True)

        # DMA 200 Alert
        if alert_col:
            c = row[alert_col - 1]
            c.fill = PatternFill("solid", fgColor=(CLR_OVER10 if "OVER" in str(c.value) else CLR_WITHIN))
            c.font = Font(bold=True)

        # Change % color
        if chg_col:
            c = row[chg_col - 1]
            try:
                v = float(c.value)
                c.font = Font(color=("375623" if v >= 0 else "9C0006"), bold=True)
            except:
                pass

        # High Volume
        if highvol_col:
            c = row[highvol_col - 1]
            if c.value == "YES":
                c.font = Font(bold=True, color="7030A0")

    # Freeze top row
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def run():
    print("=" * 60)
    print("  NSE Nifty 250 Daily DMA Screener")
    print(f"  Run date: {datetime.now().strftime('%d %b %Y %H:%M IST')}")
    print("=" * 60)

    today = get_last_trading_day()
    print(f"\n[1/4] Last trading day: {today.strftime('%d %b %Y (%A)')}")

    # Download today's bhavcopy
    today_bhav = download_bhavcopy(today)

    # Build 200-day history for DMA
    price_df = build_price_history(today, days_needed=210)

    # Screen
    result_df = screen_stocks(today, price_df, today_bhav)

    # Export
    print("\n[4/4] Exporting Excel report…")
    excel_path = export_excel(result_df, today)

    print("\n" + "=" * 60)
    print(f"  ✅ DONE! Report saved to: {excel_path}")
    print("=" * 60)


if __name__ == "__main__":
    run()

