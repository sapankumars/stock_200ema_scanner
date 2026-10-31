"""
200 EMA Crossover Scanner - Daily, Weekly & Monthly Timeframe
================================================================
Pura NSE stock list scan karke jo stocks apni 200 EMA (Exponential
Moving Average) ko UPAR ya NEECHE cross kar rahe hain (latest candle
par), unko yeh script find karta hai — Daily, Weekly aur Monthly,
teeno timeframe par — aur result Excel file mein 3 alag sheets
(EMA200_Daily, EMA200_Weekly, EMA200_Monthly) mein save karta hai.
Har sheet ke andar Bullish aur Bearish stocks alag-alag block mein
dikhte hain.

NEW ENTRY TRACKING:
Script har run ke baad apna result "ema200_history.json" file mein
save kar leta hai. Agle run mein jo stock pichhli baar list mein
NAHI tha (naya cross hua hai), uski row Excel mein YELLOW highlight
hogi aur "New Entry" column mein "NEW" likha aayega. Is history file
ko delete mat karo agar aap yeh tracking chalu rakhna chahte ho.

Requirements:
    pip install yfinance pandas openpyxl requests

Usage:
    python stock_200ema_scanner.py
"""

import pandas as pd
import yfinance as yf
import requests
import io
import json
import os
import time
from datetime import datetime

EMA_PERIOD = 200

# History file: pichhle run ke signals yahan save hote hain, taaki
# agle run mein pata chal sake ki koi NAYA stock list mein aaya hai
# ya nahi (jo pichhli baar nahi tha).
HISTORY_FILE = "ema200_history.json"

# ---------------------------------------------------------------------
# 1. NSE ki poori equity list fetch karna
# ---------------------------------------------------------------------

NSE_EQUITY_LIST_URL = "https://archives.nseindia.com/content/equities/EQUITY_L.csv"
BACKUP_FILE = "nse_stock_list_backup.txt"


def get_nse_stock_list():
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        resp = requests.get(NSE_EQUITY_LIST_URL, headers=headers, timeout=15)
        resp.raise_for_status()
        df = pd.read_csv(io.StringIO(resp.text))
        symbols = [f"{s.strip()}.NS" for s in df["SYMBOL"].tolist()]
        print(f"NSE se {len(symbols)} stocks fetch hue.")

        with open(BACKUP_FILE, "w") as f:
            f.write("\n".join(symbols))

        return symbols

    except Exception as e:
        print(f"[Warning] NSE list download fail hui ({e}).")
        try:
            with open(BACKUP_FILE, "r") as f:
                symbols = [line.strip() for line in f if line.strip()]
            print(f"Backup file se {len(symbols)} stocks load hue.")
            return symbols
        except FileNotFoundError:
            print("Backup file bhi nahi mili. Please 'nse_stock_list_backup.txt' "
                  "banake usme symbols daal do (e.g. RELIANCE.NS ek per line).")
            return []


# ---------------------------------------------------------------------
# 1a. Helper: yfinance ka data hamesha "clean" single-level columns
#     ke saath milta hai (yeh naya yfinance MultiIndex columns deta
#     hai, jisse "truth value of a Series is ambiguous" error aata
#     hai agar isko flatten na karein).
# ---------------------------------------------------------------------

def flatten_columns(data):
    if isinstance(data.columns, pd.MultiIndex):
        # yfinance ka format: (field, ticker) -> sirf 'field' rakho
        data = data.copy()
        data.columns = data.columns.get_level_values(0)
    return data


# ---------------------------------------------------------------------
# 2. EMA crossover check
# ---------------------------------------------------------------------

def check_ema_crossover(df, period=EMA_PERIOD):
    """
    df: OHLC dataframe with 'Close' column (single-level columns)
    Returns: (signal, last_close, last_date)
    signal = 'Crossed ABOVE 200 EMA' / 'Crossed BELOW 200 EMA' / None
    """
    if len(df) < period + 2:
        return None, None, None

    df = df.copy()

    close = df["Close"]
    # Safety: agar phir bhi 2D aa jaye (duplicate column names waghera)
    if isinstance(close, pd.DataFrame):
        close = close.iloc[:, 0]

    df["EMA"] = close.ewm(span=period, adjust=False).mean()
    df["Close_"] = close
    df = df.dropna(subset=["Close_", "EMA"])

    if len(df) < 2:
        return None, None, None

    prev_close, prev_ema = float(df["Close_"].iloc[-2]), float(df["EMA"].iloc[-2])
    curr_close, curr_ema = float(df["Close_"].iloc[-1]), float(df["EMA"].iloc[-1])

    signal = None
    if prev_close <= prev_ema and curr_close > curr_ema:
        signal = "Crossed ABOVE 200 EMA (Bullish)"
    elif prev_close >= prev_ema and curr_close < curr_ema:
        signal = "Crossed BELOW 200 EMA (Bearish)"

    last_date = df.index[-1]
    return signal, curr_close, last_date


def scan_timeframe(symbols, interval_label, resample_rule=None):
    results = []
    total = len(symbols)

    for i, symbol in enumerate(symbols, 1):
        try:
            data = yf.download(symbol, period="5y", interval="1d",
                                progress=False, auto_adjust=True)
            if data is None or data.empty:
                continue

            data = flatten_columns(data)

            if resample_rule:
                data = data.resample(resample_rule).agg({
                    "Open": "first", "High": "max",
                    "Low": "min", "Close": "last",
                    "Volume": "sum"
                }).dropna()

            signal, last_close, last_date = check_ema_crossover(data)

            if signal:
                results.append({
                    "Symbol": symbol,
                    "Timeframe": interval_label,
                    "Signal": signal,
                    "Last Close": round(float(last_close), 2),
                    "Date": last_date.strftime("%Y-%m-%d")
                })
                print(f"  [FOUND] {symbol} -> {signal} ({interval_label})")

            if i % 100 == 0:
                print(f"  ... {i}/{total} stocks scanned ({interval_label})")

        except Exception as e:
            print(f"  [error] {symbol}: {e}")

        # Thoda delay taaki Yahoo Finance rate-limit na kare
        time.sleep(0.1)

    return results


def load_history():
    """
    Pichhle run ki history load karta hai.
    Returns: set of (Symbol, Timeframe, Signal, Date) tuples.
    """
    if not os.path.exists(HISTORY_FILE):
        return set()
    try:
        with open(HISTORY_FILE, "r") as f:
            data = json.load(f)
        return set(tuple(item) for item in data)
    except Exception as e:
        print(f"[Warning] History file read nahi ho payi ({e}). Fresh maan kar chal rahe hain.")
        return set()


def save_history(all_results):
    """
    Is run ke saare results (Daily+Weekly+Monthly) ko history file
    mein save karta hai, taaki agli baar naye entries pehchane ja sakein.
    """
    tuples = [
        (r["Symbol"], r["Timeframe"], r["Signal"], r["Date"])
        for r in all_results
    ]
    with open(HISTORY_FILE, "w") as f:
        json.dump(tuples, f, indent=2)


def split_bull_bear(results, timeframe_label, prev_history):
    """
    results (list of dicts) ko Bullish aur Bearish do DataFrames mein
    split karta hai, aur ek "New Entry" column add karta hai jo batata
    hai ki yeh signal pichhle run mein nahi tha (naya hai).
    """
    cols = ["Symbol", "Signal", "Last Close", "Date", "New Entry"]
    if not results:
        empty = pd.DataFrame(columns=cols)
        return empty, empty

    df = pd.DataFrame(results)
    df["New Entry"] = df.apply(
        lambda r: "NEW" if (r["Symbol"], timeframe_label, r["Signal"], r["Date"])
        not in prev_history else "",
        axis=1
    )
    df = df[["Symbol", "Signal", "Last Close", "Date", "New Entry"]]

    df_bull = df[df["Signal"].str.contains("Bullish")] \
        .sort_values(["New Entry", "Symbol"], ascending=[False, True]).reset_index(drop=True)
    df_bear = df[df["Signal"].str.contains("Bearish")] \
        .sort_values(["New Entry", "Symbol"], ascending=[False, True]).reset_index(drop=True)
    return df_bull, df_bear


def write_bull_bear_sheet(writer, sheet_name, df_bull, df_bear):
    """
    Ek hi sheet mein Bullish list left side (col A) aur Bearish list
    right side (col F se) mein, apne apne heading ke saath likhta hai.
    """
    from openpyxl.styles import Font, PatternFill

    workbook = writer.book
    worksheet = workbook.create_sheet(sheet_name)

    BULL_START_COL = 1   # A
    BEAR_START_COL = 6   # F  (1 col gap after 4-col bullish block)
    HEADER_ROW = 1
    TITLE_ROW = 0

    new_bull_count = int((df_bull["New Entry"] == "NEW").sum()) if len(df_bull) else 0
    new_bear_count = int((df_bear["New Entry"] == "NEW").sum()) if len(df_bear) else 0

    bull_title = worksheet.cell(
        row=1, column=BULL_START_COL,
        value=f"BULLISH (Crossed ABOVE) - {len(df_bull)} stocks  |  {new_bull_count} NEW")
    bear_title = worksheet.cell(
        row=1, column=BEAR_START_COL,
        value=f"BEARISH (Crossed BELOW) - {len(df_bear)} stocks  |  {new_bear_count} NEW")
    bull_title.font = Font(bold=True, color="006100")
    bear_title.font = Font(bold=True, color="9C0006")
    bull_title.fill = PatternFill("solid", fgColor="C6EFCE")
    bear_title.fill = PatternFill("solid", fgColor="FFC7CE")

    cols = ["Symbol", "Signal", "Last Close", "Date", "New Entry"]
    new_fill = PatternFill("solid", fgColor="FFFF00")  # yellow highlight for NEW rows
    new_font = Font(bold=True, color="9C6500")

    def write_block(df, start_col):
        header_row = 2
        for j, col_name in enumerate(cols):
            c = worksheet.cell(row=header_row, column=start_col + j, value=col_name)
            c.font = Font(bold=True)
        for i, row in enumerate(df.itertuples(index=False), start=header_row + 1):
            is_new = (row[-1] == "NEW")  # "New Entry" is last column
            for j, val in enumerate(row):
                cell = worksheet.cell(row=i, column=start_col + j, value=val)
                if is_new:
                    cell.fill = new_fill
                    if j == len(row) - 1:
                        cell.font = new_font

    write_block(df_bull, BULL_START_COL)
    write_block(df_bear, BEAR_START_COL)

    # Column widths thoda adjust
    widths = {"A": 16, "B": 24, "C": 12, "D": 12, "E": 11,
              "F": 16, "G": 24, "H": 12, "I": 12, "J": 11}
    for col, w in widths.items():
        worksheet.column_dimensions[col].width = w



# ---------------------------------------------------------------------
# Telegram: credentials GitHub Secrets se environment variables ke
# through aate hain (TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID).
# ---------------------------------------------------------------------

def send_telegram_message(text):
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        print("[Warning] Telegram secrets set nahi hain, message skip.")
        return
    try:
        r = requests.post(f"https://api.telegram.org/bot{token}/sendMessage",
                          data={"chat_id": chat_id, "text": text}, timeout=30)
        r.raise_for_status()
    except Exception as e:
        print(f"[Warning] Telegram message fail: {e}")


def send_telegram_file(file_path, caption=""):
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        print("[Warning] Telegram secrets set nahi hain, file skip.")
        return
    try:
        with open(file_path, "rb") as f:
            r = requests.post(f"https://api.telegram.org/bot{token}/sendDocument",
                              data={"chat_id": chat_id, "caption": caption},
                              files={"document": f}, timeout=120)
        r.raise_for_status()
        print("Telegram par Excel file bhej di gayi.")
    except Exception as e:
        print(f"[Warning] Telegram file fail: {e}")


def main():
    symbols = get_nse_stock_list()
    if not symbols:
        print("Stock list khali hai, script rok raha hoon.")
        send_telegram_message("200 EMA Scanner: stock list nahi mili, scan nahi hua.")
        return

    print(f"\nScanning {len(symbols)} stocks for 200 EMA crossover...\n")

    print("== Daily Timeframe Scan ==")
    daily_results = scan_timeframe(symbols, "Daily", resample_rule=None)

    print("\n== Weekly Timeframe Scan ==")
    weekly_results = scan_timeframe(symbols, "Weekly", resample_rule="W")

    print("\n== Monthly Timeframe Scan ==")
    monthly_results = scan_timeframe(symbols, "Monthly", resample_rule="ME")

    if not (daily_results or weekly_results or monthly_results):
        print("\nKoi 200 EMA crossover nahi mila is baar.")
        send_telegram_message("200 EMA Scanner: aaj koi crossover nahi mila.")
        return

    # Pichhle run ki history load karo (naye entries pehchaanne ke liye)
    prev_history = load_history()

    daily_bull, daily_bear = split_bull_bear(daily_results, "Daily", prev_history)
    weekly_bull, weekly_bear = split_bull_bear(weekly_results, "Weekly", prev_history)
    monthly_bull, monthly_bear = split_bull_bear(monthly_results, "Monthly", prev_history)

    # Is run ka poora data history mein save kar do (agle run ke liye)
    save_history(daily_results + weekly_results + monthly_results)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_file = f"ema200_crossover_results_{timestamp}.xlsx"

    # Empty placeholder DataFrame use karke ExcelWriter create karte hain,
    # phir actual sheets openpyxl ke through banate hain (custom layout ke liye).
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        write_bull_bear_sheet(writer, "EMA200_Daily", daily_bull, daily_bear)
        write_bull_bear_sheet(writer, "EMA200_Weekly", weekly_bull, weekly_bear)
        write_bull_bear_sheet(writer, "EMA200_Monthly", monthly_bull, monthly_bear)
        # pandas ExcelWriter default ek empty "Sheet" bana deta hai, use hata do
        if "Sheet" in writer.book.sheetnames:
            del writer.book["Sheet"]

    def new_count(df):
        return int((df["New Entry"] == "NEW").sum()) if len(df) else 0

    print(f"\nDone! Results saved to: {out_file}")
    print(f"  Daily   -> Bullish: {len(daily_bull)} ({new_count(daily_bull)} NEW), "
          f"Bearish: {len(daily_bear)} ({new_count(daily_bear)} NEW)")
    print(f"  Weekly  -> Bullish: {len(weekly_bull)} ({new_count(weekly_bull)} NEW), "
          f"Bearish: {len(weekly_bear)} ({new_count(weekly_bear)} NEW)")
    print(f"  Monthly -> Bullish: {len(monthly_bull)} ({new_count(monthly_bull)} NEW), "
          f"Bearish: {len(monthly_bear)} ({new_count(monthly_bear)} NEW)")

    caption = (
        "200 EMA Scan complete\n"
        f"Daily: {len(daily_bull)} Bull ({new_count(daily_bull)} NEW) / "
        f"{len(daily_bear)} Bear ({new_count(daily_bear)} NEW)\n"
        f"Weekly: {len(weekly_bull)} Bull ({new_count(weekly_bull)} NEW) / "
        f"{len(weekly_bear)} Bear ({new_count(weekly_bear)} NEW)\n"
        f"Monthly: {len(monthly_bull)} Bull ({new_count(monthly_bull)} NEW) / "
        f"{len(monthly_bear)} Bear ({new_count(monthly_bear)} NEW)"
    )
    send_telegram_file(out_file, caption)


if __name__ == "__main__":
    main()