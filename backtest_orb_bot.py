"""
Intraday ORB + VWAP Strategy BACKTEST (Updated with your requests)
"""

import argparse
import os
from datetime import datetime, timedelta
from io import StringIO

import pandas as pd
import pytz
import requests
import yfinance as yf
from openpyxl import Workbook

# ============================== CONFIG ===================================

BACKTEST_CAPITAL = 300000.0
RISK_PER_TRADE_PCT = 10.0
MIN_RR_RATIO = 1.5

OPENING_RANGE_MINUTES = 15
MAX_ENTRY_HOUR = 13          # No entries after 1:00 PM
CANDLE_INTERVAL = "5m"
MAX_LOOKBACK_DAYS = 59

VOLUME_MULTIPLIER = 2.0      # Strengthened
EMA_FAST, EMA_SLOW = 9, 21
RSI_PERIOD = 14
RSI_BULLISH_MIN, RSI_BULLISH_MAX = 50, 70
RSI_BEARISH_MIN, RSI_BEARISH_MAX = 30, 50
MACD_FAST, MACD_SLOW, MACD_SIGNAL = 12, 26, 9
SWING_LOOKBACK_CANDLES = 12

IST = pytz.timezone("Asia/Kolkata")
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "backtest_journal.xlsx")
NIFTY100_SOURCE_URL = "https://niftyindices.com/IndexConstituent/ind_nifty100list.csv"

FALLBACK_TICKERS = [
    "RELIANCE.NS", "TCS.NS", "HDFCBANK.NS", "ICICIBANK.NS", "INFY.NS",
    "HINDUNILVR.NS", "ITC.NS", "SBIN.NS", "BHARTIARTL.NS", "KOTAKBANK.NS",
    "LT.NS", "BAJFINANCE.NS", "HCLTECH.NS", "AXISBANK.NS", "ASIANPAINT.NS",
    "MARUTI.NS", "SUNPHARMA.NS", "TITAN.NS", "ULTRACEMCO.NS", "WIPRO.NS",
]

# ============================ END CONFIG ==================================

def get_nifty_100_tickers():
    try:
        resp = requests.get(NIFTY100_SOURCE_URL, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
        resp.raise_for_status()
        df = pd.read_csv(StringIO(resp.text))
        symbol_col = "Symbol" if "Symbol" in df.columns else df.columns[2]
        tickers = [f"{s.strip()}.NS" for s in df[symbol_col].dropna().tolist()]
        print(f"Fetched live Nifty 100 list: {len(tickers)} tickers.")
        return tickers
    except Exception as e:
        print(f"[warn] Live Nifty 100 fetch failed ({e}). Using fallback.")
        return FALLBACK_TICKERS


# ... [All indicator functions remain the same: compute_vwap, ema, rsi, macd, get_opening_range, classic_pivots] ...

def simulate_ticker_day(day_df, pivots, capital_start_of_day):
    if len(day_df) < max(EMA_SLOW, MACD_SLOW, RSI_PERIOD) // 3 + 4:
        return None

    day_df = compute_vwap(day_df)
    day_df = compute_ema(day_df, EMA_FAST, "EMA_FAST")
    day_df = compute_ema(day_df, EMA_SLOW, "EMA_SLOW")
    day_df = compute_rsi(day_df, RSI_PERIOD)
    day_df = compute_macd(day_df, MACD_FAST, MACD_SLOW, MACD_SIGNAL)

    or_high, or_low = get_opening_range(day_df)
    if or_high is None:
        return None

    session_start = day_df.index[0]
    or_cutoff = session_start + pd.Timedelta(minutes=OPENING_RANGE_MINUTES)

    avg_volume_full_day = day_df["Volume"].mean()
    if pd.isna(avg_volume_full_day) or avg_volume_full_day == 0:
        return None

    entry_idx = None
    signal = None

    for i in range(len(day_df)):
        row = day_df.iloc[i]
        if row.name <= or_cutoff:
            continue

        # === No trade after 1:00 PM ===
        if row.name.hour >= MAX_ENTRY_HOUR:
            break

        avg_volume = day_df["Volume"].iloc[:i].mean() if i > 0 else avg_volume_full_day
        if pd.isna(avg_volume) or avg_volume == 0:
            continue

        strong_volume = row["Volume"] >= VOLUME_MULTIPLIER * avg_volume

        bull_trend_ok = row["EMA_FAST"] > row["EMA_SLOW"]
        bear_trend_ok = row["EMA_FAST"] < row["EMA_SLOW"]
        bull_rsi_ok = RSI_BULLISH_MIN <= row["RSI"] <= RSI_BULLISH_MAX
        bear_rsi_ok = RSI_BEARISH_MIN <= row["RSI"] <= RSI_BEARISH_MAX
        bull_macd_ok = row["MACD"] > row["MACD_SIGNAL"]
        bear_macd_ok = row["MACD"] < row["MACD_SIGNAL"]

        price_above_vwap = row["Close"] > row["VWAP"]
        price_below_vwap = row["Close"] < row["VWAP"]

        if (row["Close"] > or_high and price_above_vwap and strong_volume
                and bull_trend_ok and bull_rsi_ok and bull_macd_ok):
            entry_idx, signal = i, "BULLISH"
            break
        elif (row["Close"] < or_low and price_below_vwap and strong_volume
                and bear_trend_ok and bear_rsi_ok and bear_macd_ok):
            entry_idx, signal = i, "BEARISH"
            break

    if entry_idx is None:
        return None

    # ... [rest of entry calculation, stop/target, walk-forward remains same] ...

    # Square off at ~2:30 PM
    if outcome is None:
        square_off_time = None
        for j in range(entry_idx + 1, len(day_df)):
            future_row = day_df.iloc[j]
            if future_row.name.hour >= 14 and future_row.name.minute >= 30:
                square_off_time = future_row.name
                break
        if square_off_time is None:
            last_row = day_df.iloc[-1]
            square_off_time = last_row.name
            future_row = last_row
        else:
            future_row = day_df[day_df.index == square_off_time].iloc[0]

        exit_price = float(future_row["Close"])
        exit_time = square_off_time
        if signal == "BULLISH":
            outcome = "SQUARED_OFF_WIN" if exit_price > price else "SQUARED_OFF_LOSS"
        else:
            outcome = "SQUARED_OFF_WIN" if exit_price < price else "SQUARED_OFF_LOSS"

    # ... [PnL calculation same] ...

    return { ... }   # same as before


# Rest of the file (run_backtest, write_excel, main, etc.) remains unchanged
