"""
9/20 EMA Crossover + Volume + RSI Strategy BACKTEST
Same structure as previous scripts for easy deployment.
"""

import argparse
import os
import sys
import threading
import time
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer
from io import StringIO
from urllib.parse import urlparse, parse_qs

import pandas as pd
import pytz
import requests
import yfinance as yf
from openpyxl import Workbook

# ============================== CONFIG ===================================

BACKTEST_CAPITAL = 300000.0
RISK_PER_TRADE_PCT = 10.0
MAX_DAILY_RISK_PCT = 30.0

CANDLE_INTERVAL = "5m"
MAX_LOOKBACK_DAYS = 59

# Strategy Parameters
EMA_FAST = 9
EMA_SLOW = 20
RSI_PERIOD = 14
RSI_MAX_FOR_LONG = 65      # Avoid entering when overbought
VOL_MULT = 1.4
SL_PCT = 0.5
RR_TARGET = 2.0

MAX_ENTRY_HOUR, MAX_ENTRY_MINUTE = 13, 30
EXIT_HOUR, EXIT_MINUTE = 14, 45

IST = pytz.timezone("Asia/Kolkata")
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "ema_crossover_journal.xlsx")
NIFTY100_SOURCE_URL = "https://niftyindices.com/IndexConstituent/ind_nifty100list.csv"

FALLBACK_TICKERS = ["RELIANCE.NS", "HDFCBANK.NS", "ICICIBANK.NS", "SBIN.NS", "BHARTIARTL.NS"]

# ============================ END CONFIG ==================================

def get_nifty_100_tickers():
    try:
        resp = requests.get(NIFTY100_SOURCE_URL, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
        df = pd.read_csv(StringIO(resp.text))
        symbol_col = "Symbol" if "Symbol" in df.columns else df.columns[2]
        return [f"{s.strip()}.NS" for s in df[symbol_col].dropna().tolist()]
    except:
        return FALLBACK_TICKERS

def fetch_intraday_history(ticker, days):
    try:
        df = yf.download(ticker, period=f"{days}d", interval=CANDLE_INTERVAL, progress=False)
        if df.empty or len(df) < 100:
            return None
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)
        return df
    except:
        return None

def simulate_ticker(ticker, df):
    df = df.copy()
    df['EMA9'] = df['Close'].ewm(span=EMA_FAST, adjust=False).mean()
    df['EMA20'] = df['Close'].ewm(span=EMA_SLOW, adjust=False).mean()
    delta = df['Close'].diff()
    gain = delta.where(delta > 0, 0).rolling(RSI_PERIOD).mean()
    loss = -delta.where(delta < 0, 0).rolling(RSI_PERIOD).mean()
    rs = gain / loss
    df['RSI'] = 100 - (100 / (1 + rs))
    df['AvgVol'] = df['Volume'].rolling(10).mean()

    trades = []
    in_position = False
    entry_price = stop_price = target_price = None
    entry_time = None
    current_day = None

    for i in range(30, len(df)):
        row = df.iloc[i]
        row_time = row.name.time()
        row_day = row.name.date()

        if row_day != current_day:
            current_day = row_day
            if in_position:
                exit_price = float(row['Close'])
                trades.append(_build_trade(ticker, "LONG", entry_time, entry_price, stop_price, target_price, row.name, exit_price, "EOD"))
                in_position = False

        if in_position:
            if row_time >= pd.Timestamp(f"{EXIT_HOUR}:{EXIT_MINUTE}").time():
                exit_price = float(row['Close'])
                trades.append(_build_trade(ticker, "LONG", entry_time, entry_price, stop_price, target_price, row.name, exit_price, "TIME_EXIT"))
                in_position = False
                continue
            price = float(row['Close'])
            if price >= target_price:
                trades.append(_build_trade(ticker, "LONG", entry_time, entry_price, stop_price, target_price, row.name, price, "TARGET_HIT"))
                in_position = False
            elif price <= stop_price:
                trades.append(_build_trade(ticker, "LONG", entry_time, entry_price, stop_price, target_price, row.name, price, "STOP_HIT"))
                in_position = False
            continue

        # Crossover + Filters
        crossover = (row['EMA9'] > row['EMA20']) and (df.iloc[i-1]['EMA9'] <= df.iloc[i-1]['EMA20'])
        rsi_ok = row['RSI'] < RSI_MAX_FOR_LONG
        vol_ok = row['Volume'] > VOL_MULT * row['AvgVol']
        early_enough = row_time <= pd.Timestamp(f"{MAX_ENTRY_HOUR}:{MAX_ENTRY_MINUTE}").time()

        if crossover and rsi_ok and vol_ok and early_enough:
            entry_price = float(row['Close'])
            stop_price = entry_price * (1 - SL_PCT / 100)
            risk = entry_price - stop_price
            target_price = entry_price + RR_TARGET * risk
            entry_time = row.name
            in_position = True

    return trades

def _build_trade(ticker, direction, entry_time, entry_price, stop_price, target, exit_time, exit_price, reason):
    risk = entry_price - stop_price
    return {
        "EntryDate": entry_time.strftime("%Y-%m-%d"),
        "EntryTime": entry_time.strftime("%H:%M:%S"),
        "Ticker": ticker,
        "Direction": direction,
        "Entry": round(entry_price, 2),
        "InitialStop": round(stop_price, 2),
        "Target": round(target, 2),
        "RiskPerShare": round(risk, 2),
        "ExitTime": exit_time.strftime("%H:%M:%S"),
        "ExitPrice": round(exit_price, 2),
        "ExitReason": reason,
    }

def apply_position_sizing(raw_trades):
    # Same logic as previous scripts
    by_day = {}
    for t in raw_trades:
        by_day.setdefault(t["EntryDate"], []).append(t)

    capital = BACKTEST_CAPITAL
    final_trades = []
    capital_curve = [capital]

    for day in sorted(by_day.keys()):
        day_trades = sorted(by_day[day], key=lambda t: t["EntryTime"])
        capital_start = capital
        committed = 0.0
        day_pnl = 0.0
        taken = 0

        for t in day_trades:
            risk_amount = capital_start * (RISK_PER_TRADE_PCT / 100)
            if committed + risk_amount > capital_start * (MAX_DAILY_RISK_PCT / 100):
                continue
            committed += risk_amount

            qty = max(1, int(risk_amount // t["RiskPerShare"])) if t["RiskPerShare"] > 0 else 0
            if qty == 0: continue

            pnl = (t["ExitPrice"] - t["Entry"]) * qty
            pnl_pct = round((pnl / (t["Entry"] * qty)) * 100, 2) if qty else 0
            outcome = "WIN" if pnl > 0 else "LOSS"

            t["Qty"] = qty
            t["RiskAmount"] = round(risk_amount, 2)
            t["PnL"] = round(pnl, 2)
            t["PnLPct"] = pnl_pct
            t["Outcome"] = outcome
            day_pnl += pnl
            taken += 1
            final_trades.append(t)

        if taken > 0:
            capital += day_pnl
            for t in final_trades[-taken:]:
                t["CapitalAfter"] = round(capital, 2)
            capital_curve.append(capital)

    return final_trades, capital, capital_curve

def run_backtest(tickers, days):
    print(f"Running 9/20 EMA Crossover on {len(tickers)} tickers...")
    all_raw = []
    for i, ticker in enumerate(tickers, 1):
        print(f"  [{i}/{len(tickers)}] {ticker}")
        df = fetch_intraday_history(ticker, days)
        if df is not None:
            raw = simulate_ticker(ticker, df)
            all_raw.extend(raw)
    return apply_position_sizing(all_raw)

# Excel + Web Server (identical to previous)
TRADE_COLUMNS = ["EntryDate", "EntryTime", "Ticker", "Direction", "Entry", "InitialStop", "Target",
                 "Qty", "RiskAmount", "ExitTime", "ExitPrice", "ExitReason", "Outcome", "PnL", "PnLPct", "CapitalAfter"]

def max_drawdown_pct(curve):
    peak = curve[0]
    max_dd = 0
    for v in curve:
        peak = max(peak, v)
        dd = (peak - v) / peak * 100
        max_dd = max(max_dd, dd)
    return round(max_dd, 2)

def write_excel(trades, start_cap, end_cap, curve):
    wb = Workbook()
    ws = wb.active
    ws.title = "Trades"
    ws.append(TRADE_COLUMNS)
    for t in trades:
        ws.append([t.get(c, "") for c in TRADE_COLUMNS])

    wins = sum(1 for t in trades if t.get("Outcome") == "WIN")
    total = len(trades)
    win_rate = round(wins / total * 100, 2) if total else 0

    summary = wb.create_sheet("Summary")
    for row in [
        ("Starting Capital", start_cap),
        ("Ending Capital", round(end_cap, 2)),
        ("Total Return %", round((end_cap - start_cap)/start_cap*100, 2)),
        ("Total Trades", total),
        ("Wins", wins),
        ("Losses", total - wins),
        ("Win Rate %", win_rate),
        ("Max Drawdown %", max_drawdown_pct(curve)),
    ]:
        summary.append(row)

    wb.save(OUTPUT_FILE)
    print(f"Results saved → {OUTPUT_FILE}")

# Web Server (unchanged)
backtest_status = {"running": False, "started_at": None, "finished_at": None, "result_summary": None, "error": None}

def _run_background(days, tickers_arg):
    backtest_status["running"] = True
    backtest_status["started_at"] = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
    try:
        tickers = [t.strip() for t in tickers_arg.split(",") if t.strip()] if tickers_arg else get_nifty_100_tickers()
        trades, end_cap, curve = run_backtest(tickers, min(days, MAX_LOOKBACK_DAYS))
        write_excel(trades, BACKTEST_CAPITAL, end_cap, curve)
        wins = sum(1 for t in trades if t.get("Outcome") == "WIN")
        backtest_status["result_summary"] = f"{len(trades)} trades | Win Rate ~{round(wins/len(trades)*100,1) if trades else 0}%"
    except Exception as e:
        backtest_status["error"] = str(e)
    finally:
        backtest_status["running"] = False
        backtest_status["finished_at"] = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path.startswith("/run-backtest"):
            self._trigger()
        elif self.path.startswith("/status"):
            self._status()
        elif self.path.startswith("/download"):
            self._download()
        else:
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"EMA Crossover Backtest Service\nUse /run-backtest?days=30&tickers=...\n")

    def _trigger(self):
        if backtest_status["running"]:
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"Already running.")
            return
        query = parse_qs(urlparse(self.path).query)
        days = int(query.get("days", [MAX_LOOKBACK_DAYS])[0])
        tickers_arg = query.get("tickers", [None])[0]
        threading.Thread(target=_run_background, args=(days, tickers_arg), daemon=True).start()
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Backtest started.")

    def _status(self):
        self.send_response(200)
        self.end_headers()
        msg = f"Running: {backtest_status['running']}\nResult: {backtest_status['result_summary']}"
        self.wfile.write(msg.encode())

    def _download(self):
        if os.path.exists(OUTPUT_FILE):
            with open(OUTPUT_FILE, "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", "attachment; filename=ema_crossover_journal.xlsx")
            self.end_headers()
            self.wfile.write(data)
        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b"No file yet.")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--tickers", type=str, default=None)
    parser.add_argument("--days", type=int, default=MAX_LOOKBACK_DAYS)
    parser.add_argument("--once", action="store_true")
    args = parser.parse_args()

    if args.once:
        tickers = [t.strip() for t in (args.tickers or "").split(",") if t.strip()] or get_nifty_100_tickers()
        trades, end_cap, curve = run_backtest(tickers, args.days)
        write_excel(trades, BACKTEST_CAPITAL, end_cap, curve)
        return

    port = int(os.environ.get("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), Handler)
    print(f"Server running on port {port}")
    server.serve_forever()

if __name__ == "__main__":
    main()
