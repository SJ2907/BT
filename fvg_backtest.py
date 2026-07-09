"""
Fair Value Gap (FVG) Strategy BACKTEST - FIXED VERSION
All 3 candles (c1, c2, c3) must be bullish for Bullish FVG
"""

import argparse
import os
import sys
import threading
import time
import uuid
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

DISPLACEMENT_BODY_MULT = 1.5
DISPLACEMENT_VOL_MULT = 2.0
MIN_GAP_PCT = 0.15

FVG_RETEST_MAX_CANDLES = 15
SL_BUFFER_PCT = 0.1
TARGET_RR = 2.5
MEASURED_MOVE_MULT = 1.0

TRAIL_TO_BE_R = 1.2
TRAIL_LOOKBACK_CANDLES = 4

NO_ENTRY_AFTER_HOUR, NO_ENTRY_AFTER_MINUTE = 14, 0
EXIT_HOUR, EXIT_MINUTE = 15, 0

IST = pytz.timezone("Asia/Kolkata")
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "fvg_journal.xlsx")
NIFTY100_SOURCE_URL = "https://niftyindices.com/IndexConstituent/ind_nifty100list.csv"

FALLBACK_TICKERS = [
    "RELIANCE.NS", "TCS.NS", "HDFCBANK.NS", "ICICIBANK.NS", "INFY.NS",
    "HINDUNILVR.NS", "ITC.NS", "SBIN.NS", "BHARTIARTL.NS", "KOTAKBANK.NS",
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

def fetch_intraday_history(ticker, days):
    try:
        df = yf.download(ticker, period=f"{days}d", interval=CANDLE_INTERVAL, progress=False, auto_adjust=False)
    except Exception as e:
        print(f"  [warn] fetch failed for {ticker}: {e}")
        return None
    if df is None or df.empty:
        return None
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    return df

def precompute_indicators(df):
    body = (df["Close"] - df["Open"])
    df["_avg_body_prior"] = body.abs().shift(1).rolling(10).mean()
    df["_avg_vol"] = df["Volume"].rolling(20).mean()
    return df

def simulate_ticker(ticker, df):
    df = precompute_indicators(df)
    raw_trades = []
    all_signals = []

    in_position = False
    entry_price = stop_price = trailing_stop = target = None
    entry_time = direction = None
    signal_id = None
    current_day = None
    pending = None

    n = len(df)
    for i in range(n):
        row = df.iloc[i]
        row_time = row.name.time()
        row_day = row.name.date()

        if row_day != current_day:
            current_day = row_day
            pending = None
            in_position = False

        if i < 25:
            continue

        if in_position:
            if row_time >= pd.Timestamp(f"{EXIT_HOUR}:{EXIT_MINUTE}").time():
                exit_price = float(row["Close"])
                raw_trades.append(_build_trade(signal_id, ticker, direction, entry_time, entry_price, stop_price, target, row.name, exit_price, "EOD_SQUAREOFF"))
                in_position = False
                continue

            price = float(row["Close"])
            risk = abs(entry_price - stop_price)
            if direction == "BULLISH":
                if price >= target:
                    raw_trades.append(_build_trade(signal_id, ticker, direction, entry_time, entry_price, stop_price, target, row.name, price, "TARGET_HIT"))
                    in_position = False
                    continue
                if (price - entry_price) / risk >= TRAIL_TO_BE_R:
                    trailing_stop = max(trailing_stop, entry_price)
                    lookback_low = float(df.iloc[max(0, i - TRAIL_LOOKBACK_CANDLES):i]["Low"].min())
                    trailing_stop = max(trailing_stop, lookback_low)
                if price <= trailing_stop:
                    raw_trades.append(_build_trade(signal_id, ticker, direction, entry_time, entry_price, stop_price, target, row.name, price, "TRAIL_STOP_HIT"))
                    in_position = False
                    continue
            else:
                if price <= target:
                    raw_trades.append(_build_trade(signal_id, ticker, direction, entry_time, entry_price, stop_price, target, row.name, price, "TARGET_HIT"))
                    in_position = False
                    continue
                if (entry_price - price) / risk >= TRAIL_TO_BE_R:
                    trailing_stop = min(trailing_stop, entry_price)
                    lookback_high = float(df.iloc[max(0, i - TRAIL_LOOKBACK_CANDLES):i]["High"].max())
                    trailing_stop = min(trailing_stop, lookback_high)
                if price >= trailing_stop:
                    raw_trades.append(_build_trade(signal_id, ticker, direction, entry_time, entry_price, stop_price, target, row.name, price, "TRAIL_STOP_HIT"))
                    in_position = False
                    continue
            continue

        if pending is not None:
            candles_waited = i - pending["detected_idx"]
            if candles_waited > FVG_RETEST_MAX_CANDLES:
                pending = None
            else:
                zone_top, zone_bottom = pending["zone_top"], pending["zone_bottom"]
                if pending["direction"] == "BULLISH":
                    if row["Close"] < zone_bottom * (1 - SL_BUFFER_PCT / 100):
                        pending = None
                    elif row["Low"] <= zone_top and row["Close"] > row["Open"] and row["Close"] > zone_bottom:
                        entry_price = float(row["Close"])
                        stop_price = zone_bottom * (1 - SL_BUFFER_PCT / 100)
                        risk = entry_price - stop_price
                        if risk > 0:
                            c2_range = pending["c2_range"]
                            measured_target = entry_price + c2_range * MEASURED_MOVE_MULT
                            rr_target = entry_price + TARGET_RR * risk
                            target = max(measured_target, rr_target)
                            trailing_stop = stop_price
                            direction = "BULLISH"
                            entry_time = row.name
                            signal_id = pending["signal_id"]
                            in_position = True
                            for s in all_signals:
                                if s["SignalId"] == signal_id:
                                    s["EntryTriggered"] = True
                                    s["EntryTime"] = entry_time.strftime("%H:%M:%S")
                        pending = None
                # Bearish logic can be added similarly if needed
            if in_position:
                continue

        # --- Gap Detection with 3 Bullish Candles ---
        if pending is None and i >= 2 and row_time <= pd.Timestamp(f"{NO_ENTRY_AFTER_HOUR}:{NO_ENTRY_AFTER_MINUTE}").time():
            c1 = df.iloc[i - 2]
            c2 = df.iloc[i - 1]
            c3 = row

            # ALL 3 CANDLES MUST BE BULLISH
            if not (c1["Close"] > c1["Open"] and c2["Close"] > c2["Open"] and c3["Close"] > c3["Open"]):
                continue

            avg_body = c2["_avg_body_prior"]
            avg_vol = c2["_avg_vol"]
            if pd.isna(avg_body) or pd.isna(avg_vol) or avg_body <= 0:
                continue

            c2_body = abs(c2["Close"] - c2["Open"])
            displaced = c2_body >= DISPLACEMENT_BODY_MULT * avg_body and c2["Volume"] >= DISPLACEMENT_VOL_MULT * avg_vol
            if not displaced:
                continue

            price_ref = float(c3["Close"])

            if c1["High"] < c3["Low"]:
                gap_pct = (c3["Low"] - c1["High"]) / price_ref * 100
                if gap_pct >= MIN_GAP_PCT:
                    is_fvg = c3["Close"] <= c2["High"]
                    sig_id = str(uuid.uuid4())[:8]
                    all_signals.append({
                        "SignalId": sig_id, "Date": row.name.strftime("%Y-%m-%d"),
                        "DetectedTime": row.name.strftime("%H:%M:%S"), "Ticker": ticker,
                        "GapType": "FVG", "Direction": "BULLISH",
                        "ZoneLow": round(float(c1["High"]), 2),
                        "ZoneHigh": round(float(c3["Low"]), 2),
                        "EntryTriggered": False, "EntryTime": ""
                    })
                    if is_fvg:
                        pending = {
                            "direction": "BULLISH",
                            "zone_top": float(c3["Low"]),
                            "zone_bottom": float(c1["High"]),
                            "c2_range": float(c2["High"] - c2["Low"]),
                            "detected_idx": i,
                            "signal_id": sig_id
                        }

    return raw_trades, all_signals

def _build_trade(signal_id, ticker, direction, entry_time, entry_price, stop_price, target, exit_time, exit_price, exit_reason):
    risk = abs(entry_price - stop_price)
    return {
        "SignalId": signal_id,
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
        "ExitReason": exit_reason,
    }

def apply_position_sizing(raw_trades):
    by_day = {}
    for t in raw_trades:
        by_day.setdefault(t["EntryDate"], []).append(t)

    capital = BACKTEST_CAPITAL
    final_trades = []
    capital_curve = [capital]

    for day in sorted(by_day.keys()):
        day_trades = sorted(by_day[day], key=lambda t: t["EntryTime"])
        capital_start_of_day = capital
        committed_risk = 0.0
        day_pnl = 0.0
        taken_count = 0

        for t in day_trades:
            risk_amount = capital_start_of_day * (RISK_PER_TRADE_PCT / 100)
            if committed_risk + risk_amount > capital_start_of_day * (MAX_DAILY_RISK_PCT / 100):
                continue
            committed_risk += risk_amount
            qty = max(1, int(risk_amount // t["RiskPerShare"])) if t["RiskPerShare"] > 0 else 0
            if qty == 0:
                continue

            pnl = (t["ExitPrice"] - t["Entry"]) * qty
            pnl_pct = round((pnl / (t["Entry"] * qty)) * 100, 2) if t["Entry"] * qty else 0
            outcome = "WIN" if pnl > 0 else "LOSS"

            t["Qty"] = qty
            t["RiskAmount"] = round(risk_amount, 2)
            t["PnL"] = round(pnl, 2)
            t["PnLPct"] = pnl_pct
            t["Outcome"] = outcome
            day_pnl += pnl
            taken_count += 1
            final_trades.append(t)

        if taken_count > 0:
            capital += day_pnl
            for t in final_trades[-taken_count:]:
                t["CapitalAfter"] = round(capital, 2)
            capital_curve.append(capital)

    return final_trades, capital, capital_curve

def run_backtest(tickers, days):
    print(f"Fetching {days}-day data for {len(tickers)} tickers...")
    all_raw_trades = []
    all_signals = []
    usable = 0
    for i, ticker in enumerate(tickers, 1):
        print(f"  [{i}/{len(tickers)}] {ticker}...")
        df = fetch_intraday_history(ticker, days)
        if df is None or len(df) < 100:
            continue
        usable += 1
        try:
            raw, signals = simulate_ticker(ticker, df)
            all_raw_trades.extend(raw)
            all_signals.extend(signals)
        except Exception as e:
            print(f"  [warn] error for {ticker}: {e}")

    print(f"Usable tickers: {usable}/{len(tickers)}. Gaps found: {len(all_signals)}. Applying sizing...")
    final_trades, ending_capital, curve = apply_position_sizing(all_raw_trades)
    return final_trades, ending_capital, curve, all_signals

# ====================== EXCEL OUTPUT ======================

TRADE_COLUMNS = ["EntryDate", "EntryTime", "Ticker", "Direction", "Entry", "InitialStop",
                  "Target", "Qty", "RiskAmount", "ExitTime", "ExitPrice", "ExitReason",
                  "Outcome", "PnL", "PnLPct", "CapitalAfter"]
SIGNAL_COLUMNS = ["Date", "DetectedTime", "Ticker", "GapType", "Direction", "ZoneLow",
                   "ZoneHigh", "EntryTriggered", "EntryTime"]

def max_drawdown_pct(curve):
    peak = curve[0]
    max_dd = 0.0
    for value in curve:
        peak = max(peak, value)
        dd = (peak - value) / peak * 100 if peak else 0
        max_dd = max(max_dd, dd)
    return round(max_dd, 2)

def write_excel(trades, starting_capital, ending_capital, capital_curve, all_signals):
    wb = Workbook()
    ws = wb.active
    ws.title = "Trades"
    ws.append(TRADE_COLUMNS)
    for t in trades:
        ws.append([t.get(c, "") for c in TRADE_COLUMNS])

    wins = sum(1 for t in trades if t["Outcome"] == "WIN")
    total = len(trades)
    win_rate = round(wins / total * 100, 2) if total else 0
    total_pnl = round(sum(t["PnL"] for t in trades), 2)
    return_pct = round((ending_capital - starting_capital) / starting_capital * 100, 2)

    summary_ws = wb.create_sheet("Summary")
    for row in [
        ("Starting Capital", starting_capital),
        ("Ending Capital", round(ending_capital, 2)),
        ("Total Return %", return_pct),
        ("Total Trades", total),
        ("Wins", wins),
        ("Losses", total - wins),
        ("Win Rate %", win_rate),
        ("Total P&L", total_pnl),
        ("Max Drawdown %", max_drawdown_pct(capital_curve)),
        ("Risk Per Trade %", RISK_PER_TRADE_PCT),
        ("Max Daily Risk Cap %", MAX_DAILY_RISK_PCT),
    ]:
        summary_ws.append(row)

    signals_ws = wb.create_sheet("AllSignals")
    signals_ws.append(SIGNAL_COLUMNS)
    for s in sorted(all_signals, key=lambda x: (x["Date"], x["DetectedTime"])):
        signals_ws.append([s.get(c, "") for c in SIGNAL_COLUMNS])

    wb.save(OUTPUT_FILE)
    print(f"Results saved to {OUTPUT_FILE}")

# ====================== WEB SERVER ======================

backtest_status = {"running": False, "started_at": None, "finished_at": None, "result_summary": None, "error": None}

def _run_backtest_background(days, tickers_arg):
    backtest_status["running"] = True
    backtest_status["started_at"] = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
    try:
        tickers = [t.strip() for t in tickers_arg.split(",") if t.strip()] if tickers_arg else get_nifty_100_tickers()
        days = min(days, MAX_LOOKBACK_DAYS)
        trades, ending_capital, curve, all_signals = run_backtest(tickers, days)
        write_excel(trades, BACKTEST_CAPITAL, ending_capital, curve, all_signals)
        wins = sum(1 for t in trades if t["Outcome"] == "WIN")
        backtest_status["result_summary"] = f"{len(trades)} trades, win rate {round(wins/len(trades)*100,1) if trades else 0}%, ended at Rs.{round(ending_capital,2)}"
    except Exception as e:
        backtest_status["error"] = str(e)
    finally:
        backtest_status["running"] = False
        backtest_status["finished_at"] = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")

class _Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path.startswith("/run-backtest"):
            self._trigger()
        elif self.path.startswith("/status"):
            self._status()
        elif self.path.startswith("/download"):
            self._download()
        else:
            self._home()

    def _home(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        self.wfile.write(b"FVG Backtest Service\n/run-backtest?days=30\n/status\n/download\n")

    def _trigger(self):
        if backtest_status["running"]:
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"Already running.")
            return
        query = parse_qs(urlparse(self.path).query)
        days = int(query.get("days", [MAX_LOOKBACK_DAYS])[0])
        tickers_arg = query.get("tickers", [None])[0]
        threading.Thread(target=_run_backtest_background, args=(days, tickers_arg), daemon=True).start()
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Backtest started.")

    def _status(self):
        self.send_response(200)
        self.end_headers()
        msg = f"Running: {backtest_status['running']}\nResult: {backtest_status['result_summary']}\nError: {backtest_status['error']}"
        self.wfile.write(msg.encode())

    def _download(self):
        if os.path.exists(OUTPUT_FILE):
            with open(OUTPUT_FILE, "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", "attachment; filename=fvg_journal.xlsx")
            self.end_headers()
            self.wfile.write(data)
        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b"No file yet.")

    def log_message(self, format, *args):
        pass

def start_web_server():
    port = int(os.environ.get("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), _Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    print(f"FVG backtest server listening on port {port}.")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--tickers", type=str, default=None)
    parser.add_argument("--days", type=int, default=MAX_LOOKBACK_DAYS)
    parser.add_argument("--once", action="store_true")
    args = parser.parse_args()

    if args.once:
        days = min(args.days, MAX_LOOKBACK_DAYS)
        tickers = [t.strip() for t in args.tickers.split(",") if t.strip()] if args.tickers else get_nifty_100_tickers()
        trades, ending_capital, curve, all_signals = run_backtest(tickers, days)
        write_excel(trades, BACKTEST_CAPITAL, ending_capital, curve, all_signals)
        print("Backtest completed. Check fvg_journal.xlsx")
        return

    start_web_server()
    print("Waiting for /run-backtest...")

if __name__ == "__main__":
    main()
