"""
L99 Strategy BACKTEST -- RSI Reversal + MACD Cross + EMA/VWAP Trend
=====================================================================
(Indian Stock Market - NSE)

Fully SEPARATE from the live bot and every other backtest script here.
Doesn't touch or import any of them.

Adapted from a reference script, rebuilt on our validated backtest
framework -- the reference had a real bug (stop-loss/target were never
carried forward past the entry candle, so `row['low'] <= previous_row's
stop_loss` was comparing against NaN on almost every candle, meaning
stops/targets essentially never triggered as written) and, like the other
reference scripts we've tested, used a disconnected approximate equity
calculation that didn't match its own entry/exit rules. This version
tracks the actual stop/target for each open trade properly, and simulates
real per-trade P&L with position sizing.

STRATEGY LOGIC (long only, matching the reference)
--------------------------------------------------
Entry requires ALL of:
  1. RSI(14) below 25 for the last 7 CONSECUTIVE candles (deep, sustained
     oversold -- not just a single dip)
  2. A fresh MACD bullish crossover on the current candle (MACD crosses
     above its signal line THIS candle, not just already above)
  3. Volume RSI(14) between 40 and 55 (volume momentum in a neutral zone
     -- not a volume spike, not dried up)
  4. Close above EMA5, EMA20, AND EMA44 (short/medium-term trend alignment)
  5. Close above session VWAP
  6. Before 12:30 PM (no new entries after this)

Risk management: 1% fixed stop below entry, target at 1:2 risk:reward
(2% above entry). Mandatory flat-by 2:30 PM regardless of stop/target.

IMPORTANT LIMITATIONS
--------------------------------------------------------------------------
- Same ~60-day cap on 5-minute data as the other backtests (Yahoo's limit).
- No brokerage/STT/slippage modeled.
- Today's Nifty 100 list used for the whole window (minor survivorship bias).
- The stacked entry conditions (deep sustained oversold + trading above
  all three EMAs + above VWAP, simultaneously) are somewhat in tension
  with each other -- expect this to fire rarely. That's a property of the
  conditions as specified, not a bug in this implementation.
- MAX_DAILY_RISK_PCT caps total risk committed across all simultaneous
  signals on one day.
- Exits check High/Low for actual level touches (not just candle Close).
- Research simulation only, not a guarantee of live results.

REQUIREMENTS
------------
    pip install yfinance pandas pytz openpyxl requests

HOW TO RUN
----------
    python3 l99_backtest.py --once                      # run now, in this terminal, then exit
    python3 l99_backtest.py --once --tickers RELIANCE.NS,TCS.NS
    python3 l99_backtest.py                             # web-service mode (for Render):
                                                         #   binds a port, waits for a
                                                         #   trigger at /run-backtest
"""

import argparse
import os
import sys
import threading
import time
import uuid
from datetime import datetime, timedelta
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

RSI_PERIOD = 14
RSI_OVERSOLD_THRESHOLD = 25
RSI_OVERSOLD_CONSECUTIVE = 7   # candles in a row below the threshold

MACD_FAST, MACD_SLOW, MACD_SIGNAL = 12, 26, 9

VOL_RSI_MIN, VOL_RSI_MAX = 40, 55

EMA_FAST, EMA_MED, EMA_SLOW = 5, 20, 44

STOP_PCT = 1.0     # fixed 1% stop below entry
TARGET_RR = 2.0    # 1:2 risk:reward

NO_ENTRY_AFTER_HOUR, NO_ENTRY_AFTER_MINUTE = 12, 30
EXIT_HOUR, EXIT_MINUTE = 14, 30

IST = pytz.timezone("Asia/Kolkata")
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "l99_journal.xlsx")
NIFTY100_SOURCE_URL = "https://niftyindices.com/IndexConstituent/ind_nifty100list.csv"

FALLBACK_TICKERS = [
    "RELIANCE.NS", "TCS.NS", "HDFCBANK.NS", "ICICIBANK.NS", "INFY.NS",
    "HINDUNILVR.NS", "ITC.NS", "SBIN.NS", "BHARTIARTL.NS", "KOTAKBANK.NS",
]

# ============================ END CONFIG ==================================


def get_nifty_100_tickers():
    try:
        resp = requests.get(NIFTY100_SOURCE_URL,
                             headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
        resp.raise_for_status()
        df = pd.read_csv(StringIO(resp.text))
        symbol_col = "Symbol" if "Symbol" in df.columns else df.columns[2]
        tickers = [f"{s.strip()}.NS" for s in df[symbol_col].dropna().tolist()]
        if len(tickers) < 90:
            raise ValueError(f"Only got {len(tickers)} tickers, expected ~100")
        print(f"Fetched live Nifty 100 list: {len(tickers)} tickers.")
        return tickers
    except Exception as e:
        print(f"[warn] Live Nifty 100 fetch failed ({e}). Using a small fallback list.")
        return FALLBACK_TICKERS


def fetch_intraday_history(ticker, days):
    try:
        df = yf.download(ticker, period=f"{days}d", interval=CANDLE_INTERVAL,
                          progress=False, auto_adjust=False)
    except Exception as e:
        print(f"  [warn] fetch failed for {ticker}: {e}")
        return None
    if df is None or df.empty:
        return None
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    return df


# ------------------------------ INDICATORS ---------------------------------

def compute_rsi(series, period):
    delta = series.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.ewm(alpha=1 / period, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1 / period, adjust=False).mean()
    rs = avg_gain / avg_loss.replace(0, pd.NA)
    result = 100 - (100 / (1 + rs))
    return result.fillna(50)


def precompute_indicators(df):
    """All computed on the CONTINUOUS multi-day series (causal, no
    look-ahead) EXCEPT VWAP, which resets every session -- computed
    separately per day and reassembled."""
    df["RSI"] = compute_rsi(df["Close"], RSI_PERIOD)
    df["VolRSI"] = compute_rsi(df["Volume"], RSI_PERIOD)

    ema_fast = df["Close"].ewm(span=MACD_FAST, adjust=False).mean()
    ema_slow = df["Close"].ewm(span=MACD_SLOW, adjust=False).mean()
    df["MACD"] = ema_fast - ema_slow
    df["MACD_SIGNAL"] = df["MACD"].ewm(span=MACD_SIGNAL, adjust=False).mean()

    df["EMA_FAST"] = df["Close"].ewm(span=EMA_FAST, adjust=False).mean()
    df["EMA_MED"] = df["Close"].ewm(span=EMA_MED, adjust=False).mean()
    df["EMA_SLOW"] = df["Close"].ewm(span=EMA_SLOW, adjust=False).mean()

    # Session VWAP: reset at the start of each trading day
    typical_price = (df["High"] + df["Low"] + df["Close"]) / 3
    day_groups = df.index.date
    tpv = typical_price * df["Volume"]
    df["VWAP"] = tpv.groupby(day_groups).cumsum() / df["Volume"].groupby(day_groups).cumsum()

    return df


def _time_at_or_after(row_time, hour, minute):
    return row_time >= pd.Timestamp(f"{hour}:{minute}").time()


# ------------------------------ STRATEGY CORE ------------------------------

def simulate_ticker(ticker, df):
    """Runs the full multi-day series for one ticker. Returns
    (raw_trades, all_signals)."""
    df = precompute_indicators(df)
    raw_trades = []
    all_signals = []

    in_position = False
    entry_price = stop_price = target = entry_time = None
    signal_id = None
    current_day = None

    n = len(df)
    for i in range(n):
        row = df.iloc[i]
        row_time = row.name.time()
        row_day = row.name.date()

        if row_day != current_day:
            current_day = row_day
            in_position = False

        if i < RSI_OVERSOLD_CONSECUTIVE + 5:
            continue  # not enough history yet for the 7-candle RSI check

        if in_position:
            if _time_at_or_after(row_time, EXIT_HOUR, EXIT_MINUTE):
                exit_price = float(row["Close"])
                raw_trades.append(_build_trade(signal_id, ticker, entry_time, entry_price,
                                                stop_price, target, row.name, exit_price,
                                                "EOD_SQUAREOFF"))
                in_position = False
                continue

            # Check High/Low touches, not just Close
            if row["Open"] <= stop_price:
                exit_price, reason = float(row["Open"]), "STOP_HIT"
            elif row["Low"] <= stop_price:
                exit_price, reason = stop_price, "STOP_HIT"
            elif row["Open"] >= target:
                exit_price, reason = float(row["Open"]), "TARGET_HIT"
            elif row["High"] >= target:
                exit_price, reason = target, "TARGET_HIT"
            else:
                exit_price, reason = None, None

            if exit_price is not None:
                raw_trades.append(_build_trade(signal_id, ticker, entry_time, entry_price,
                                                stop_price, target, row.name, exit_price, reason))
                in_position = False
            continue

        # --- Not in a position: check entry conditions ---
        if _time_at_or_after(row_time, NO_ENTRY_AFTER_HOUR, NO_ENTRY_AFTER_MINUTE):
            continue

        rsi_window = df["RSI"].iloc[i - RSI_OVERSOLD_CONSECUTIVE:i]
        if pd.isna(rsi_window).any() or len(rsi_window) < RSI_OVERSOLD_CONSECUTIVE:
            continue
        rsi_oversold = (rsi_window < RSI_OVERSOLD_THRESHOLD).all()
        if not rsi_oversold:
            continue

        prev = df.iloc[i - 1]
        macd_cross = (row["MACD"] > row["MACD_SIGNAL"]) and (prev["MACD"] <= prev["MACD_SIGNAL"])
        if not macd_cross:
            continue

        if pd.isna(row["VolRSI"]) or not (VOL_RSI_MIN <= row["VolRSI"] <= VOL_RSI_MAX):
            continue

        ema_aligned = row["Close"] > row["EMA_FAST"] and row["Close"] > row["EMA_MED"] and row["Close"] > row["EMA_SLOW"]
        if not ema_aligned:
            continue

        if not (row["Close"] > row["VWAP"]):
            continue

        # All conditions met -- signal!
        sig_id = str(uuid.uuid4())[:8]
        all_signals.append({
            "SignalId": sig_id, "Date": row.name.strftime("%Y-%m-%d"),
            "DetectedTime": row.name.strftime("%H:%M:%S"), "Ticker": ticker,
            "RSI": round(float(row["RSI"]), 1), "VolRSI": round(float(row["VolRSI"]), 1),
            "EntryTriggered": True, "EntryTime": row.name.strftime("%H:%M:%S"),
        })

        entry_price = float(row["Close"])
        stop_price = entry_price * (1 - STOP_PCT / 100)
        risk = entry_price - stop_price
        target = entry_price + TARGET_RR * risk
        entry_time = row.name
        signal_id = sig_id
        in_position = True

    return raw_trades, all_signals


def _build_trade(signal_id, ticker, entry_time, entry_price, stop_price, target,
                  exit_time, exit_price, exit_reason):
    risk = entry_price - stop_price
    return {
        "SignalId": signal_id,
        "EntryDate": entry_time.strftime("%Y-%m-%d"), "EntryTime": entry_time.strftime("%H:%M:%S"),
        "Ticker": ticker,
        "Entry": round(entry_price, 2), "InitialStop": round(stop_price, 2),
        "Target": round(target, 2), "RiskPerShare": round(risk, 2),
        "ExitTime": exit_time.strftime("%H:%M:%S"), "ExitPrice": round(exit_price, 2),
        "ExitReason": exit_reason,
    }


# ------------------------------ POSITION SIZING -----------------------------

def apply_position_sizing(raw_trades):
    by_day = {}
    for t in raw_trades:
        by_day.setdefault(t["EntryDate"], []).append(t)

    capital = BACKTEST_CAPITAL
    final_trades = []
    capital_curve = [capital]
    taken_signal_ids = set()

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
            taken_signal_ids.add(t["SignalId"])
            final_trades.append(t)

        if taken_count > 0:
            capital += day_pnl
            for t in final_trades[-taken_count:]:
                t["CapitalAfter"] = round(capital, 2)
            capital_curve.append(capital)
            skipped = len(day_trades) - taken_count
            skip_note = f" ({skipped} skipped by daily risk cap)" if skipped else ""
            print(f"  {day}: {taken_count} trade(s) taken{skip_note}, "
                  f"day P&L Rs.{round(day_pnl,2)}, capital now Rs.{round(capital,2)}")

    return final_trades, capital, capital_curve, taken_signal_ids


def run_backtest(tickers, days):
    print(f"Fetching {days}-day intraday history for {len(tickers)} tickers...")
    all_raw_trades, all_signals = [], []
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
            print(f"  [warn] simulation error for {ticker}: {e}")

    print(f"\nGot usable data for {usable}/{len(tickers)} tickers. "
          f"{len(all_signals)} signals detected. Applying position sizing...")

    final_trades, ending_capital, curve, taken_ids = apply_position_sizing(all_raw_trades)
    for s in all_signals:
        s["TakenAsTrade"] = s["SignalId"] in taken_ids
    return final_trades, ending_capital, curve, all_signals


# ------------------------------ EXCEL OUTPUT --------------------------------

TRADE_COLUMNS = ["EntryDate", "EntryTime", "Ticker", "Entry", "InitialStop", "Target",
                  "Qty", "RiskAmount", "ExitTime", "ExitPrice", "ExitReason", "Outcome",
                  "PnL", "PnLPct", "CapitalAfter"]
SIGNAL_COLUMNS = ["Date", "DetectedTime", "Ticker", "RSI", "VolRSI", "EntryTriggered",
                   "EntryTime", "TakenAsTrade"]


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
        ("Starting Capital", starting_capital), ("Ending Capital", round(ending_capital, 2)),
        ("Total Return %", return_pct), ("Total Trades", total), ("Wins", wins),
        ("Losses", total - wins), ("Win Rate %", win_rate), ("Total P&L", total_pnl),
        ("Max Drawdown %", max_drawdown_pct(capital_curve)),
        ("Risk Per Trade %", RISK_PER_TRADE_PCT), ("Max Daily Risk Cap %", MAX_DAILY_RISK_PCT),
        ("Total Signals Detected", len(all_signals)),
        ("No new entries after", f"{NO_ENTRY_AFTER_HOUR}:{NO_ENTRY_AFTER_MINUTE:02d}"),
        ("Mandatory exit by", f"{EXIT_HOUR}:{EXIT_MINUTE:02d}"),
        ("Fixed stop %", STOP_PCT), ("Target R:R", TARGET_RR),
    ]:
        summary_ws.append(row)

    by_symbol = {}
    for t in trades:
        s = by_symbol.setdefault(t["Ticker"], {"trades": 0, "wins": 0, "pnl": 0.0})
        s["trades"] += 1
        s["wins"] += 1 if t["Outcome"] == "WIN" else 0
        s["pnl"] += t["PnL"]
    symbol_ws = wb.create_sheet("BySymbol")
    symbol_ws.append(["Ticker", "Trades", "Wins", "WinRate%", "TotalPnL"])
    for ticker, s in sorted(by_symbol.items(), key=lambda x: -x[1]["pnl"]):
        wr = round(s["wins"] / s["trades"] * 100, 2) if s["trades"] else 0
        symbol_ws.append([ticker, s["trades"], s["wins"], wr, round(s["pnl"], 2)])

    signals_ws = wb.create_sheet("AllSignals")
    signals_ws.append(SIGNAL_COLUMNS)
    for s in sorted(all_signals, key=lambda x: (x["Date"], x["DetectedTime"])):
        signals_ws.append([s.get(c, "") for c in SIGNAL_COLUMNS])

    wb.save(OUTPUT_FILE)


# --------------------------- STANDALONE WEB SERVER --------------------------

backtest_status = {"running": False, "started_at": None, "finished_at": None,
                    "result_summary": None, "error": None}


def _run_backtest_background(days, tickers_arg):
    backtest_status["running"] = True
    backtest_status["started_at"] = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
    backtest_status["error"] = None
    try:
        tickers = ([t.strip() for t in tickers_arg.split(",") if t.strip()]
                   if tickers_arg else get_nifty_100_tickers())
        days = min(days, MAX_LOOKBACK_DAYS)
        trades, ending_capital, curve, all_signals = run_backtest(tickers, days)
        if trades or all_signals:
            write_excel(trades, BACKTEST_CAPITAL, ending_capital, curve, all_signals)
            wins = sum(1 for t in trades if t["Outcome"] == "WIN")
            wr = round(wins / len(trades) * 100, 1) if trades else 0
            backtest_status["result_summary"] = (
                f"{len(trades)} trades, win rate {wr}%, ended at Rs.{round(ending_capital,2)} "
                f"(started Rs.{BACKTEST_CAPITAL:,.2f}), max drawdown {max_drawdown_pct(curve)}%, "
                f"{len(all_signals)} total signals detected"
            )
        else:
            backtest_status["result_summary"] = "No signals or trades were generated -- check server logs."
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

    def do_HEAD(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()

    def _home(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        self.wfile.write((
            "L99 backtest service alive.\n\n"
            "Visit /run-backtest to start (optional: ?days=30&tickers=RELIANCE.NS,TCS.NS)\n"
            "Visit /status to check progress.\n"
            "Visit /download to get l99_journal.xlsx once finished.\n"
        ).encode("utf-8"))

    def _trigger(self):
        if backtest_status["running"]:
            self.send_response(200)
            self.send_header("Content-Type", "text/plain")
            self.end_headers()
            self.wfile.write(b"A backtest is already running. Check /status.\n")
            return
        query = parse_qs(urlparse(self.path).query)
        days = int(query.get("days", [MAX_LOOKBACK_DAYS])[0])
        tickers_arg = query.get("tickers", [None])[0]
        threading.Thread(target=_run_backtest_background, args=(days, tickers_arg), daemon=True).start()
        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        self.wfile.write(
            f"Backtest started ({days} days). Check /status for progress, /download when done.\n".encode("utf-8")
        )

    def _status(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        lines = [f"Running: {backtest_status['running']}",
                 f"Started at: {backtest_status['started_at']}",
                 f"Finished at: {backtest_status['finished_at']}",
                 f"Result: {backtest_status['result_summary']}",
                 f"Error (if any): {backtest_status['error']}"]
        self.wfile.write(("\n".join(lines) + "\n").encode("utf-8"))

    def _download(self):
        if not os.path.exists(OUTPUT_FILE):
            self.send_response(404)
            self.send_header("Content-Type", "text/plain")
            self.end_headers()
            self.wfile.write(b"No l99_journal.xlsx yet -- visit /run-backtest first.\n")
            return
        with open(OUTPUT_FILE, "rb") as f:
            data = f.read()
        self.send_response(200)
        self.send_header("Content-Type",
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", "attachment; filename=l99_journal.xlsx")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, format, *args):
        pass


def start_web_server():
    port = int(os.environ.get("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), _Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    print(f"L99 backtest server listening on port {port}.")


def main():
    parser = argparse.ArgumentParser(description="Backtest the L99 RSI/MACD/EMA/VWAP strategy")
    parser.add_argument("--tickers", type=str, default=None)
    parser.add_argument("--days", type=int, default=MAX_LOOKBACK_DAYS)
    parser.add_argument("--once", action="store_true",
                         help="Run once immediately and exit (Shell tab / local use). "
                              "Without this flag, starts a web server for Render deployment.")
    args = parser.parse_args()

    if args.once:
        days = min(args.days, MAX_LOOKBACK_DAYS)
        tickers = ([t.strip() for t in args.tickers.split(",") if t.strip()]
                   if args.tickers else get_nifty_100_tickers())
        trades, ending_capital, curve, all_signals = run_backtest(tickers, days)
        if not trades and not all_signals:
            print("\nNo signals or trades were generated -- check the [warn] lines above.")
            return
        write_excel(trades, BACKTEST_CAPITAL, ending_capital, curve, all_signals)
        wins = sum(1 for t in trades if t["Outcome"] == "WIN")
        wr = round(wins / len(trades) * 100, 1) if trades else 0
        print(f"\n{'='*60}\nBACKTEST COMPLETE")
        print(f"  Total signals:      {len(all_signals)}")
        print(f"  Total trades:       {len(trades)}")
        print(f"  Win rate:           {wr}%")
        print(f"  Starting capital:   Rs.{BACKTEST_CAPITAL:,.2f}")
        print(f"  Ending capital:     Rs.{ending_capital:,.2f}")
        print(f"  Max drawdown:       {max_drawdown_pct(curve)}%")
        print(f"  Full detail in:     {OUTPUT_FILE}\n{'='*60}")
        return

    start_web_server()
    print("Waiting for a backtest to be triggered via /run-backtest. Press Ctrl+C to stop.")
    try:
        while True:
            time.sleep(60)
    except KeyboardInterrupt:
        print("\nStopped by user. Goodbye!")
        sys.exit(0)


if __name__ == "__main__":
    main()
