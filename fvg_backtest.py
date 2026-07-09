"""
Fair Value Gap (FVG) Strategy BACKTEST (Indian Stock Market - NSE)
=====================================================================
Fully SEPARATE from the live bot, the ORB backtest (BT), and the
breakout-pullback backtest. Doesn't touch or import any of them.

STRATEGY LOGIC
--------------
1. GAP DETECTION (3 consecutive candles: c1, c2, c3):
     Bullish gap: c1.High < c3.Low  (a void/imbalance left behind by c2)
     Bearish gap: c1.Low  > c3.High
   c2 (the middle, "displacement" candle) must show real force -- an
   unusually large body AND unusually high volume vs recent averages --
   to filter out tiny, meaningless gaps.

2. CLASSIFICATION (this is what separates FVG from a Breakaway Gap):
     FVG: c3 closes WITHIN or touching c2's own range/body -- price didn't
          run away, it's still "filling orders" near the impulsive candle.
     (A close that instead runs well beyond c2's high/low is a Breakaway
     Gap, not an FVG -- see bag_backtest.py for that side of the same
     pattern.)

3. FVG ENTRY -- do NOT chase: wait for price to retrace back INTO the gap
   zone (between c1's extreme and c3's extreme) and show a reaction candle
   in the original direction (closes bullish, holding above the zone's
   lower edge, for a bullish FVG). That reaction candle's close is the
   entry. If price hasn't reacted within FVG_RETEST_MAX_CANDLES, or closes
   decisively through the zone instead, the setup is abandoned.

4. RISK MANAGEMENT: stop just beyond the far edge of the gap zone; target
   = the larger of (c2's own range projected forward) or a fixed R-multiple.
   Trailing stop only activates after real profit (past TRAIL_TO_BE_R), and
   then trails behind a short rolling window -- NOT a single-candle trail,
   which was proven (in the breakout-pullback backtest) to choke trades
   almost immediately on ordinary noise.

IMPORTANT LIMITATIONS
--------------------------------------------------------------------------
- Same ~60-day cap on 5-minute data as the other backtests (Yahoo's limit).
- No brokerage/STT/slippage modeled.
- Today's Nifty 100 list used for the whole window (minor survivorship bias).
- MAX_DAILY_RISK_PCT caps total risk committed across all simultaneous
  signals on one day, so a cluster of same-day gaps can't silently stack
  into an outsized single-day loss.
- This "FVG works" idea is popular in retail/ICT trading circles but has
  no strong independent evidence behind it -- that's exactly why this
  script exists: to test it on real data rather than assume it.
- Research simulation only, not a guarantee of live results.

REQUIREMENTS
------------
    pip install yfinance pandas pytz openpyxl requests

HOW TO RUN
----------
    python3 fvg_backtest.py --once                      # run now, in this terminal, then exit
    python3 fvg_backtest.py --once --tickers RELIANCE.NS,TCS.NS
    python3 fvg_backtest.py                             # web-service mode (for Render):
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

# --- Gap detection ---
DISPLACEMENT_BODY_MULT = 1.5   # c2's body must be >= this x the recent average body
DISPLACEMENT_VOL_MULT = 2.0    # c2's volume must be >= this x the 20-candle average
MIN_GAP_PCT = 0.15             # gap size must be at least this % of price, to skip noise

# --- FVG entry (wait for retracement + reaction) ---
FVG_RETEST_MAX_CANDLES = 15    # give up waiting for a reaction after this many candles
SL_BUFFER_PCT = 0.1
TARGET_RR = 2.5
MEASURED_MOVE_MULT = 1.0       # project c2's own range forward by this multiple

# --- Trailing stop (fixed design -- see breakout_pullback_backtest.py notes) ---
TRAIL_TO_BE_R = 1.2
TRAIL_LOOKBACK_CANDLES = 4

NO_ENTRY_AFTER_HOUR, NO_ENTRY_AFTER_MINUTE = 14, 0    # no NEW entries after 2:00 PM
EXIT_HOUR, EXIT_MINUTE = 15, 0                         # mandatory flat-by time

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


def precompute_indicators(df):
    body = (df["Close"] - df["Open"])
    df["_avg_body_prior"] = body.abs().shift(1).rolling(10).mean()
    df["_avg_vol"] = df["Volume"].rolling(20).mean()
    return df


def _time_at_or_after(row_time, hour, minute):
    return row_time >= pd.Timestamp(f"{hour}:{minute}").time()


def _within_time(row_time, hour, minute):
    return row_time <= pd.Timestamp(f"{hour}:{minute}").time()


# ------------------------------ STRATEGY CORE ------------------------------

def simulate_ticker(ticker, df):
    """Returns (raw_trades, all_signals) -- raw_trades only contains signals
    that actually got an entry triggered (for position sizing later);
    all_signals logs EVERY gap detected that day, whether or not price ever
    retraced into it and triggered an entry."""
    df = precompute_indicators(df)
    raw_trades = []
    all_signals = []

    in_position = False
    entry_price = stop_price = trailing_stop = target = None
    entry_time = direction = None
    signal_id = None
    current_day = None

    # Active (unfilled) FVG setup awaiting a retracement + reaction
    pending = None  # dict or None

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

        # --- Manage an open position first ---
        if in_position:
            if _time_at_or_after(row_time, EXIT_HOUR, EXIT_MINUTE):
                exit_price = float(row["Close"])
                raw_trades.append(_build_trade(signal_id, ticker, direction, entry_time,
                                                entry_price, stop_price, target, row.name,
                                                exit_price, "EOD_SQUAREOFF"))
                in_position = False
                continue

            risk = abs(entry_price - stop_price)
            # FIX: check High/Low for whether price actually TOUCHED the
            # stop/target, exiting at that level (or the Open if it gapped
            # through) -- not just at whatever the candle's Close happened
            # to be. See breakout_pullback_backtest.py's notes for the full
            # explanation of why this matters (it was overstating losses).
            if direction == "BULLISH":
                if row["Open"] <= trailing_stop:
                    exit_price = float(row["Open"])
                elif row["Low"] <= trailing_stop:
                    exit_price = trailing_stop
                else:
                    exit_price = None
                if exit_price is not None:
                    raw_trades.append(_build_trade(signal_id, ticker, direction, entry_time,
                                                    entry_price, stop_price, target, row.name,
                                                    exit_price, "STOP_HIT"))
                    in_position = False
                    continue

                if row["Open"] >= target:
                    exit_price = float(row["Open"])
                elif row["High"] >= target:
                    exit_price = target
                else:
                    exit_price = None
                if exit_price is not None:
                    raw_trades.append(_build_trade(signal_id, ticker, direction, entry_time,
                                                    entry_price, stop_price, target, row.name,
                                                    exit_price, "TARGET_HIT"))
                    in_position = False
                    continue

                if (float(row["High"]) - entry_price) / risk >= TRAIL_TO_BE_R:
                    trailing_stop = max(trailing_stop, entry_price)
                    lookback_low = float(df.iloc[max(0, i - TRAIL_LOOKBACK_CANDLES):i]["Low"].min())
                    trailing_stop = max(trailing_stop, lookback_low)
            else:
                if row["Open"] >= trailing_stop:
                    exit_price = float(row["Open"])
                elif row["High"] >= trailing_stop:
                    exit_price = trailing_stop
                else:
                    exit_price = None
                if exit_price is not None:
                    raw_trades.append(_build_trade(signal_id, ticker, direction, entry_time,
                                                    entry_price, stop_price, target, row.name,
                                                    exit_price, "STOP_HIT"))
                    in_position = False
                    continue

                if row["Open"] <= target:
                    exit_price = float(row["Open"])
                elif row["Low"] <= target:
                    exit_price = target
                else:
                    exit_price = None
                if exit_price is not None:
                    raw_trades.append(_build_trade(signal_id, ticker, direction, entry_time,
                                                    entry_price, stop_price, target, row.name,
                                                    exit_price, "TARGET_HIT"))
                    in_position = False
                    continue

                if (entry_price - float(row["Low"])) / risk >= TRAIL_TO_BE_R:
                    trailing_stop = min(trailing_stop, entry_price)
                    lookback_high = float(df.iloc[max(0, i - TRAIL_LOOKBACK_CANDLES):i]["High"].max())
                    trailing_stop = min(trailing_stop, lookback_high)
            continue

        # --- Not in a position: monitor a pending FVG, or look for a new one ---
        if pending is not None:
            candles_waited = i - pending["detected_idx"]
            if candles_waited > FVG_RETEST_MAX_CANDLES:
                pending = None
            else:
                zone_top, zone_bottom = pending["zone_top"], pending["zone_bottom"]
                if pending["direction"] == "BULLISH":
                    # Invalidated if price closes decisively through the zone
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
                else:  # BEARISH pending
                    if row["Close"] > zone_top * (1 + SL_BUFFER_PCT / 100):
                        pending = None
                    elif row["High"] >= zone_bottom and row["Close"] < row["Open"] and row["Close"] < zone_top:
                        entry_price = float(row["Close"])
                        stop_price = zone_top * (1 + SL_BUFFER_PCT / 100)
                        risk = stop_price - entry_price
                        if risk > 0:
                            c2_range = pending["c2_range"]
                            measured_target = entry_price - c2_range * MEASURED_MOVE_MULT
                            rr_target = entry_price - TARGET_RR * risk
                            target = min(measured_target, rr_target)
                            trailing_stop = stop_price
                            direction = "BEARISH"
                            entry_time = row.name
                            signal_id = pending["signal_id"]
                            in_position = True
                            for s in all_signals:
                                if s["SignalId"] == signal_id:
                                    s["EntryTriggered"] = True
                                    s["EntryTime"] = entry_time.strftime("%H:%M:%S")
                        pending = None
            if in_position:
                continue

        # --- Look for a brand-new gap (only if not already tracking one) ---
        if pending is None and i >= 2 and not _time_at_or_after(row_time, NO_ENTRY_AFTER_HOUR, NO_ENTRY_AFTER_MINUTE):
            c1 = df.iloc[i - 2]
            c2 = df.iloc[i - 1]
            c3 = row  # current candle acts as c3

            avg_body = c2["_avg_body_prior"]
            avg_vol = c2["_avg_vol"]
            if pd.isna(avg_body) or pd.isna(avg_vol) or avg_body <= 0:
                continue
            c2_body = abs(c2["Close"] - c2["Open"])
            displaced = c2_body >= DISPLACEMENT_BODY_MULT * avg_body and c2["Volume"] >= DISPLACEMENT_VOL_MULT * avg_vol
            if not displaced:
                continue

            price_ref = float(c3["Close"])

            # All 3 candles must share the same direction (per the reference
            # image): bullish gap needs c1, c2, AND c3 all green; bearish
            # gap needs all 3 red. A big impulsive c2 alone isn't enough.
            all_bullish = c1["Close"] > c1["Open"] and c2["Close"] > c2["Open"] and c3["Close"] > c3["Open"]
            all_bearish = c1["Close"] < c1["Open"] and c2["Close"] < c2["Open"] and c3["Close"] < c3["Open"]

            if c1["High"] < c3["Low"] and all_bullish:
                gap_pct = (c3["Low"] - c1["High"]) / price_ref * 100
                if gap_pct >= MIN_GAP_PCT:
                    is_fvg = c3["Close"] <= c2["High"]  # stays within/touching c2's range
                    sig_id = str(uuid.uuid4())[:8]
                    all_signals.append({
                        "SignalId": sig_id, "Date": row.name.strftime("%Y-%m-%d"),
                        "DetectedTime": row.name.strftime("%H:%M:%S"), "Ticker": ticker,
                        "GapType": "FVG" if is_fvg else "BAG (not this strategy)",
                        "Direction": "BULLISH", "ZoneLow": round(float(c1["High"]), 2),
                        "ZoneHigh": round(float(c3["Low"]), 2), "EntryTriggered": False,
                        "EntryTime": "",
                    })
                    if is_fvg:
                        pending = {"direction": "BULLISH", "zone_top": float(c3["Low"]),
                                   "zone_bottom": float(c1["High"]),
                                   "c2_range": float(c2["High"] - c2["Low"]),
                                   "detected_idx": i, "signal_id": sig_id}
            elif c1["Low"] > c3["High"] and all_bearish:
                gap_pct = (c1["Low"] - c3["High"]) / price_ref * 100
                if gap_pct >= MIN_GAP_PCT:
                    is_fvg = c3["Close"] >= c2["Low"]
                    sig_id = str(uuid.uuid4())[:8]
                    all_signals.append({
                        "SignalId": sig_id, "Date": row.name.strftime("%Y-%m-%d"),
                        "DetectedTime": row.name.strftime("%H:%M:%S"), "Ticker": ticker,
                        "GapType": "FVG" if is_fvg else "BAG (not this strategy)",
                        "Direction": "BEARISH", "ZoneLow": round(float(c3["High"]), 2),
                        "ZoneHigh": round(float(c1["Low"]), 2), "EntryTriggered": False,
                        "EntryTime": "",
                    })
                    if is_fvg:
                        pending = {"direction": "BEARISH", "zone_top": float(c1["Low"]),
                                   "zone_bottom": float(c3["High"]),
                                   "c2_range": float(c2["High"] - c2["Low"]),
                                   "detected_idx": i, "signal_id": sig_id}

    return raw_trades, all_signals


def _build_trade(signal_id, ticker, direction, entry_time, entry_price, stop_price,
                  target, exit_time, exit_price, exit_reason):
    risk = abs(entry_price - stop_price)
    return {
        "SignalId": signal_id,
        "EntryDate": entry_time.strftime("%Y-%m-%d"), "EntryTime": entry_time.strftime("%H:%M:%S"),
        "Ticker": ticker, "Direction": direction,
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

            if t["Direction"] == "BULLISH":
                pnl = (t["ExitPrice"] - t["Entry"]) * qty
            else:
                pnl = (t["Entry"] - t["ExitPrice"]) * qty
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
          f"{len(all_signals)} gaps detected ({len(all_raw_trades)} triggered an entry). "
          f"Applying position sizing...")

    final_trades, ending_capital, curve, taken_ids = apply_position_sizing(all_raw_trades)
    for s in all_signals:
        s["TakenAsTrade"] = s["SignalId"] in taken_ids
    return final_trades, ending_capital, curve, all_signals


# ------------------------------ EXCEL OUTPUT --------------------------------

TRADE_COLUMNS = ["EntryDate", "EntryTime", "Ticker", "Direction", "Entry", "InitialStop",
                  "Target", "Qty", "RiskAmount", "ExitTime", "ExitPrice", "ExitReason",
                  "Outcome", "PnL", "PnLPct", "CapitalAfter"]
SIGNAL_COLUMNS = ["Date", "DetectedTime", "Ticker", "GapType", "Direction", "ZoneLow",
                   "ZoneHigh", "EntryTriggered", "EntryTime", "TakenAsTrade"]


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
        ("Total Gaps Detected", len(all_signals)),
        ("Gaps That Triggered Entry", sum(1 for s in all_signals if s["EntryTriggered"])),
        ("No new entries after", f"{NO_ENTRY_AFTER_HOUR}:{NO_ENTRY_AFTER_MINUTE:02d}"),
        ("Mandatory exit by", f"{EXIT_HOUR}:{EXIT_MINUTE:02d}"),
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
                f"{len(all_signals)} total gaps detected"
            )
        else:
            backtest_status["result_summary"] = "No gaps or trades were generated -- check server logs."
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
            "FVG backtest service alive.\n\n"
            "Visit /run-backtest to start (optional: ?days=30&tickers=RELIANCE.NS,TCS.NS)\n"
            "Visit /status to check progress.\n"
            "Visit /download to get fvg_journal.xlsx once finished.\n"
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
            self.wfile.write(b"No fvg_journal.xlsx yet -- visit /run-backtest first.\n")
            return
        with open(OUTPUT_FILE, "rb") as f:
            data = f.read()
        self.send_response(200)
        self.send_header("Content-Type",
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", "attachment; filename=fvg_journal.xlsx")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, format, *args):
        pass


def start_web_server():
    port = int(os.environ.get("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), _Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    print(f"FVG backtest server listening on port {port}.")


def main():
    parser = argparse.ArgumentParser(description="Backtest the Fair Value Gap strategy")
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
            print("\nNo gaps or trades were generated -- check the [warn] lines above.")
            return
        write_excel(trades, BACKTEST_CAPITAL, ending_capital, curve, all_signals)
        wins = sum(1 for t in trades if t["Outcome"] == "WIN")
        wr = round(wins / len(trades) * 100, 1) if trades else 0
        print(f"\n{'='*60}\nBACKTEST COMPLETE")
        print(f"  Total gaps detected: {len(all_signals)}")
        print(f"  Total trades:        {len(trades)}")
        print(f"  Win rate:            {wr}%")
        print(f"  Starting capital:    Rs.{BACKTEST_CAPITAL:,.2f}")
        print(f"  Ending capital:      Rs.{ending_capital:,.2f}")
        print(f"  Max drawdown:        {max_drawdown_pct(curve)}%")
        print(f"  Full detail in:      {OUTPUT_FILE}\n{'='*60}")
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
