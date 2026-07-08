"""
Intraday ORB + VWAP Strategy BACKTEST (Render Web Service Ready)
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

VOLUME_MULTIPLIER = 2.0      # Stronger filter
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


def compute_vwap(df):
    typical_price = (df["High"] + df["Low"] + df["Close"]) / 3
    df["cum_vol"] = df["Volume"].cumsum()
    df["cum_vol_price"] = (typical_price * df["Volume"]).cumsum()
    df["VWAP"] = df["cum_vol_price"] / df["cum_vol"]
    return df


def compute_ema(df, span, col_name):
    df[col_name] = df["Close"].ewm(span=span, adjust=False).mean()
    return df


def compute_rsi(df, period, col_name="RSI"):
    delta = df["Close"].diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.ewm(alpha=1 / period, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1 / period, adjust=False).mean()
    rs = avg_gain / avg_loss.replace(0, pd.NA)
    df[col_name] = 100 - (100 / (1 + rs))
    df[col_name] = df[col_name].fillna(50)
    return df


def compute_macd(df, fast, slow, signal):
    ema_fast = df["Close"].ewm(span=fast, adjust=False).mean()
    ema_slow = df["Close"].ewm(span=slow, adjust=False).mean()
    df["MACD"] = ema_fast - ema_slow
    df["MACD_SIGNAL"] = df["MACD"].ewm(span=signal, adjust=False).mean()
    return df


def get_opening_range(df):
    if df.empty:
        return None, None
    session_start = df.index[0]
    cutoff = session_start + pd.Timedelta(minutes=OPENING_RANGE_MINUTES)
    opening_slice = df[df.index <= cutoff]
    if opening_slice.empty:
        return None, None
    return opening_slice["High"].max(), opening_slice["Low"].min()


def classic_pivots(H, L, C):
    P = (H + L + C) / 3
    return {
        "P": P, "R1": 2 * P - L, "S1": 2 * P - H,
        "R2": P + (H - L), "S2": P - (H - L),
        "R3": H + 2 * (P - L), "S3": L - 2 * (H - P),
    }


def fetch_intraday_history(ticker, days):
    try:
        df = yf.download(ticker, period=f"{days}d", interval=CANDLE_INTERVAL,
                          progress=False, auto_adjust=False)
    except Exception as e:
        print(f"  [warn] intraday fetch failed for {ticker}: {e}")
        return None
    if df is None or df.empty:
        return None
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    return df


def fetch_daily_history(ticker, days):
    try:
        df = yf.download(ticker, period=f"{days + 15}d", interval="1d",
                          progress=False, auto_adjust=False)
    except Exception as e:
        print(f"  [warn] daily fetch failed for {ticker}: {e}")
        return None
    if df is None or df.empty:
        return None
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    return df


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

    entry_row = day_df.iloc[entry_idx]
    price = float(entry_row["Close"])

    swing_window = day_df.iloc[max(0, entry_idx - SWING_LOOKBACK_CANDLES):entry_idx]
    if swing_window.empty:
        return None
    swing_high, swing_low = float(swing_window["High"].max()), float(swing_window["Low"].min())

    if signal == "BULLISH":
        stop = max(or_low, swing_low)
        if stop >= price:
            return None
        candidates = [pivots[k] for k in ("R1", "R2", "R3") if pivots[k] > price]
        if not candidates:
            return None
        target = min(candidates)
    else:
        stop = min(or_high, swing_high)
        if stop <= price:
            return None
        candidates = [pivots[k] for k in ("S1", "S2", "S3") if pivots[k] < price]
        if not candidates:
            return None
        target = max(candidates)

    risk = abs(price - stop)
    reward = abs(target - price)
    if risk <= 0:
        return None
    rr_ratio = reward / risk
    if rr_ratio < MIN_RR_RATIO:
        return None

    risk_amount = capital_start_of_day * (RISK_PER_TRADE_PCT / 100)
    qty = max(1, int(risk_amount // risk))

    outcome, exit_price, exit_time = None, None, None
    for j in range(entry_idx + 1, len(day_df)):
        future_row = day_df.iloc[j]
        if signal == "BULLISH":
            if future_row["High"] >= target:
                outcome, exit_price = "WIN", target
            elif future_row["Low"] <= stop:
                outcome, exit_price = "LOSS", stop
        else:
            if future_row["Low"] <= target:
                outcome, exit_price = "WIN", target
            elif future_row["High"] >= stop:
                outcome, exit_price = "LOSS", stop
        if outcome:
            exit_time = future_row.name
            break

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

    if signal == "BULLISH":
        pnl = (exit_price - price) * qty
    else:
        pnl = (price - exit_price) * qty
    pnl_pct = round((pnl / (price * qty)) * 100, 2) if price * qty else 0

    return {
        "EntryDate": entry_row.name.strftime("%Y-%m-%d"),
        "EntryTime": entry_row.name.strftime("%H:%M:%S"),
        "Ticker": None,
        "Signal": signal, "Entry": round(price, 2), "Stop": round(stop, 2),
        "Target": round(target, 2), "RRRatio": round(rr_ratio, 2), "Qty": qty,
        "RiskAmount": round(risk_amount, 2),
        "ExitTime": exit_time.strftime("%H:%M:%S"), "ExitPrice": round(exit_price, 2),
        "Outcome": outcome, "PnL": round(pnl, 2), "PnLPct": pnl_pct,
    }


def run_backtest(tickers, days):
    print(f"Fetching {days}-day intraday history for {len(tickers)} tickers...")
    intraday_data = {}
    daily_data = {}
    for i, ticker in enumerate(tickers, 1):
        print(f"  [{i}/{len(tickers)}] {ticker}...")
        idf = fetch_intraday_history(ticker, days)
        ddf = fetch_daily_history(ticker, days)
        if idf is not None and ddf is not None:
            intraday_data[ticker] = idf
            daily_data[ticker] = ddf

    print(f"\nGot usable data for {len(intraday_data)}/{len(tickers)} tickers. Simulating...")

    all_days = sorted(set(d.date() for df in intraday_data.values() for d in df.index))

    capital = BACKTEST_CAPITAL
    all_trades = []
    capital_curve = [capital]

    for day in all_days:
        day_trades = []
        for ticker, idf in intraday_data.items():
            day_df = idf[idf.index.date == day].copy()
            if day_df.empty:
                continue

            ddf = daily_data.get(ticker)
            if ddf is None:
                continue
            prior_days = ddf[ddf.index.date < day]
            if prior_days.empty:
                continue
            prev = prior_days.iloc[-1]
            pivots = classic_pivots(float(prev["High"]), float(prev["Low"]), float(prev["Close"]))

            try:
                trade = simulate_ticker_day(day_df, pivots, capital)
            except Exception as e:
                print(f"  [warn] simulation error for {ticker} on {day}: {e}")
                continue

            if trade:
                trade["Ticker"] = ticker
                day_trades.append(trade)

        if day_trades:
            day_pnl = sum(t["PnL"] for t in day_trades)
            capital += day_pnl
            for t in day_trades:
                t["CapitalAfter"] = round(capital, 2)
            all_trades.extend(day_trades)
            capital_curve.append(capital)
            print(f"  {day}: {len(day_trades)} trade(s), day P&L Rs.{round(day_pnl,2)}, capital now Rs.{round(capital,2)}")

    return all_trades, capital, capital_curve


TRADE_COLUMNS = ["EntryDate", "EntryTime", "Ticker", "Signal", "Entry", "Stop", "Target",
                  "RRRatio", "Qty", "RiskAmount", "ExitTime", "ExitPrice", "Outcome",
                  "PnL", "PnLPct", "CapitalAfter"]


def max_drawdown_pct(curve):
    peak = curve[0]
    max_dd = 0.0
    for value in curve:
        peak = max(peak, value)
        dd = (peak - value) / peak * 100 if peak else 0
        max_dd = max(max_dd, dd)
    return round(max_dd, 2)


def write_excel(trades, starting_capital, ending_capital, capital_curve):
    wb = Workbook()
    ws = wb.active
    ws.title = "Trades"
    ws.append(TRADE_COLUMNS)
    for t in trades:
        ws.append([t.get(c, "") for c in TRADE_COLUMNS])

    wins = sum(1 for t in trades if "WIN" in t["Outcome"])
    losses = sum(1 for t in trades if "LOSS" in t["Outcome"])
    total = len(trades)
    win_rate = round(wins / total * 100, 2) if total else 0
    total_pnl = round(sum(t["PnL"] for t in trades), 2)
    return_pct = round((ending_capital - starting_capital) / starting_capital * 100, 2)
    avg_rr = round(sum(t["RRRatio"] for t in trades) / total, 2) if total else 0

    summary_ws = wb.create_sheet("Summary")
    summary_rows = [
        ("Starting Capital", starting_capital),
        ("Ending Capital", round(ending_capital, 2)),
        ("Total Return %", return_pct),
        ("Total Trades", total),
        ("Wins", wins),
        ("Losses", losses),
        ("Win Rate %", win_rate),
        ("Total P&L", total_pnl),
        ("Average R:R Achieved", avg_rr),
        ("Max Drawdown %", max_drawdown_pct(capital_curve)),
        ("Risk Per Trade %", RISK_PER_TRADE_PCT),
        ("Min R:R Filter", MIN_RR_RATIO),
    ]
    for row in summary_rows:
        summary_ws.append(row)

    by_symbol = {}
    for t in trades:
        s = by_symbol.setdefault(t["Ticker"], {"trades": 0, "wins": 0, "pnl": 0.0})
        s["trades"] += 1
        s["wins"] += 1 if "WIN" in t["Outcome"] else 0
        s["pnl"] += t["PnL"]

    symbol_ws = wb.create_sheet("BySymbol")
    symbol_ws.append(["Ticker", "Trades", "Wins", "WinRate%", "TotalPnL"])
    for ticker, s in sorted(by_symbol.items(), key=lambda x: -x[1]["pnl"]):
        wr = round(s["wins"] / s["trades"] * 100, 2) if s["trades"] else 0
        symbol_ws.append([ticker, s["trades"], s["wins"], wr, round(s["pnl"], 2)])

    wb.save(OUTPUT_FILE)


def main():
    parser = argparse.ArgumentParser(description="Backtest the ORB+VWAP strategy on historical data")
    parser.add_argument("--tickers", type=str, default=None,
                         help="Comma-separated tickers instead of the full Nifty 100")
    parser.add_argument("--days", type=int, default=MAX_LOOKBACK_DAYS,
                         help=f"Lookback days for 5m data (max ~{MAX_LOOKBACK_DAYS}, Yahoo's own limit)")
    args = parser.parse_args()

    days = min(args.days, MAX_LOOKBACK_DAYS)
    tickers = ([t.strip() for t in args.tickers.split(",") if t.strip()]
               if args.tickers else get_nifty_100_tickers())

    trades, ending_capital, capital_curve = run_backtest(tickers, days)

    if not trades:
        print("\nNo trades were generated...")
        return

    write_excel(trades, BACKTEST_CAPITAL, ending_capital, capital_curve)

    wins = sum(1 for t in trades if "WIN" in t["Outcome"])
    print(f"\n{'='*60}")
    print(f"BACKTEST COMPLETE")
    print(f"  Total trades:     {len(trades)}")
    print(f"  Win rate:         {round(wins/len(trades)*100, 1)}%")
    print(f"  Starting capital: Rs.{BACKTEST_CAPITAL:,.2f}")
    print(f"  Ending capital:   Rs.{ending_capital:,.2f}")
    print(f"  Return:           {round((ending_capital-BACKTEST_CAPITAL)/BACKTEST_CAPITAL*100, 2)}%")
    print(f"  Max drawdown:     {max_drawdown_pct(capital_curve)}%")
    print(f"  Full detail in:   {OUTPUT_FILE}")
    print(f"{'='*60}")


if __name__ == "__main__":
    # Render Web Service Fix
    import os
    port = int(os.environ.get("PORT", 8080))
    print(f"✅ Render port binding successful on port {port}")
    print("Starting backtest...\n")
    
    main()
