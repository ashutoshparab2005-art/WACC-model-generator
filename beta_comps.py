import contextlib
import io
import yfinance as yf
import pandas as pd
import numpy as np
from macro_data import validate_and_get_company_info, prompt_float

def get_market_benchmark(country: str, ticker: str) -> str:
    """Selects the standard equity benchmark index using global mappings."""
    suffix = ticker.split('.')[-1] if '.' in ticker else ""
    
    # Global Index Dictionary mapping Yahoo Finance suffixes
    suffix_map = {
        "NS": "^NSEI",   # India (NSE)
        "BO": "^BSESN",  # India (BSE)
        "SI": "^STI",    # Singapore
        "HK": "^HSI",    # Hong Kong
        "AX": "^AXJO",   # Australia
        "L": "^FTSE",    # UK
        "DE": "^GDAXI",  # Germany
        "T": "^N225",    # Japan
        "TO": "^GSPTSE", # Canada
        "PA": "^FCHI",   # France
        "MI": "^FTSEMIB",# Italy
        "SA": "^BVSP",   # Brazil
        "KS": "^KS11",   # South Korea
        "TW": "^TWII"    # Taiwan
    }
    
    # Check if market is standard US/India to skip the manual override prompt
    is_core = country.lower() in ["india", "united states"] or suffix in ["NS", "BO"]
    
    if suffix in suffix_map:
        default_bench = suffix_map[suffix]
    elif country.lower() == "india":
        default_bench = "^NSEI"
    elif country.lower() == "united states":
        default_bench = "^GSPC"
    else:
        default_bench = "^GSPC"  # Global fallback
        
    # Prompt for manual override if outside standard core markets
    if not is_core:
        print(f"\n  [?] Detected {country} for {ticker}. Mapped to local index: {default_bench}")
        override = input(f"  Enter Benchmark Index [Default {default_bench}]: ").strip().upper()
        return override if override else default_bench
        
    return default_bench

def prompt_period_and_frequency():
    print("\n--- REGRESSION CONFIGURATION ---")
    period_str = input("Enter lookback period in years [Default 5]: ").strip()
    while True:
        if not period_str:
            years = 5
            break
        try:
            years = int(period_str)
            if years > 0:
                break
            period_str = input("Invalid, enter years again (must be > 0): ").strip()
        except ValueError:
            period_str = input("Invalid, enter years again: ").strip()

    print("Available frequencies: [1] Monthly (1mo) | [2] Weekly (1wk) | [3] Daily (1d)")
    freq_choice = input("Select sampling frequency [Default 1 (Monthly)]: ").strip()
    while True:
        if freq_choice in ["", "1", "1mo", "monthly", "Monthly"]:
            freq = "1mo"
            break
        elif freq_choice in ["2", "1wk", "weekly", "Weekly"]:
            freq = "1wk"
            break
        elif freq_choice in ["3", "1d", "daily", "Daily"]:
            freq = "1d"
            break
        else:
            freq_choice = input("Invalid, enter choice again (1, 2, or 3): ").strip()

    return years, freq

def prompt_peers(macro_constants: dict) -> list:
    print("\n--- PEER GROUP SELECTION ---")
    target_ticker = macro_constants["ticker"]
    country_taxes = {macro_constants["country"]: macro_constants["tax_rate"]}
    
    num_str = input("How many peer companies to compare? [Default 4]: ").strip()
    while True:
        if not num_str:
            num_peers = 4
            break
        try:
            num_peers = int(num_str)
            if num_peers > 0:
                break
            num_str = input("Invalid, enter number of peers again: ").strip()
        except ValueError:
            num_str = input("Invalid, enter number of peers again: ").strip()

    peers_data = []
    print(f"Please enter {num_peers} peer ticker symbols:")

    for i in range(1, num_peers + 1):
        p_input = input(f"Enter Peer {i} Ticker: ").strip().upper()
        while True:
            if not p_input:
                p_input = input(f"Invalid, enter Peer {i} ticker again: ").strip().upper()
                continue
            if p_input == target_ticker:
                p_input = input(f"Target cannot be its own peer. Enter again: ").strip().upper()
                continue
            if any(p['ticker'] == p_input for p in peers_data):
                p_input = input(f"Ticker already added. Enter again: ").strip().upper()
                continue

            is_valid, info = validate_and_get_company_info(p_input)
            if is_valid:
                p_name = info.get("shortName") or info.get("longName") or p_input
                p_country = info.get("country", "Unknown")
                
                # Fetch local benchmark for this specific peer
                p_bench = get_market_benchmark(p_country, p_input)
                
                print(f"  ✓ Added: {p_name} ({p_input}) | Country: {p_country} | Index: {p_bench}")
                
                # Cross-border tax check
                if p_country not in country_taxes:
                    print(f"\n  🌍 Cross-Border Peer Detected!")
                    print(f"  We need the corporate tax rate for {p_country} to unlever {p_input}'s beta.")
                    default_tax = 21.0 if p_country.lower() == "united states" else 25.0
                    
                    p_tax = prompt_float(
                        initial_prompt=f"  Enter {p_country} Marginal Tax Rate % [Default {default_tax}%]: ",
                        default_val=default_tax,
                        field_name="tax rate"
                    )
                    country_taxes[p_country] = p_tax
                    print(f"  ✓ {p_country} Tax Rate saved as {p_tax * 100:.2f}%\n")
                
                peers_data.append({
                    "ticker": p_input,
                    "country": p_country,
                    "tax_rate": country_taxes[p_country],
                    "benchmark": p_bench
                })
                break

            p_input = input(f"Invalid, enter Peer {i} ticker again: ").strip().upper()

    return peers_data

def fetch_historical_returns(tickers: list, benchmarks: list, years: int, freq: str) -> tuple:
    # Combine stock tickers and all unique benchmark indices
    all_tickers = list(dict.fromkeys(tickers + benchmarks))
    end_date = pd.Timestamp.today()
    start_date = end_date - pd.DateOffset(years=years)

    print(f"\nDownloading {years}Y {freq} data for {len(all_tickers)} symbols from Yahoo Finance...")
    with contextlib.redirect_stderr(io.StringIO()):
        data = yf.download(all_tickers, start=start_date, end=end_date, interval=freq, progress=False)

    if isinstance(data.columns, pd.MultiIndex):
        top_level = data.columns.get_level_values(0)
        second_level = data.columns.get_level_values(1) if data.columns.nlevels > 1 else []
        
        if "Adj Close" in top_level:
            adj_close = data["Adj Close"]
        elif "Close" in top_level:
            adj_close = data["Close"]
        elif "Adj Close" in second_level:
            adj_close = data.xs("Adj Close", axis=1, level=1)
        elif "Close" in second_level:
            adj_close = data.xs("Close", axis=1, level=1)
        else:
            adj_close = data
    else:
        if "Adj Close" in data.columns:
            adj_close = data["Adj Close"]
        elif "Close" in data.columns:
            adj_close = data["Close"]
        else:
            adj_close = data

    if isinstance(adj_close, pd.DataFrame):
        adj_close = adj_close.dropna(how="all", axis=1)
    
    # Keep NaNs: each regression later uses its own (stock, benchmark) pair. Dropping every row
    # that has a NaN anywhere would silently cut all betas down to the shortest history in the set.
    returns = adj_close.pct_change(fill_method=None)
    return adj_close, returns


EXPECTED_OBS_PER_YEAR = {"1mo": 12, "1wk": 52, "1d": 252}
SHORT_HISTORY_THRESHOLD = 0.75  # flag if fewer than 75% of the requested observations exist


def regress_beta(prices: pd.DataFrame, sym: str, bench: str) -> dict:
    """
    OLS beta of `sym` on `bench` using ONLY dates where both have a price. Mirrors what the
    Excel regression sheet does (returns between consecutive jointly-available dates), so the
    console summary and the workbook can never disagree.
    """
    out = {"beta": np.nan, "r2": np.nan, "se": np.nan, "n": 0}
    if sym not in prices.columns or bench not in prices.columns:
        return out
    pair = prices[[sym, bench]].dropna()
    rets = pair.pct_change().dropna()
    n = len(rets)
    out["n"] = n
    if n < 3:
        return out
    x, y = rets[bench].values, rets[sym].values
    sxx = ((x - x.mean()) ** 2).sum()
    if sxx <= 0:
        return out
    syy = ((y - y.mean()) ** 2).sum()
    beta = ((x - x.mean()) * (y - y.mean())).sum() / sxx
    r2 = (beta ** 2) * sxx / syy if syy > 0 else np.nan
    sse = max(syy - (beta ** 2) * sxx, 0.0)
    out.update(beta=beta, r2=r2, se=np.sqrt(sse / (n - 2) / sxx))
    return out


def prompt_exclusions(df: pd.DataFrame, years: int, freq: str) -> list:
    """
    Flags peers whose price history is too short for a meaningful beta (recent IPOs) and lets the
    user drop them. The target is never dropped, only warned about.
    """
    expected = years * EXPECTED_OBS_PER_YEAR[freq]
    floor = int(SHORT_HISTORY_THRESHOLD * expected)
    drop = []
    for _, r in df.iterrows():
        if r["Obs"] >= floor:
            continue
        lo, hi = r["Raw Levered Beta"] - 1.96 * r["Beta SE"], r["Raw Levered Beta"] + 1.96 * r["Beta SE"]
        msg = (f"\n  [!] {r['Ticker']}: only {r['Obs']} observations (expected ~{expected}). "
               f"Raw beta {r['Raw Levered Beta']:.2f}, approx. 95% CI [{lo:.2f}, {hi:.2f}], R² {r['R2']:.2f}.")
        print(msg)
        if r["Role"] == "Target":
            print("      Target beta is not used directly, but treat the regression sheet with caution.")
            continue
        ans = input(f"      Exclude {r['Ticker']} from the peer set? (y/n) [y]: ").strip().lower()
        if ans in ("", "y", "yes"):
            drop.append(r["Ticker"])
    return drop

def run_beta_comps(macro_constants: dict):
    target = macro_constants["ticker"]
    country = macro_constants["country"]

    years, freq = prompt_period_and_frequency()
    target_bench = get_market_benchmark(country, target)
    peers_data = prompt_peers(macro_constants)

    peer_tickers = [p["ticker"] for p in peers_data]
    peer_benchmarks = [p["benchmark"] for p in peers_data]

    all_symbols = [target] + peer_tickers
    all_benchmarks = [target_bench] + peer_benchmarks

    prices, returns = fetch_historical_returns(all_symbols, all_benchmarks, years, freq)

    comps_data = []

    for sym in all_symbols:
        # Determine local benchmark for this specific stock
        if sym == target:
            sym_bench = target_bench
            sym_tax = macro_constants["tax_rate"]
            sym_country = macro_constants["country"]
        else:
            p_match = next((p for p in peers_data if p["ticker"] == sym), None)
            sym_bench = p_match["benchmark"]
            sym_tax = p_match["tax_rate"]
            sym_country = p_match["country"]

        # Run regression against the local market (pairwise sample, same as the Excel sheet)
        reg = regress_beta(prices, sym, sym_bench)
        raw_levered_beta = reg["beta"]

        if pd.notnull(raw_levered_beta):
            adjusted_beta = (2/3) * raw_levered_beta + (1/3) * 1.0
        else:
            adjusted_beta = np.nan

        comps_data.append({
            "Ticker": sym,
            "Role": "Target" if sym == target else "Peer",
            "Country": sym_country,
            "Tax Rate": sym_tax,
            "Benchmark": sym_bench,
            "Raw Levered Beta": round(raw_levered_beta, 4),
            "Adjusted Beta": round(adjusted_beta, 4),
            "Obs": reg["n"],
            "R2": round(reg["r2"], 4) if pd.notnull(reg["r2"]) else np.nan,
            "Beta SE": round(reg["se"], 4) if pd.notnull(reg["se"]) else np.nan,
        })

    df = pd.DataFrame(comps_data)

    dropped = prompt_exclusions(df, years, freq)
    if dropped:
        remaining_peers = df[(df["Role"] == "Peer") & (~df["Ticker"].isin(dropped))]
        if remaining_peers.empty:
            print("  [!] Excluding these would leave no peers; keeping them.")
            dropped = []
        else:
            df = df[~df["Ticker"].isin(dropped)].reset_index(drop=True)
            print(f"  ✓ Excluded from peer set: {', '.join(dropped)}")

    return {
        "summary_table": df,
        "raw_prices": prices,
        "returns": returns,
        "benchmark": target_bench, # Target's benchmark is passed for the WACC CoE calculation
        "period_years": years,
        "frequency": freq,
        "excluded": dropped
    }

if __name__ == "__main__":
    from macro_data import get_macro_inputs
    macro = get_macro_inputs()
    beta_results = run_beta_comps(macro)

    print("\n=== RAW & ADJUSTED BETA SUMMARY ===")
    print(beta_results["summary_table"].to_string(index=False))