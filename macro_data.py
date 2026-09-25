import contextlib
import io
import yfinance as yf


def validate_and_get_company_info(ticker_symbol: str):
    """
    Validates the ticker with Yahoo Finance while suppressing raw HTTP error dumps.
    Returns (True, info_dict) if valid, or (False, None) if invalid.
    """
    try:
        # Suppress stderr to keep terminal clean from raw 404 logs
        with contextlib.redirect_stderr(io.StringIO()):
            t = yf.Ticker(ticker_symbol)
            info = t.info

            if not info or ("shortName" not in info and "regularMarketPrice" not in info and "currentPrice" not in info):
                hist = t.history(period="5d")
                if hist.empty:
                    return False, None

            return True, info
    except Exception:
        return False, None


def prompt_float(initial_prompt: str, default_val, field_name: str) -> float:
    """
    Prompts for a percentage and returns it as a fraction (6.5 -> 0.065).
    If default_val is None the field is REQUIRED: blank input re-prompts instead of
    silently falling back to a default that may not fit the market.
    """
    val_str = input(initial_prompt).strip()
    while True:
        if not val_str:
            if default_val is None:
                val_str = input(f"Required, enter {field_name}: ").strip()
                continue
            return default_val / 100.0
        try:
            return float(val_str) / 100.0
        except ValueError:
            val_str = input(f"Invalid, enter {field_name} again: ").strip()


def get_macro_inputs():
    """
    Prompts for target ticker, risk-free rate, and tax rate with strict input validation.
    Rf and tax rate only have defaults for India / US. Every other market must be entered
    explicitly, so a stale default can never leak into a model for the wrong country.
    """
    print("\n--- MODEL INPUTS & MACRO CONSTANTS ---")

    # 1. Ticker Input Loop
    ticker_input = input("Enter Target Ticker (e.g., RELIANCE.NS, AAPL): ").strip().upper()
    while True:
        if not ticker_input:
            ticker_input = input("Invalid, enter ticker again: ").strip().upper()
            continue

        is_valid, info = validate_and_get_company_info(ticker_input)
        if is_valid:
            ticker = ticker_input
            company_name = info.get("shortName") or info.get("longName") or ticker
            country = info.get("country")
            if not country:
                if ticker.endswith(".NS") or ticker.endswith(".BO"):
                    country = "India"
                else:
                    country = input("Country not returned by Yahoo Finance. Enter country: ").strip() or "Unknown"
            print(f"✓ Found: {company_name} | Country: {country}")
            break

        ticker_input = input("Invalid, enter ticker again: ").strip().upper()

    # Regional defaults exist only where we have a defensible starting point
    is_india = country.lower() == "india" or ticker.endswith(".NS") or ticker.endswith(".BO")
    is_us = country.lower() == "united states"
    default_rf = 6.95 if is_india else 4.25 if is_us else None
    default_tax = 25.17 if is_india else 21.00 if is_us else None

    rf_hint = (f"[Default {default_rf}% - verify vs current 10Y yield]" if default_rf is not None
               else "[required - 10Y government yield in the currency of the cash flows]")
    tax_hint = (f"[Default {default_tax}%]" if default_tax is not None
                else "[required - marginal corporate rate]")

    # 2. Risk-Free Rate Input
    risk_free_rate = prompt_float(
        initial_prompt=f"Enter Risk-Free Rate % {rf_hint}: ",
        default_val=default_rf,
        field_name="risk-free rate"
    )

    # 3. Corporate Tax Rate Input
    tax_rate = prompt_float(
        initial_prompt=f"Enter Marginal Tax Rate % {tax_hint}: ",
        default_val=default_tax,
        field_name="tax rate"
    )

    return {
        "ticker": ticker,
        "company_name": company_name,
        "country": country,
        "risk_free_rate": risk_free_rate,
        "tax_rate": tax_rate
    }


if __name__ == "__main__":
    constants = get_macro_inputs()
    print("\n=== CONFIRMED MODEL CONSTANTS ===")
    print(f"Company:        {constants['company_name']}")
    print(f"Ticker:         {constants['ticker']}")
    print(f"Country:        {constants['country']}")
    print(f"Risk-Free Rate: {constants['risk_free_rate'] * 100:.2f}%")
    print(f"Tax Rate:       {constants['tax_rate'] * 100:.2f}%")