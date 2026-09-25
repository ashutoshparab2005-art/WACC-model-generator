# WACC-model-generator
A Python command-line tool that generates a fully formula-linked, institutional-style WACC model in Excel, pulling live data from Yahoo Finance instead of requiring manual data entry.
Given a target ticker and a set of peer companies, the tool builds a multi-tab workbook — WACC summary, cost of equity, and one beta regression sheet per company — with live Excel formulas rather than pasted values, so every number can be traced back to its source and audited.

What it does
Peer comps & beta. Runs beta regressions for the target and each peer against its own local market index, over a user-chosen lookback period and sampling frequency. Reports R², standard error, and a 95% confidence interval for every beta, and flags peers with too little price history (e.g. recent IPOs) so you can decide whether to exclude them.
Unlevering / relevering. Unlevers each peer's beta using its own debt/equity and tax rate, takes the peer median (excluding the target), and relevers it to the target's own capital structure.
Cost of debt. Three modes: risk-free rate plus a credit spread (derived from a synthetic credit rating based on interest coverage), historical interest expense ÷ total debt, or a manual input. Flags cases where the resulting cost of debt falls at or below the risk-free rate, since that's not economically defensible.
Equity risk premium. Taken as an explicit, sourced input rather than derived from noisy historical index returns. The historical arithmetic and geometric estimates (with standard error) are still shown on the Cost of Equity sheet for reference.
Currency handling. Converts debt into a company's trading currency when its reporting currency differs (common for cross-listed or foreign-domiciled peers).
Output. CAPM-based cost of equity, final WACC, and a sensitivity table across beta and ERP — all as live formulas in an auditable .xlsx file.
Usage
bash
pip install yfinance openpyxl pandas numpy
python main.py

The tool runs interactively: enter the target ticker, risk-free rate, tax rate, cost-of-debt method, peer tickers, and ERP when prompted. The finished workbook is saved to output/WACC_Model_<TICKER>_<date>.xlsx.

Open it in Excel and save once before sharing — previewers (GitHub, Google Drive, etc.) show cached values, and a fresh file has none until Excel calculates and saves it.

Files
File	Purpose
main.py	Orchestrates the run and builds the Excel workbook (WACC, Cost of Equity, and regression sheets).
beta_comps.py	Handles peer selection, benchmark mapping, historical price/return fetching, and beta regressions.
macro_data.py	Prompts for and validates the target ticker, risk-free rate, and tax rate.
Known limitations
Each company is regressed against its own local index (e.g. NSE peers against NSEI, Hong Kong peers against HSI), so betas from different markets aren't on a strictly comparable scale when peers span countries.
The credit-spread table used for the synthetic rating is an approximation modeled on public interest-coverage/rating tables and should be refreshed against a current source before being cited in any serious context.
Peer selection is left entirely to the user — the tool flags weak or short-history betas but can't judge whether a peer is actually comparable to the target.
Market data is only as reliable as what Yahoo Finance returns; sparse or missing financials (particularly for smaller or newly listed companies) can leave gaps that require manual input.
