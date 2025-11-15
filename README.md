# Stock Data Analysis Tool

Python toolkit for pulling fundamentals and valuation metrics from multiple APIs and assembling them into an analysis-ready Excel sheet. The scripts automate calls to Seeking Alpha, Financial Modeling Prep (FMP), and Alpha Vantage, compute growth and quality ratios, and output a unified table suitable for screening and further modeling.

## What this tool does

- Fetches **cash-flow, income-statement, metrics, and valuation** data from Seeking Alpha for a list of tickers.
- Computes **cash flow growth**, **EPS growth**, and **revenue growth** using annual financials.
- Pulls **marketcap, P/E, price/cash-flow, price/book, EV/EBITDA, revenue growth, gross margin, net margin, ROE** from Seeking Alpha metrics.
- Enriches the dataset with **key metrics and TTM ratios** (P/E, price-to-cash-flow, price-to-sales, ROE) from Financial Modeling Prep.
- Adds **quarterly EPS history** from Alpha Vantage (current and previous quarters, plus a couple of estimated EPS points).
- Derives blended and change metrics such as **Blended_Avg_PE** and **P/E Growth**.
- Writes everything into a **structured Excel sheet** (plus optional CSV/TXT in some scripts), matching a predefined column layout used for manual review and downstream models.

## Core workflow

1. **User inputs**  
   - Seeking Alpha API key (RapidAPI)  
   - Financial Modeling Prep API key  
   - Alpha Vantage API key  
   - Target Excel file path and sheet name  
   - Comma-separated list of stock symbols to analyze  

2. **Seeking Alpha financials and metrics**  
   - Cash-flow endpoint: reads the *Cash Flow From Operating Activities* section and computes cash-flow growth over selected years.  
   - Income-statement endpoint: extracts **diluted EPS** and **total revenue** for 2022 and 2023 and computes percentage growth.  
   - Metrics endpoint: pulls marketcap, valuation and profitability fields, and pivots into a stock-by-metric table.  
   - Valuation endpoint: reads **priceSales** and **pegRatio** for each symbol.

3. **FMP key metrics and ratios**  
   - `key-metrics` endpoint: grabs historical P/E ratios, computes **Latest_PE**, **Last_Year_PE**, and a **Blended_Avg_PE**.  
   - `ratios-ttm` endpoint: fetches TTM ratios like **P/E Ratio TTM**, **Price to Cash Flow Ratio TTM**, **Price to Sales Ratio TTM**, **Return on Equity TTM**.

4. **Alpha Vantage EPS series**  
   - `EARNINGS` function: reads `quarterlyEarnings` for each symbol and collects:  
     - `current_eps`  
     - `prev_eps_1`, `prev_eps_2`, `prev_eps_3`  
     - `est_eps_1`, `est_eps_2` (additional EPS figures used in the sheet)

5. **Data assembly and shaping**  
   - Merges the Seeking Alpha metrics, FMP key metrics/ratios, and Alpha Vantage EPS data on `Symbol`.  
   - Computes **P/E Growth** from `Latest_PE` and `Last_Year_PE`.  
   - Adds placeholder columns (e.g., `EPS-g TTM-manual`, `Gross Margin TTM`, `Operating_NI_margin`) for manual population.  
   - Reorders columns to match an existing Excel template and writes them starting at a fixed row (e.g., row 5) on the chosen sheet.

## Key modules and functions

- `initialize_connection()`  
  Creates an HTTPS connection to Seeking Alpha via RapidAPI and returns connection + headers.

- `make_api_request(endpoint)`  
  Generic wrapper around the Seeking Alpha endpoints. Sends GET, parses JSON, closes the connection, and returns the decoded object.

- `get_jsonparsed_data(url)`  
  Uses `urllib` with an explicit SSL context (via `certifi`) to fetch and parse JSON from FMP endpoints.

- `calculate_metrics(symbol)`  
  Central function for a single ticker:
  - Calls cash-flow, income-statement, metrics, and valuation endpoints.
  - Computes cash-flow growth, EPS and revenue growth, and gathers all numeric metrics into a `stock_data` dict.

- `get_combined_data(symbols, api_key1, api_key2)`  
  Combines FMP key metrics, FMP TTM ratios, and Alpha Vantage EPS series for the given list of symbols.

- `calculate_pe_growth(row)`  
  Computes the relative change between `Latest_PE` and `Last_Year_PE`.

## Setup

### Requirements

- Python 3.9+ (tested with 3.x)
- Packages:
  - `pandas`
  - `requests`
  - `certifi`
  - `openpyxl` (for Excel writing)
- Network access to:
  - Seeking Alpha (via RapidAPI)
  - Financial Modeling Prep
  - Alpha Vantage

Install dependencies (example):

```bash
pip install pandas requests certifi openpyxl
