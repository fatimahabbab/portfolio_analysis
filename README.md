# Portfolio Analysis Project

## Overview

This project performs portfolio analysis on a set of stock tickers, including data collection, returns calculation, portfolio performance evaluation, and visualization of key metrics using Excel and Python.

## Files and Folders

- **`PortfolioAnalysis.xlsx`**: The main Excel workbook containing all data and analysis.
- **`fetch_data.py`**: Python script to fetch historical stock data from Alpha Vantage.
- **`README.md`**: This file, providing an overview and instructions.

## Excel Workbook Structure

### Sheet 1: Stock List

Contains a list of stock tickers and their corresponding company names.

| Ticker | Company Name          |
|--------|-----------------------|
| AAPL   | Apple Inc.            |
| MSFT   | Microsoft Corporation |
| GOOGL  | Alphabet Inc.         |
| AMZN   | Amazon.com, Inc.      |
| FB     | Meta Platforms, Inc.  |
| ...    | ...                   |

### Sheet 2: Raw Data

Contains historical stock data for each ticker, imported from CSV files.

| Date      | Open  | High  | Low   | Close | Adj Close |
|-----------|-------|-------|-------|-------|-----------|
| 1/2/2023  | 133.52| 134.64| 131.48| 133.00| 132.69    |
| 1/3/2023  | 133.00| 135.53| 132.01| 134.68| 134.36    |
| ...       | ...   | ...   | ...   | ...   | ...       |

### Sheet 3: Returns Calculation

Calculates daily returns for each stock.

| Date     | AAPL  | MSFT  | GOOGL | AMZN  | FB    |
|----------|-------|-------|-------|-------|-------|
| 1/2/2023 |       |       |       |       |       |
| 1/3/2023 | 0.0200| 0.0150| 0.0100| 0.0180| 0.0120|
| ...      | ...   | ...   | ...   | ...   | ...   |

Daily returns formula: `(Adj Close Today - Adj Close Yesterday) / Adj Close Yesterday`

### Sheet 4: Portfolio Analysis

Evaluates portfolio returns, mean return, volatility, and Sharpe ratio.

| Date     | Daily Return |
|----------|--------------|
| 1/2/2023 |              |
| 1/3/2023 | 0.0150       |
| ...      | ...          |

| Metric       | Value                   |
|--------------|-------------------------|
| Mean Return  | =AVERAGE(B2:B100)       |
| Volatility   | =STDEV(B2:B100)         |
| Sharpe Ratio | =(Mean Return - Risk Free Rate) / Volatility |

### Sheet 5: Dashboard

Summarizes key metrics and visualizations for easy interpretation.

| Metric        | Value  |
|---------------|--------|
| Mean Return   | 0.0102 |
| Volatility    | 0.0187 |
| Sharpe Ratio  | 0.5432 |

| Date     | Portfolio Value |
|----------|-----------------|
| 1/2/2023 | 100000          |
| 1/3/2023 | 101500          |
| ...      | ...             |

## Usage Instructions

### Setting Up the Excel Workbook

1. Open `PortfolioAnalysis.xlsx`.
2. In the **Stock List** sheet, update the list of stock tickers and company names as needed.
3. Use the **fetch_data.py** script to download historical stock data and import it into the **Raw Data** sheet.
4. Ensure the **Returns Calculation** sheet formulas are correctly referencing the appropriate columns in the **Raw Data** sheet.
5. Adjust the portfolio weights in the **Portfolio Analysis** sheet.
6. Review the **Dashboard** sheet for key metrics and visualizations.

### Running the Python Script

1. Ensure you have Python installed on your system.
2. Install necessary libraries by running:
   ```sh
   pip install pandas requests
