# Cryptocurrency Live Data Fetcher

## Overview

The **Cryptocurrency Live Data Fetcher** is a Python application that fetches live data for the top 50 cryptocurrencies from the CoinMarketCap API, performs basic data analysis, and updates the results in a live-updating Excel spreadsheet. This tool is designed for cryptocurrency enthusiasts and analysts looking to monitor market trends efficiently.

## Features

- Fetches live cryptocurrency data including:
  - Cryptocurrency Name
  - Symbol
  - Current Price (in USD)
  - Market Capitalization
  - 24-hour Trading Volume
  - Price Change (24-hour, percentage)
  
- Performs basic analysis on the fetched data:
  - Identifies the top 5 cryptocurrencies by market capitalization
  - Calculates the average price of the top 50 cryptocurrencies
  - Analyzes the highest and lowest 24-hour percentage price changes
  
- Updates an Excel spreadsheet every minute with the latest data

## Requirements

- Python 3.x
- Required Python packages:
  - `requests`
  - `pandas`
  - `openpyxl`

## Installation

1. **Clone the Repository**:
   ```bash
   git clone <repository-url>
   cd <repository-directory>
