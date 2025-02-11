# Cryptocurrency Market Analysis Report

## Overview
This report provides an analysis of the top 50 cryptocurrencies by market capitalization. The data is sourced from the CoinGecko API and is automatically updated every 5 minutes to ensure real-time market insights.

## Data Collection
### Data Source
- API: CoinGecko Public API
- Update Frequency: Every 5 minutes
- Number of Cryptocurrencies: Top 50 by market capitalization

### Data Points Collected
1. Cryptocurrency Name
2. Symbol
3. Current Price (USD)
4. Market Capitalization
5. 24-hour Trading Volume
6. Price Change (24-hour percentage)

## Analysis Components

### 1. Market Leaders Analysis
- Top 5 cryptocurrencies by market capitalization are tracked and displayed in a dedicated Excel sheet
- Provides quick insight into market dominance and leadership
- Updates automatically to reflect market changes

### 2. Price Analysis
- Average price calculation across all 50 cryptocurrencies
- Helps identify overall market pricing trends
- Updated every 5 minutes to reflect market movements

### 3. Volatility Tracking
- Monitors 24-hour price changes
- Identifies cryptocurrencies with highest and lowest price changes
- Helps understand market volatility and momentum

## Technical Implementation

### Data Storage
- Format: Microsoft Excel (.xlsx)
- File Name: crypto_data_live.xlsx
- Sheet Structure:
  1. Live Data: Complete dataset for all 50 cryptocurrencies
  2. Analysis: Key metrics and statistical analysis
  3. Top 5 by Market Cap: Detailed view of market leaders

### Automation
- Python script runs continuously
- Automatic updates every 5 minutes
- Error handling for API and file system operations
- Formatted data presentation with proper currency and percentage formatting

## How to Use the Analysis Tools

### Excel Sheet (crypto_data_live.xlsx)
1. Open the Excel file to view current data
2. Data refreshes automatically every 5 minutes
3. Three sheets provide different levels of analysis:
   - Live Data: Complete market overview
   - Analysis: Key metrics and insights
   - Top 5: Focus on market leaders

### Python Script (crypto_analyzer.py)
1. Ensures continuous data updates
2. Handles all API communications
3. Manages Excel file updates
4. Provides error handling and logging

## Conclusions
This implementation provides a robust system for:
- Real-time cryptocurrency market monitoring
- Automated data collection and analysis
- Clear presentation of market insights
- Tracking of key market indicators

The system is designed to be:
- Reliable: Error handling and validation
- Current: 5-minute update frequency
- Comprehensive: Coverage of top 50 cryptocurrencies
- User-friendly: Clear Excel-based presentation


# Cryptocurrency Data Analyzer

This project fetches and analyzes live cryptocurrency data for the top 50 cryptocurrencies by market capitalization. It creates a live-updating Excel sheet with price data and analysis.

## Features

- Fetches live data from CoinGecko API
- Updates automatically every 5 minutes
- Tracks key metrics including:
  - Current Price (USD)
  - Market Capitalization
  - 24-hour Trading Volume
  - Price Change (24-hour percentage)
- Provides real-time analysis:
  - Top 5 cryptocurrencies by market cap
  - Average price of top 50 cryptocurrencies
  - Highest and lowest 24-hour price changes

## Setup

1. Install required packages:
   ```
   pip install -r requirements.txt
   ```

2. Run the analyzer:
   ```
   python crypto.py
   ```

## Output

The script generates an Excel file named `crypto_data_live.xlsx` with three sheets:
- Live Data: Current cryptocurrency data
- Analysis: Key metrics and statistics
- Top 5 by Market Cap: Detailed view of the top 5 cryptocurrencies

## Notes

- The Excel file updates every 5 minutes automatically
- Press Ctrl+C to stop the program
- Internet connection is required for live updates
- Uses the free CoinGecko API (no API key required)
