# VBA-challenge
Quarterly Stock Analysis VBA Script
Overview
This VBA script, QuarterlyStockAnalysis, is designed to analyze stock data across multiple quarters, summarizing each stock's performance and calculating key metrics such as quarterly change, percentage change, and total stock volume. It also tracks the stock with the greatest percentage increase, percentage decrease, and total volume for each quarter.

## How It Works
The script loops through each worksheet in the workbook that represents a quarter (identified by worksheet names starting with "Q"), processes the stock data, and outputs a summary for each quarter on the same worksheet.

## Key Metrics:
Ticker: Stock symbol being analyzed.
Quarterly Change: Difference between the closing and opening price of the stock in that quarter.
Percent Change: The percentage change from the opening to the closing price.
Total Stock Volume: The total volume of the stock traded during the quarter.

## In addition, the script highlights:
1. Greatest % Increase: The stock with the highest percentage increase.
2. Greatest % Decrease: The stock with the greatest percentage decrease.
3. Greatest Total Volume: The stock with the highest trading volume during the quarter.

## Summary Output
For each worksheet (quarter), the script outputs the following information:

  1. Ticker: In column H.
  2. Quarterly Change: In column I (colored green for positive and red for negative).
  3. Percent Change: In column J.
  4. Total Stock Volume: In column K.

It also outputs the stocks with:
1. Greatest % Increase: In cells P2 to R2.
2. Greatest % Decrease: In cells P3 to R3.
3. Greatest Total Volume: In cells P4 to R4.
   
## Installation and Setup
1. Open your Excel workbook containing quarterly stock data.
  Ensure that each quarter's data is stored in a separate worksheet, with the worksheet name starting with "Q" (e.g., "Q1", "Q2").
2. The stock data in each worksheet should be arranged as follows:
  Column A: Stock ticker
  Column C: Opening price
  Column F: Closing price
  Column G: Stock volume
3. Open the VBA editor (Alt + F11), and insert the script into a module.
4. Run the QuarterlyStockAnalysis macro to generate the summary for each quarter.
   
## Script Breakdown
1. Worksheet Loop: The script loops through all worksheets, focusing only on those named with "Q*".
2. Data Processing: For each stock, the script calculates the quarterly change, percent change, and total volume, and outputs this to the summary section on the worksheet.
3. Color Formatting: Positive quarterly changes are colored green, while negative changes are colored red.
4. Tracking Maximums: The script tracks the stock with the greatest percent increase, percent decrease, and total volume for each quarter and outputs this information separately.
   
## Example Usage
If you have stock data for multiple quarters in worksheets named "Q1", "Q2", etc., running this script will:
1. Analyze the opening and closing prices to calculate the quarterly changes.
2. Calculate the total trading volume for each stock.
3. Identify the stocks with the largest percentage gains and losses, as well as the highest total volume.
   
