# VBA-challenge
This repository contains VBA scripts for Multiple Year Stock Data excel sheet (Module 2 Challenge).

# Project Overview
This Excel VBA project automates the analysis of stock data across multiple worksheets, representing different financial quarters. The script performs the following tasks:
  1. Loops through all worksheets (each representing a different quarter).
  2. Extracts the ticker symbol for each stock.
  3. Calculates:
       - The quarterly price change (difference between the opening price at the start of the quarter and the closing price at the end).
       - The percentage change in price from the opening to the closing of each quarter.
       - The total stock volume traded during the quarter.
  4. Identifies:
       - The stock with the greatest percentage increase.
       - The stock with the greatest percentage decrease.
       - The stock with the greatest total volume.
    
         
# Features
- Quarterly Stock Analysis: Automatically processes stock data for each quarter to provide detailed performance metrics.
- Efficiency Enhancements: Uses Application.ScreenUpdating = False and Application.Calculation = xlCalculationManual to optimize performance by reducing unnecessary screen updates and recalculations.
- Summary Output: Outputs key findings (greatest percentage change, highest volume) for easy interpretation.

# Requirements
 - Excel with VBA enabled.
 - .xlsm file format to support macros.

# Installation and Setup
1. Clone the Repository:
    - Create a new repository in GitHub named VBA-challenge.
    - Upload the provided .xlsm and .bas files from the project folder.
2. Running the Script:
    - Download the .xlsm file and open it in Excel.
    - Ensure macros are enabled.
    - To execute the script, go to the VBA editor (Alt + F11), find the module containing the script, and run Q1 from the macro menu.

# Usage Instructions
- Ensure that the stock data for each quarter is placed in separate worksheets.
- The script will automatically process all worksheets.
- Once the script is executed, results will be printed in new columns, including:
   - Ticker symbols.
   - Quarterly changes and percentage changes.
   - Total stock volumes.
Additionally, a summary table will display the stocks with the greatest changes and volume across all worksheets.
