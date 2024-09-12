# VBA-challenge
This repository contains VBA scripts for Multiple Year Stock Data excel sheet (Module 2 Challenge).

# Project Overview
This Excel VBA project automates the analysis of stock data across multiple worksheets (each representing different quarters). The script performs the following tasks:
  1. Loops through all worksheets (each representing a different quarter).
  2. Extracts the ticker symbol for each stock.
  3. Calculates:
       - The quarterly price change (difference between the opening price at the start of the quarter and the closing price at the end which is based on each tricker symbol).
       - The percentage change in price from the opening to the closing of each quarter.
       - The total stock volume traded during the quarter.
  4. Identifies:
       - The ticker with the greatest percentage increase.
       - The ticker with the greatest percentage decrease.
       - The ticker with the greatest total volume.
    
         
# Features
- Quarterly Stock Analysis: Automatically processes stock data for each quarter to provide detailed performance metrics.
- Efficiency Enhancements: Uses Application.ScreenUpdating = False and Application.Calculation = xlCalculationManual to optimize performance by reducing unnecessary screen updates and recalculations.
    - I used this link (https://stackoverflow.com/questions/58799143/how-to-speed-up-the-following-excel-vba-code) to enhance the code since my excel way timing out.
- Summary Output: Outputs key findings (greatest percentage change, highest volume) for easy interpretation.

# Installation
Running the Script:
  - Download the .xlsm file and open it in Excel.
      1. Open your terminal (Git Bash, Command Prompt, or any Git client).
      2. Use the cd command to navigate to the directory where you want to clone the repository.
      3. Run the following command to clone the repository: git clone link_provided
  - Open the Excel file
  - Ensure macros are enabled.
  - To execute the script, go to the VBA editor (Alt + F11), find module1 containing the script, and run Multiple_year_stock_data from the macro menu. Or Can execute the script by pressing the "Calculating the Tables" button (found in "Q1" excel sheet) and can reset the results by pressing "Reseting the Tables" button (found in "Q1" excel sheet). 

# Usage Instructions
- The script will automatically process all worksheets.
- Once the script is executed, results will be printed in new columns, including:
   - Ticker symbols.
   - Quarterly changes and percentage changes.
   - Total stock volumes.
   - Greatest percenteage increase, Greatest percenteage decrease, and Greatest total volume.
