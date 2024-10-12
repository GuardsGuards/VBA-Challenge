VBA Challenge - Stock Market Data Analysis
Overview
This project is part of Module 2 Challenge, where we use VBA scripting to analyze stock market data. The script loops through stock data for each quarter and outputs key information such as quarterly changes, percentage changes, and total stock volume. Additionally, it identifies the stock with the greatest percentage increase, the greatest percentage decrease, and the greatest total volume.

Features
Ticker Symbol: Extracted from stock data for each quarter.
Quarterly Change: Calculated as the difference between the closing price at the end of the quarter and the opening price at the beginning of the quarter.
Percentage Change: Calculated as the percentage difference between the closing price and the opening price.
Total Stock Volume: Sum of all stock volumes for the quarter.
Greatest % Increase/Decrease: Identifies the stocks with the greatest percentage increase and decrease.
Greatest Total Volume: Identifies the stock with the highest total volume.
Conditional Formatting: Highlights positive changes in green and negative changes in red.
Multi-Sheet Functionality: The script runs on all worksheets, allowing for analysis across multiple quarters.
Files
VBA Script: The script file (VBA_Challenge.vba) contains all the logic to perform the analysis.
Screenshots: Example screenshots showing the output of the script.
alphabetical_testing.xlsx: A smaller dataset for testing the VBA script during development.
How to Use
Download the Dataset: Place the dataset in the same folder as the script.
Run the Script: Open the Excel file and run the script using the Developer tab in Excel. The script will loop through each quarter and display the results on the corresponding sheet.
Review the Output: The results, including ticker symbols, quarterly changes, percentage changes, and total stock volume, will be displayed. Conditional formatting will highlight positive and negative changes.
Key VBA Functions
Looping Across Worksheets: The script automatically loops through all the worksheets to process the data from each quarter.
Conditional Formatting: Applied to the percentage and quarterly change columns to indicate performance.
Tracking Greatest Values: The script keeps track of the stock with the greatest percentage increase, decrease, and total volume.
Results
The script generates the following key results for each quarter:

List of ticker symbols and their respective quarterly and percentage changes.
Total volume of stocks for each ticker.
Greatest percentage increase and decrease for the quarter.
The stock with the greatest total volume.
