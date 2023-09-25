# VBA-challenge
Multiple Year Stock Analysis:

Objective: Draw Yearly Summarized information from daily transactional data stock tickers.

VBA Script Process Flow:

	Analyze Raw Stock Data Formats. 
•	Read Each Work sheets for year after year Stock Market Transactional Data for each Ticker.
•	Clear contents from all Summary Cells to avoid Garbage data in Cells.
•	Create Headers for Summary information (Ticker, Yearly Price Change, Percentage Change and Volume) to be displayed after Macro Run.
•	Create Headers for Greatest % Increase, Greatest % Decrease and Greatest Volume.
•	Sort Raw data based on Ticker and Date in Ascending Order
•	Get First Day in Year for Market Open from Date Column.
•	Get Last Day of Year for Market Open from Date Column.
•	Read Each Row of Sheet excluding Headers with reference to Work Sheet Object of Workbook in Excel.

	With Reference to each Row:
•	Get Ticker, first Open price for ticker, last of year close price. Summarize Volume for each ticker for entire year.
•	Calculate Yearly change in Price and Compute percent change in price for each ticker.
•	Conditional Format Negative and Positive change in Price with Red and Green color. 
•	Get Greatest % Increase, Greatest % Decrease and Greatest Volume among tickers.


