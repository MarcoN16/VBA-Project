# Multiple Year Stock Data Analysis

This Excel file contains stock data spanning multiple years. I've developed a macro that streamlines the process of separating and analyzing the dataset for each spreadsheet(year) with a single click.
To start the macro press the button on the first page, or run it from the developer tab

# Instructions
The script is divided into two sections:

# Creation of List:
1 - Separation of Tickers: 
  This step categorizes each ticker, representing a specific stock, and creating a separate list. 
2 - Yearly change:  
  It determines the difference between the opening price at the beginning of a given year and the closing price at the end   of that year.
3 - Percentage Change: 
  it computes the percentage change from the opening price at the start of a year to the closing price at the end of that year.
4 - Total Stock Volume: 
  it calculates the total stock volume of each stock.

# Analysis of Dataset:
1 - Conditional Formatting: 
  The script performs conditional formatting to highlight positive changes in green and negative changes in red within the yearly change column.
2 - Identification of Extremes: 
  It identifies both the maximum and minimum values in the percentage change column, reporting the respective stock tickers and their associated values.
3 - Highest Total Stock Volume: 
  the script identifies and reports the stock ticker and value of the highest total stock volume for each year.

# Note
Upon running the script, a message box will appear, prompting you to confirm whether you'd like to proceed with the analysis:
- Click "Yes" 
to perform an in-depth analysis of the dataset.
- Click "No" 
to generate only a list of the stock data without performing additional analysis.

The analysis is set for the period from January 2nd to December 31st
