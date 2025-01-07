# VBA Stock Analysis Challenge

This is Joe Portnoy's submission for the VBA Challenge. All code was authored by Joe Portnoy with assistance on debugging and formatting from ChatGPT.

## Retrieval of Data

The script loops through thousands of rows in the worksheet to retrieve unique ticker names, finds the open price and the close price of unique ticker names in each quarter, it calculates the change over the quarter for each unique ticker, calculates the percent change and finally adds the total volume of stocks for each day of trading.

## Column Creation

For each retrieval and calculation, the script creates new column headers and adds the calculated values into the rows below.

## Conditional Formatting
Conditional formatting is applied to the Quarterly Change column when a ticker has a positive calculation, the cell turns green, when a ticker has a negative calculation, the cell turns red. Finally, when there is zero change, no formatting is applied.

## Calculated Values
The script will also create two new columns labeled Ticker and Value in columns O and P to pull ticker for the Greatest values for % Increase, % Decrease and Total Volume. Labels are placed in rows 2-4.

## Looping Across Worksheets
Finally, the script will loop across all worksheets and perform the actions above to analyze all rows of data.