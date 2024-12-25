# VBA Stock Analysis Challenge

## Statement of work and sources
This is Joe Portnoy's submission for the VBA Challenge. All code was authored by Joe Portnoy with debugging and formatting assistance from the ChatGPT models 4o and o1.

## Files Included
- Multiple_year_stock_data.xlsm
    - Original Data Source on which to run the script. Macro embedded and called "Stock Analysis."
- Multiple_year_stock_data_SOLVED.xlsm
    - Script solution applied and embedded into file
- Multiple_year_stock_analysis.vba
    - VBA challenge solution annotated

## Folders Included
- images
    - Q1.png
    - Q2.png
    - Q3.png
    - Q4.png
- Starter Code
    - alphabetical_testing.xlsx
        - Small dataset for testing
    - Multiple_year_stock_data.xlsx
        - Original dataset

## Requirements and Explanation of Solution

### Retrieval of Data

The script loops through thousands of lines of code to retrieve unique ticker names, finds the open price and the close price of unique ticker names in each quarter, it calculates the change over the quarter for each unique ticker, calculates the percent change and finally adds the total volume of stocks for each day of trading.

### Column Creation

For each retrieval and calculation, the script creates new column headers and adds the calculated values into the rows below.

### Conditional Formatting
Conditional formatting is applied to the Quarterly Change column when a ticker has a positive calculation, the cell turns green, when a ticker has a negative calculation, the cell turns red. Finally, when there is zero change, no formatting is applied.