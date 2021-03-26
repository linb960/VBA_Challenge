# VBA Refactoring Challenge
## Overview of Project
### Purpose
The purpose of this project is to gain insight into how refactoring code can affect the run time of code particularly when there are many of lines of data to analyze.  
### Background
The challenge uses this <a href="VBA_Challenge.xlsm">dataset</a> to analyze the yearly data of twelve different stocks and output, for each stock, its ticker symbol, total volume and return for that year.  Because this dataset only uses 12 stocks the idea behind this project is to determine if it's possible to reduce the runtime so that if 100's or thousands of stocks were analyzed the execution time is shortened.

The columns of the input sheet represent the information for each stock.  They are **Ticker, Date, Open, High, Low, Close, Adj Close, Volume**. The rows of data represents each day of the year the stock market was open.  For our analysis we are looking at data from **2017 and 2018**. 

An initial project called Green Stocks was used as the basis for this challenge project.  In the initial project the code used *two for loops*, one inside another.  The inner loop checked if the stock ticker was changing to the next ticker, or if it had just changed to a new ticker.  This made it possible to add up the total volume and get the starting and ending prices. The outer loop iterated over the 12 ticker symbols.  So as a ticker symbol changed the inner loop restarted at the top of the sheet searching for the ticker symbol and determining if it was at the start date or end date for that ticker for that year.  In the outer loop, prior to going to the next ticker the information collected was output to the **All Stocks Analysis** sheet.  The **Ticker, Total Daily Volume and Return,** which was caluculate by dividing the starting price from the ending price and subtracting 1, make up the columns of the output sheet.

This challenge was designed to see if refactoring the initial projects code and use *one for loop* instead of two would make a difference in the run time.  To do this arrays are used for total volume, starting price and ending price. This makes it possible to run one loop from the top of the sheet to the bottom one time.  Each array uses a variable called tickerIndex to keep track of which ticker symbol the code is collecting data for.  The code to check if the next ticker symbol is different than the current ticker symbol is still incoporated in the refactored code but instead of having to drop out of the loop to increment to the next symbol the tickerIndex is incremented by 1.

# Results
The results of the output sheet, that is the calculations for volume and return for the challange is the same as the initial project and is expected.  The real results that we are interested in for this challenge is the output times.

### Results for 2017 sheet
This is the output sheet for the data from the 2017 sheet:
  <img src="Resources/VBA_Challenge_Output_2017.png" width="675" height="350"><br>

And the output times for 2017 for both the refactored code in the VBA Challenge and the Green Stocks code with the two loops:

|VBA Challenge Output Time for 2017|Green Stocks Output Time for 2017|
|---|---|
|<img src="Resources/VBA_Challenge_2017.png" width="250" height="200"><br>|<img src="Resources/green_stocks_2017.png" width="250" height="200"><br>|

### Results for 2018 sheet
This is the output sheet for the data from the 2018 sheet:

  <img src="Resources/VBA_Challenge_Output_2018.png" width="675" height="350"><br>

And the output times for 2018 for both the refactored code in the VBA Challenge and the Green Stocks code with the two loops:

|VBA Challenge Output Time for 2018|Green Stocks Output Time for 2018|
|---|---|
|<img src="Resources/VBA_Challenge_2018.png" width="250" height="200"><br>|<img src="Resources/green_stocks_2018.png" width="250" height="200"><br>|

### Analysis of Data
As seen above the refactored code does run faster then the original code.  While the actual data collected and shown on the **All Stocks Analysis** sheet is different for 2017 and 2018 the time of execution was quite similar.  

For 2017 the refactored code ran in .078125 seconds and the original Green Stocks code with two loops ran in .2890625 seconds.  A difference of .2109375 seconds.  

Similarly the 2018 sheet finished in .078125 seconds.  Where the original Green Stocks code for 2018 ran in slightly less time then 2017 at .2851562.

The difference of about .2 seconds between the original and refactored VBA code shows therefore that by running one loop vs. two loops did improve the execution time.

## Summary
The point of this challenge was to determine if refactoring code can help improve the runtime or execution time of the code.  Lets consider a few of the advantages and disadvantages for doing this in general and in the context of this challenge.

### Advantages in General
The advantages of refactoring code can be best understood through <a href="https://en.wikipedia.org/wiki/Big_O_notation">Big O notation.</a> Big O notation is used to indicate the approximate time it will take to 
