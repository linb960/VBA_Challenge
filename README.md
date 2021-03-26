# VBA Refactoring Challenge
## Overview of Project
### Purpose
The purpose of this project is to gain insight into how refactoring code can affect the run time of the code particularly when there are many of lines of data to analyze.  
### Background
The challenge is to analyze the yearly data of twelve different stocks and output, for each stock, its ticker symbol, total volume and return for that year.  

The columns of the input sheet represent the information for each stock.  They are Ticker, Date, Open, High, Low, Close, Adj Close, Volume.    The rows of data represents each day of the year the stock market was open.  For our analysis we are looking at data from 2017 and 2018. 

An initial project was used as the basis for this challenge project.  In the initial project the code used two for loops, one inside another.  The inner loop checked if the stock ticker was changing to the next ticker, or if it had just changed to a new ticker.  This made it possible to add up the total volume and get the starting and ending prices. The outer loop iterated over the 12 ticker symbols.  So as a ticker symbol changed the inner loop restarted at the top of the sheet searching for the ticker symbol and determining if it was at the start date or end date for that ticker for that year.  In the outer loop, prior to going to the next ticker the information collected was output to the All Stocks Analysis sheet.  The ticker symbol, total volume and the return which was caluculate by dividing the starting price from the ending price and subtracting 1 make up the columns of the output sheet.

This challenge was designed to see if refactoring the initial projects code and use one loop instead of two would make a difference in the run time.  To do this arrays are used for total volume, starting price and ending price. This makes it possible to run one loop from the top of the sheet to the bottom one time.  Each array uses a variable called tickerIndex to keep track of which ticker symbol the code is collecting data for each array.  The code to check if the next ticker symbol is different than the current ticker symbol is still incoporated in the refactored code but instead of having to drop out of the loop to increment to the next symbol the tickerIndex is incremented by 1.


# Results
