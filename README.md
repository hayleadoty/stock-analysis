# All Stocks Analysis - 2017 & 2018
## Overview of Project

### Purpose

The purpose of this project was to refactor existing code in order to run an analysis of all stocks’ total daily volume and return for 2017 and 2018. We were looking to discover whether or not refactoring the code made the run time more efficient, as we eventually want to expand the dataset to include the entire stock market.

### Results

For 2017, DQ had the lowest daily volume (35,796,200 shares) but the highest return percentage (199.4%). TERP was the only stock with a return loss of -7.2%. The stock with the highest daily volume was SPWR (782,187,000 shares), with a return of 23.1%.

![2017 Results](/Resources/2017_Results.png)

The refactored code did not run faster than the previous code in my case. The previous code ran in 0.52 seconds, and the refactored code ran in 0.59 seconds:

![2017 Refactored Code Run Time](/Resources/VBA_Challenge_2017.png)

For 2018, AY had the lowest daily volume (83,079,900 shares). All but two stocks showed a negative return, with the lowest being SPWR (-44.6%). The stock with the highest return was RUN at 84.0%. The stock with the highest daily volume was ENPH, at 607,473,500 shares.

![2018 Results](/Resources/2018_Results.png)

The refactored code did not run faster than the previous code for this year, either. The previous code ran in 0.53 seconds, and the refactored code ran in 0.60 seconds:

![2018 Refactored Code Run Time](/Resources/VBA_Challenge_2018.png)

We adjusted our code to create a ticker index:
  ```
  tickerIndex = tickers(i)
  ```
And three output arrays:
   ```
   Dim tickerVolumes As Long
   Dim tickerStartingPrices As Single
   Dim tickerEndingPrices As Single
   ```
   
This way, we can add one or more new stocks to our dataset. We can then use our refactored code to run an analysis on any number of stocks, all the way up to the entire stock market.

### Summary

In general, refactoring code allows us to save time by reusing working code in more streamlined or simplified way. We can also make existing code run faster through refactoring. A negative of refactoring code could be that sometimes you are working with code that was written by someone else, so it could be confusing or contain bugs.
