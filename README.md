# Green Stocks Analysis Using VBA & Excel

## Overview

### Purpose

The purpose of this exercize was to take existing code from our stock analysis and refactor it for functionality and efficiency purposes. After doing this, we can determine whether refactoring the code made the script run faster, by implementing a timer function. The end goal is for Steve to utilize the same dataset you made for him to run analysis on every stock, not just the 12 we previously used.

### Results

First, we set 'tickerIndex' to zero before looping it over the dataset rows. Then, four arrays were made for 'tickers', 'tickerVolumes', 'tickerStartingPrices', and 'tickerEndingPrices'. Then, we use 'tickerIndex' to access the stock ticker index for the arrays we just made. Next, a for loop it utilized to set 'tickerVolumes' to zero, and if the ticker doesn't match, 'tickerIndex' is increased. Another loop is used to loop over all the rows in the spreadsheet. A time pop up message is utilized to give the speed of each year's analysis, as seen below.

### 2017 Results

![VBA_Challenge_2017](https://user-images.githubusercontent.com/77767984/117514961-b5fb6100-af5a-11eb-90ca-398acf205cf1.png)


### 2018 Results

![VBA_Challenge_2018](https://user-images.githubusercontent.com/77767984/117514965-b98ee800-af5a-11eb-848a-c77c412aa2b1.png)


### Summary

One obvious advantage of code refactoring is the opportunity to "clean house", or tidy up some messy or confusing code. Another is increased efficiency, as well as readable code, which will make it easier for others who didn't originally write the code to be able to dive in and contribute. Disadvantages could include the extra time spent to refactor code you already spent time to write.
