# Refactoring All Stock Analysis VBA Code

## Overview of Project

### Purpose
The purpose of the project was to assist Steve determine what stock was worth investing in by finding the total daily volume and yearly return for each stock. We refactored our “All Stock Analysis” VBA code to loop through the entire stock market data and determine if refactoring the VBA code would yield a faster output. 

## Results 

### Analysis 
Below are the procedures I took to refactor our All Stock Analysis VBA Code.

1. Created ticker index and three output arrays for "tickerVolumes", "TickerStartingPrices" and "TickerEndingPrices". TickerVolumes were considered a long data type while TickerStartingPrices and TickerEndingPrices were considered to be single data types. 
<img width="269" alt="Refactored Code - Output Arrays " src="https://user-images.githubusercontent.com/91925639/138619828-b0feb975-331f-4dbe-b0c2-1ac70be78efc.png">
2. Created a for-loop to initilaize the tickerVolumes to zero and created a for loop to loop over all the rows in the 2017 and 2018  data. 
<img width="408" alt="Refactored Code - Ticker Volume " src="https://user-images.githubusercontent.com/91925639/138619869-9f9ba61d-6b4f-4cf0-bcfb-f8b36b26bbfb.png">
3. Inside the for-loop, increased the volume for the current ticker. Then created an if-then statement to check if the current row is the first row in the selected tickerIndex. If the statement is true then the tickerStartingPrices and tickerEndingPrices would be assigned. Lastly, increased the tickerIndex if the next row's ticker did not match the previous row's ticker.
<img width="517" alt="Refactored Code - Ticker Index" src="https://user-images.githubusercontent.com/91925639/138619897-1846110e-3e1f-4337-a57d-bba5623cd8fc.png">
4. Created a for-loop to loop through the three arrays (tickerVolumes, tickerStartingPrices and tickerEndingPrices) to return outputs for "Ticker", "Total Daily Volume", and "Return".
<img width="535" alt="Refactored Code - Loop through Arrays " src="https://user-images.githubusercontent.com/91925639/138619926-2a788185-9e59-4738-8013-631bf9dcf289.png">
5. Formatting for our final All Stock Analysis for 2017 and 2018 data. 
<img width="604" alt="Refactored Code - Formatting" src="https://user-images.githubusercontent.com/91925639/138619951-517e6622-7010-441c-ad97-bfd8c2510ca0.png">
     
### All Stock Analysis Output for 2017
Ticker "TERP" was the only stock with a negative return in 2017.

![Stock Analysis Output - 2017](https://user-images.githubusercontent.com/91925639/138618321-9fb79d67-7ecb-4601-9deb-2f56e712cbee.png)

### All Stock Analysis Output for 2018
Tickers "ENPH" and "RUN" were the only stocks with positive returns in 2018.

![Stock Analysis OUtput - 2018](https://user-images.githubusercontent.com/91925639/138618377-0b78cf53-1b13-44b3-a49a-62bbafeeec8f.png)

### Time Eplapsed for the VBA Code - 2017
![VBA_Challenge_2017](https://user-images.githubusercontent.com/91925639/138574339-f2825ddb-373e-4961-804b-77c6af0290fd.png)
### Time Eplapsed for the VBA Code - 2018
![VBA_Challenge_2018](https://user-images.githubusercontent.com/91925639/138574353-78502403-48f5-467c-9f0a-04e74dfe5d5f.png)

## Summary

### Advantages & Disadvantages of Refactoring Code 

### Advantages & Disadvantages of Original vs Refactored VBA Script 

