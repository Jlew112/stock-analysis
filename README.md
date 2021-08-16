# stock-analysis

## Overview
We have built a macro to cycle through a list of tickers and sell prices in order to help Steve's parents figure out which stock to invest in. We've compiled data on 12 different stock options over the course of 2017 and 2018, and I think we've been able to make a recomendation.

## Breaking down the data with Results
There were 12 companies with 12 tickers. We are able to cycle through the list and show which of these were successful in 2018 and 2017 based on their total price change over the course of the year, and the total volume of sales made at each ticker. 2017 found great returns fo the group. with all but 2 showing overal growth in the price. 2018 was a different story. Almost the entire field brought losses between -3.5% and -62.6%. There was a bright spot with the Tickers, ENPH and RUN. Both posting +80% increases and high volumes of trades.
![alt text](https://github.com/Jlew112/stock-analysis/blob/main/Resources/2017.PNG)
![alt text](https://github.com/Jlew112/stock-analysis/blob/main/Resources/2018.PNG)

## Summary
The refactored code worked well. The advantage being that we didn't have to start from complete scratch, cutting down the actual time needed to write the code for this piece of the project. However, the slight changes in variable names, and the addition of the arrays did complicate the issue. Debugging was difficult, but necessary as we were able to finally get the macro to work.
While the original code did work, it was limited because it didn't take into account the total list of data. And if more data was going to be entered, there are limitations to capturing the new data. Overall, I can see the value in VBA scripting, and I'm excited about the possibilities it has brought to light.
