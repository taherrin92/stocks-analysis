# Overview
- Steve wants to help his parents boost their retirement fund by figuring out which stocks have a positive annual return and requested an analysis of the stocks that he uploaded into Excel. To help him, I wrote a macro in VBA that calculates the trading volume and annual return of a stock for the tickers that Steve wants to analyze. The first version of the macro worked for his data. Clicking a button and typing in the year automatically calculated the trading volume and annual return, but I could visually tell that the macro was inefficient. 
- I refactored the existing code to run much more efficiently by only running through the data worksheet once. This refactored version can also handle many more tickers that Steve could analyze in one instance.

## Results 
- As mentioned above, the first version of the macro could run through the data that Steve provided, however, I could see the code visibly changing worksheets many times during its run time. This was one flag that the first version was inefficient. The reason this was happening was because the nested for-loop in the first version would run through all of the data and calculate the values for a single ticker, change worksheets and record the value before looping again for the next ticker seen here: 
![first-version code](https://github.com/taherrin92/stocks-analysis/blob/Final-Version/Code%20screenshots/First-Version-Code.png)


This meant that the code would have to run for over half a second to calculate 24 values for 3300 rows of data. 

![first-version 2017 run](https://github.com/taherrin92/stocks-analysis/blob/Final-Version/First-Version-Resources/Unrefactored_2017.png) ![first-version 2018 run](https://github.com/taherrin92/stocks-analysis/blob/Final-Version/First-Version-Resources/Unrefactored_2018.png)

- The refactored code required more lines and was more complicated to write, but it significantly shortened the time it took to calculate and record all of the data. This is because there are no nested loops and the code was set up to only have to run through the data once as seen here:

![refactored-code](https://github.com/taherrin92/stocks-analysis/blob/Final-Version/Code%20screenshots/Refactored-Code.png)
The refactored code assigns all of the values of the first version code into arrays which stores all of the calculated values as the code progresses. When the code runs to the last row containing data for that array, then the code continues to the next ticker using the code `tickerIndex = (tickerIndex + 1)` without having to start over and loop through the data again. By separating this loop from the nested loop used in the first version's code and the loop that recorded all of the data from the arrays, the worksheet only needed to be changed over one time to record all of the data in the arrays.

- This refactored code cut the run time down to about one-tenth of a second as seen here:

![final-version 2017 run](https://github.com/taherrin92/stocks-analysis/blob/Final-Version/Resources/VBA_Challenge_2017.PNG) ![final-version 2018 run](https://github.com/taherrin92/stocks-analysis/blob/Final-Version/Resources/VBA_Challenge_2018.png)


## Summary
- The advantage to refactoring code is that it takes an idea that's already been written out and tweaks it to enhance its performance. In general, this could mean cutting out rundant lines and compacting the code into some thing clear to read and runs fast. The disadvantage to refactoring code is that if the idea of the first version is rudimentary or runs based on luck, it can be very difficult to refactor and might even have to be completely redone to achieve an efficient macro. 

- The pros of refactoring the code in VBA are that the run times are significantly shortened, which allows the macro to handle a much larger volume of data than what Steve provided. The cons in this case was that meant taking a simple nested loop that was only a few lines of code as seen in the first version image, and adding in many more lines and variables that led to a few small mistakes (in my case, typing in `tickersEndingPrices(tickerIndex)` when the variable was `tickerEndingPrices(tickerIndex)`) that took a long time to find and even led to a couple of misleading error messages. 
