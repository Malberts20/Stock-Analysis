# Stock-Analysis


## Overview of Project

Two year perfomance data (years 2017-2018) was analyzed for twelve stocks to evaluate potential investment strategy.  VBA code was utilized to calculate trading volume and returns on an annualized basis for each year.   

Subsequent coding exercise focused on refactoring VBA code to loop through the dataset one time as opposed to looping through each individual stock.  Code time was measured to determine if refactoring resulted in more efficient programming.

## Results

Original VBA Code (green_stocks_final.xlsm): This code calculates annualized volume and returns and creates output for each stock individually.  Code time measured at approximately 1.5 seconds for each year.

Refactored VBA Code (VBA_Challenge.xlsm) : Code refactored to calculate volume and returns for all stocks in one loop and susequently sent to output.  Code time measured at less than 0.35 seconds for each year.

Stock performance for 2017 and 2018 (both codes yield same results) given below:


![2017_Analysis](https://user-images.githubusercontent.com/71353552/95018749-dfe45200-061e-11eb-9491-ef8b378278b9.PNG)
![2018_Analysis](https://user-images.githubusercontent.com/71353552/95018828-53865f00-061f-11eb-82f2-a2867abe0096.PNG)


Elapsed run time for the original (left) and refactored code (right) for 2017 and 2018 given below:


![ElapsedTime_Original_2017](https://user-images.githubusercontent.com/71353552/95018927-c2fc4e80-061f-11eb-9b4f-8760adc46ab2.PNG)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/71353552/95018928-c394e500-061f-11eb-962c-91507a2db5ce.PNG)

![ElapsedTime_Original_2018](https://user-images.githubusercontent.com/71353552/95018937-cee81080-061f-11eb-961b-0dbe72ca2c4f.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/71353552/95018938-cf80a700-061f-11eb-8721-5f27a82cc3d4.PNG)


## Summary

From stock analysis results (shown above), only two (2) stocks, ENPH and RUN, show positive returns for both years and would be the recommended stocks for potential investment. 

Refactored code improved programming efficiency as run times reduced from approximately 1.5 sec to under 0.35 sec for both years.

The advantages of refactored code results in a program that is more organized and is easier to understand. However, it should be noted that refactoring has diminishing returns and can be costly in terms of both time and money.  It is important that the appropriate time is committed to the requested stock analysis and resulting recommendations as opposed to significant efforts dedicated to code optimization.

