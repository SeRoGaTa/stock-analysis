# Stock-analysis with Excel VBA Macro 

***Version 1.0.0***

---

## Overview of Project
#### This project analyzes given data from 11 stocks and their closing prices and volumes among other characteristics so Steve's parents can invest in a well and solid stock.

## Purpose
#### This analysis will help Steve to analyze each stock return over 2017 and 2018 as well as their volumes, it will graphically tell which stocks has a positive return over that period and which one does not, with this Steve parents will have strong data to invest safely their savings in stocks.

---

## Results
### As part of the results from the analysis, there is one strong stock which remains in a positive return after those two periods called Enphase Energy Inc (NASDAQ "ENPH"). On their first year it had a return of 129% and in 2018 it had 81%, this means is a strong stock to invest.
### In this image you can see the result table for 2017 stock analysis, here you can see the 129% of ENPH.

<img src="https://github.com/SeRoGaTa/stock-analysis/blob/main/Challenge%202/Resources/Stocks_table_2017.png" width="400">

### In this second image you can see the result table for 2018 stock analysis, where we see ENPH ticker with a positive 81%.

<img src="https://github.com/SeRoGaTa/stock-analysis/blob/main/Challenge%202/Resources/Stocks_table_2018.png" width="400">

### Code Structure and refactored analysis
#### During the analysis phase we tested two types of codes, the first one was accurate and easy to write but it took longer to execute and the second one (refactored) were way faster that the first one.
####
#### Here you can find an example of the initial code, this code use For loops and if statements with the use of only one array for the tickers name.

<img src="https://github.com/SeRoGaTa/stock-analysis/blob/main/Challenge%202/Resources/First%20code.png" width="800">

#### This is an example of the second code, the main difference between both is that in this case, instead of review the entire set for each result row and printing it, we use 4 arrays to store each value once we review each line, so it just need to go thru the entire set once improving in a huge way the execution time of the code.

<img src="https://github.com/SeRoGaTa/stock-analysis/blob/main/Challenge%202/Resources/Second%20code.png" width="800">

#### In the following images you can see the time to execute difference between codes

2017 First code attempt

<img src="https://github.com/SeRoGaTa/stock-analysis/blob/main/Challenge%202/Resources/green_stocks_2017.png" width="300">
2018 First code attempt

<img src="https://github.com/SeRoGaTa/stock-analysis/blob/main/Challenge%202/Resources/green_stocks_2018.png" width="300">
2017 Second code attempt

<img src="https://github.com/SeRoGaTa/stock-analysis/blob/main/Challenge%202/Resources/VBA_Challenge_2017.png" width="300">
2018 Second code attempt

<img src="https://github.com/SeRoGaTa/stock-analysis/blob/main/Challenge%202/Resources/VBA_Challenge_2018.png" width="300">

---

## Summary
### As part of the refactoring process I saw an improvement of the code and that improvement encourage me to keep doing further changes to my original code. Next, I'll describe the pros and cons of the refactored code and compare it with the original.
###
* What are the advantages or disadvantages of refactoring code?
	* After refactoring my original code, my surprise and first advantage was the time to execute the code, it was in average 6 times faster or what is the same 16% of the original code execution time. This will help to have deeper analysis or increase the number of stocks to analyze. Other advantage I saw is that is always a good practice to continues improvement our work so we can find faster and easier ways to have the same result.
	* As a disadvantages, I see that this may end in a perfectionism and it could take you more time to deliver the work, we need to find the line between get it better and deliver and get it perfect and delay to deliver your work. Other disadvantage is that you could do your code as complex as is difficult to debug for future users and it might end rewriting the complete code.

* How do these pros and cons apply to refactoring the original VBA script?
	* First, I'd like to say, refactoring is a good practice from my standpoint of view. I release that we can improve our codes and make them faster so I see as a pro to have the ability to have bigger datasets that we can test due to our faster code.
	* And as a cons, I think we could lose our selves in a loop of getting better and better codes that we never delivers or feels like we aren't delivering a good result. So, at the end we need to have a hard stop and look back and decide if we want to refactor again or not.

You can locate the complete analysis file on the following link [2017-2018 Stock Analysis](https://github.com/SeRoGaTa/stock-analysis/blob/main/Challenge%202/VBA_Challenge.xlsm) and all image used in the following folder [Resources](https://github.com/SeRoGaTa/stock-analysis/tree/main/Challenge%202/Resources)
