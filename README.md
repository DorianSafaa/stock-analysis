# Green Stocks Analysis with VBA

## Overview of Project

The goal of the project is to analyze some stock data from 2017 and 2018. We want to find the total daily volume and yearly return for each stock. The total daily volume represents the total number of shares traded daily so we can know how actively a stock was traded. And the yearly return represents the percentage difference in the price of the stock at the beginning and end of the year. 

## Purpose

Our purpose in this project is to find the best stock to invest in based on the stock analysis of 2017 and 2018. We already created a code in VBA that helped us find the total daily volume and return for each stock, but we want to refactor the code so it can take less time to execute in case of analyzing more stocks. The refactored code will loop through the data one time.  

## Result

To find the total daily volume and return for every ticker, we created an index that can access the tickers list and find the total of the daily volume of a specific ticker by looping over the volume column and adding all the volumes. 
Also, we used conditional if statement to find the starting price and end price of a specified ticker and used the formula (end price)/(start price)-1 to calculate the return of each ticker.

Comparing the stocks' returns in 2017 and 2018, we can see that most of the stocks had a negative return in 2018 except ENPH and RUN. These two stocks had positive returns in both years, but for RUN the return was more by 80% in 2018, and for ENPH the return in 2018 was less by almost 50% than in 2017. 

 ![2017 Stocks](https://user-images.githubusercontent.com/66279829/154908870-4bec0ef9-436a-492f-a0ed-05c86be5d1e1.PNG)

![2018 Stocks](https://user-images.githubusercontent.com/66279829/154908932-9b68c29b-8076-42f5-a783-7815d38b8fc8.PNG)


In addition to the analysis, we calculated how long the code takes to execute and we output the elapsed time in a message box. In the original code, the time to run the code was more by 1 sec than in the refactored code. So, we can say that the refactored code decreased the run time. 

 ![VBA_Challenge_2017](https://user-images.githubusercontent.com/66279829/154908972-4dc72665-c119-43ba-99b1-0ec8e5b8566e.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/66279829/154909016-9fbed5de-e81f-4c5e-a9ed-9c09c1cb78cd.png)

## Summary
In general, refactoring leads to better quality code. Refactored code executes the program faster since we limit the repetition and the use of loops. Also, the code will be more organized and easier to follow if anyone else wants to understand the steps. 

However, sometimes refactoring a code can be time-consuming and can create more problems than with the original code.


In this case, we can confirm from the elapsed time that the refactored code helped in running the program faster than the original code. Creating an index to access the tickers, volume, and price lists helped us to loop through the data one time only. This is especially true since our lists are of the same length. If the lists had different lengths, we may need to adjust the code or use more than one index to access them. And this can be a disadvantage since it cannot be used for all scenarios. 



