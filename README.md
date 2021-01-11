# stocks-analysis - Challenge 2

## Project Overview

This project is created to analyze the data for a given stock market. This will help the end user to figure out what stock to invest in.

## Result Summary

### Results from 2017 analysis:

The image below shows the results from the 2017 stock data analysis

![2017 Stock Analysis](/Resources/VBA_Challenge_2017.PNG)

This image also shows the amount of time it took to execute the code.

Based on the above image, the user should invest in Ticker DQ as it is giving the highest return in 2017.

Ticker TERP had a negative return and hence the user should not invest in this stock.

### Results from 2018 analysis:

The image below shows the results from the 2018 stock data analysis

![2018 Stock Analysis](/Resources/VBA_Challenge_2018.PNG)

This image also shows the amount of time it took to execute the code.

Based on the above image, the user should be very careful as most of the stocks had a negative return.

Two of the stocks had a high positive return and the user should focus on those two (ENPH and RUN).

Based on the above images it shows that the code took much faster to execute as compared to the original script.

The original code has 2 for loops, one of them is nested into another. This will take the code much longer to run as the nested for loop runs through the entire sheet everytime the code moves to the next ticker.

In the refactored code, it only loops through the entire sheet once. Each time a new ticker is identified the code increases the ticker index by 1.

The following lines of code allows the refactored code to use only one for loop: 

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            tickerIndex = tickerIndex + 1
            
        End If

## Summary of refactoring code
The advantage of refactoring code is to create time efficiency. In our example, by refactoring code we reduced the code run time. In this example it is not very noticeable, however when you are dealing with large data, it is always a good idea to refactor code. 

It also allows many users to follow the code and be able to understand the code better. Refactoring code can reduce the number of lines of code needed. As an example to achieve a specific result you can use 15 lines of code, however the same can also be achieved with 8 lines of code.

The biggest disadvantage of refactoring code is that it may introduce bugs and also alter the end result that the user is looking for.

In our case it was very beneficial to refactor the code as the code does not need to loop through the entire sheet so many times. 

The disadvantage in our case is that the refactored code will only work and produce the correct result is if the data is sorted by the ticker. In the event the data was not sorted by the ticker, the refactored code will fail and not produce the correct results.

