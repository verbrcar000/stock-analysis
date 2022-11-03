# Stock Analysis

## Overview
The purpose of this analysis is to help Steve's parents decide which stocks would be best for them to invest in. It calculates the total daily
volume for each stock as well as the return, first, making it clear which stocks did well and which did not. It also makes it easy to look at both
of these variables for the year 2017 or for the year 2018, allowing the Steve to compare the two.

## Results
As seen in the images below (2017 first, 2018 second), almost all stocks did signifcantly better in 2017 as opposed to 2018. On average, stocks in 2017 produced a much higher return and higher total volumes. In 2018 the majority of stocks actually produced a negative return while in 2017, the majority produced a positive return. One stock however, actually did signifcantly better in 2018. RUN went from a 5.5% rate of return in 2017 to 84% in 2018.

![2017 Results](Results_2017.png) ![2018 Results](Results_2018.png)

As for exection times of the actual scripts, we can see clearly below that the original script (shown first) ran much more slowly than the refactored script (shown second).

![Org Script Runtime 2017](Green_Stocks_2017) ![Org Script Runtime 2018](Green_Stocks_2018)

![Refactored Script Runtime 2017](VBA_Challenge_2017) ![Refactored Script Runtime 2018](VBA_Challenge_2018)

## Summary 
Refactoring code can make the code run much more quickly and efficiently. It can also simplify the code, making it easier to understand or to debug. Two disadvantages to refactoring code might be that it can be time consuming and that it could introduce new bugs into the code. In the case of this analysis of stocks, I was able to refactor the code so that instead of looping through each of the 32,000+ data points 12 different times for each ticker, it just looped through each data point once. The script ended up running 5 times faster after being refactored. If there were even more data points to run through, this difference could have been even more significant so this proves the usefulness of refactoring.

### Note
Refactored script found in VBA Module 2.
