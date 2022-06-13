# Stock-Analysis

## Overview of Project

The purpose of this analysis was to help Steve be able to quickly analyze and identify what stocks he should advise his parents to invest in. In order to do this, we reviewed two years worth of data (2017 & 2018) and created calculations to identify the total volume of trades, as well as the annual return. The higher the percentage of the annual return, the more likely it is that it would be a good stock to invest in. If a stock has a negative annual return, he can quickly decide on whether or not it is worth the risk. 

In order to allow Steve to run the code and view the results with minimum effort, we enabled conditional formatting and auto-formatted the column widths and text. 

## Results

### Refactored Code vs Original Code
By refactoring the new code, we were able to avoid the multiple nested loops that were used in the original document. By using the tickerIndex, we were able to gather all of the values for each ticker, seamlessly. The execution time decreased by over 10% from the original code's execution time. 

![Challenge_2017_RunTime!](/Resources/VBA_Challenge_2017.png)

![Challenge_2018_RunTime!](/Resources/VBA_Challenge_2018.png)

### Stock Performance Between 2017 & 2018
A quick analysis of the results indicates that 2017 was a much stronger year for these stocks than 2018. However, Based on the exponential growth that stocks such as DQ and SEDG had in 2017, 2018 was not that significant of a drop in return, and DQ exponential increased it's Total Daily Volume.  

## Summary

### Advantages and Disadvantages of Refactoring Code 

An Advantage to refactoring code is the once you have it written the first time, you are able to find ways to rearrange it and make it better. One of the biggest disadvantages is that these alternative methods are not always "common-sense" and take a lot of research to configure the code in a way to where it still produces the same result, but in a quicker amount of time. 

### Comparing the Codes

The original VBA script provided a bit more logic in the sense that it was all a relative process within a process. Being unfamiliar with VBA, it was a lot more difficult for me to try to figure out how to create the missing puzzle pieces and then reconfigure them in a way that we hadn't necessarily done before. An advantage was knowing what the results should be and how to get them to land in the appropriate position. A disadvantage was also trying to figure out why the annual return was giving us an "overflow" error. Once you are able to identify the problem though, the solution starts to come easier and easier. 

