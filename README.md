# Stock_Analysis
Module 2 VBA Script Challenge - stock analysis 
# Stock Analysis – by Maha Shah 

## Overview of the Project 

The purpose of this stock analysis is to help Steve loop through multiple and various stocks to make an informed decision based on researched stock figures. This involved refactoring a previous code already determined for Steve and his analysis. Refactoring involves building on an already established and defined code. Code refactoring will help make it more efficient in its output, while making the code itself more dynamic and useful. 

## Results 

### 2017 and 2018 Stock Returns Comparison 

As can be seen in the pictures below, Tickers RUN and ENPH are the only two options that had a positive year to year variance in the return. It can be concluded that these are two more stable stock options with a higher positive return. Overall, 2017 was a better year for the stocks analyzed in this data set. Whereas 2018 was particularly challenging more the majority of the stocks. 

![2017Returns](https://github.com/msha789/Stock_Analysis/blob/694e89831bf17f60a04bb8d85fc16154f31f6bda/Screen%20Shot%202022-01-09%20at%206.43.47%20PM.png)

![2018Returns](https://github.com/msha789/Stock_Analysis/blob/694e89831bf17f60a04bb8d85fc16154f31f6bda/Screen%20Shot%202022-01-09%20at%206.44.11%20PM.png)

### Refactored VBA Script 

After creating the three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices, the tickerIndex was used to access the stock ticker indices. For this a loop was created to initialize the tickerVolumes to zero and if the following did not match to the before then we increased the tickerIndex. This loop also loops over all the rows in the spreadsheet. 

Lastly, all the outputs were put into the respective columns and formatted so that positive outcomes were formatted green and negative red. This all can be seen in the code pictured below. 

![Code1](https://github.com/msha789/Stock_Analysis/blob/694e89831bf17f60a04bb8d85fc16154f31f6bda/Screen%20Shot%202022-01-14%20at%209.03.16%20PM.png)

![Code2](https://github.com/msha789/Stock_Analysis/blob/694e89831bf17f60a04bb8d85fc16154f31f6bda/Screen%20Shot%202022-01-14%20at%209.03.44%20PM.png)


### Timers 

The refactoring of code resulted in a time difference in the execution of the code. The initial codes ran for 0.805seocnds for 2017 and 0.813seconds for 2018. The refactored code, since more complex, runs for 0.844 seconds for 2017 and 0.828 seconds for 2018. The refactored code runs at a slightly longer time, as can be seen in the pictures below. 

#### Original 

![OG17](https://github.com/msha789/Stock_Analysis/blob/694e89831bf17f60a04bb8d85fc16154f31f6bda/VBA%202017%20OG.png

![OG18](https://github.com/msha789/Stock_Analysis/blob/694e89831bf17f60a04bb8d85fc16154f31f6bda/VBA%202018%20OG.png)

#### Refactored 

![RF17](https://github.com/msha789/Stock_Analysis/blob/694e89831bf17f60a04bb8d85fc16154f31f6bda/VBA_Challenge_2017.png)

![RF18](https://github.com/msha789/Stock_Analysis/blob/694e89831bf17f60a04bb8d85fc16154f31f6bda/VBA_Challenge_2018.png)

## Advantages and Disadvantages of Refactoring Code 

There are a number of pros and cons to refactoring an existing code or script. 

### Advantages 

Refactoring code provides the advantage of being able to build on an already functioning script. This helps with building more efficiencies onto an already written code. This also provides the chance to catch inefficiencies built into a working code. 

Refactoring can help catch bunched up loops and nested loops. It provides the opportunity for clarification and further structuring of the code already written for better execution and analysis. 

### Disadvantages 

However, refactoring also has disadvantages. This includes a longer procedure of re-writing working and functioning codes. Sometimes these can be longer and more iterative. Another disadvantage for refactoring code is that it may be more time consuming to find erroneous areas to improve rather than write a code you know would work from scratch.  

### Pros and Cons as they apply to the Original VBA Script 
 
The refactoring of the original VBA script was a time-consuming process that resulted in a longer execution time for a similar analysis and result. Although a more sophisticated code, this resulted in the same analysis in the end. However, this was a good exercise to catch logical errors within the script and understand the logic and syntax of the VBA script at a deeper level. The nature of refactoring as a trial and error learning exercise is quite valuable. 

