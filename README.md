# Using Refactored VBA Code to Analyze Stocks
---
## **Overview of Project**
### Purpose
The purpose of this project is to refactor an existing VBA code in order to make the code more efficient and to reduce the total execution time
### Current State
The original VBA code was written to analyze 12 stock performances for the year 2017 and 2018. Both analyses took an average of 3.76 seconds to execute. This comes to an average of 0.31 seconds per stock per year. There are currently 2,600 active stocks being traded on the New York Stock Exchange (NYSE) and if an analyst is asked to analyze all stocks in any given year, it will take 832 seconds or 13.9 minutes to complete the execution. Therefore, it will be beneficial to refactor the existing VBA code to make it more efficient.
#### *Execution Time for both 2017 and 2018 based on Original VBA Code*
![Current_State_2017](https://user-images.githubusercontent.com/70525492/93630787-2b0c2d00-f9b0-11ea-97f4-95db4ee4930b.png)
![Current_State_2018](https://user-images.githubusercontent.com/70525492/93630788-2b0c2d00-f9b0-11ea-9b6f-287184d4a6e4.png)
## **Results**
The refactoring work was done by eliminating the reference to the ticker within the nested loop in the original VBA code. By eliminating reference to the ticker, it reduced an extra step in the looping process.
#### *Comparison between the Orignal Nested loop vs. Refactored Nested Loop*
|   Year   | Runtime (Current) | Runtime (Refactored) |
| -------- | ----------------- | -------------------- |
|   2017   |       3.36s       |        0.46s         |
|   2018   |       4.15s       |        0.55s         |
|  Average |       3.76s       |        0.51s         |

![VBA Original_Nested](https://user-images.githubusercontent.com/70525492/93631484-49265d00-f9b1-11ea-9455-6ceefb8944ae.png)
![Refactored_VBA](https://user-images.githubusercontent.com/70525492/93631821-dff31980-f9b1-11ea-98a7-a1b897398260.png)

The average run time on the refactored code was reduced to an average of 0.51 seconds or 0.042 seconds per stock, which is equivalent to an 87% improvement. Assume an analyst needing to run all 2,600 active stocks on NYSE, the total run time is now reduced to 109 seconds, or 1.8 minutes! To put in perspective, if an analyst needs to run all stock performances for the past 10 years, it will take the analyst 18 minutes to run with the refactored code, compare to 139 minutes in the old one!
#### *Execution Time for both 2017 and 2018 based on Refactored VBA Code*
![VBA_Challenge_2017](https://user-images.githubusercontent.com/70525492/93630422-aa4d3100-f9af-11ea-8b0c-8750dd9290e4.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/70525492/93630426-aae5c780-f9af-11ea-8a7b-6c74314b1ab0.png)
## **Summary**
### *What are the advantages or disadvantages of refactoring code?*
The advantages and disadvantages of refactoring code can be summarized as follow:
1. Advantages
    - Allows code to run more efficiently.
    - Makes the code easier to understand or to debug.
    - Improves the quality of the coding.
2. Disadvantages
    - Refactoring may take up too much time to complete, resulted in diminishing returns.
    - Time equals money and if all programmers took a long time to refactor the codes, it will definitely increase the cost of the project and may delay the projected completion time.
### *Advantages and Disadvantages of the Original and Refactored VBA Script*
#### Orignal VBA Script
1. Advantages
    - Less maintenance when more tickers are added into the data. For example, if one ticker is added, only the ticker dimension count and the For "i" loop count need to be updated
2. Disadvantages
    - Less efficient and takes more time to execute. 
#### Refactored VBA Script
1. Advantages
    - Definitely more efficient to run
2. Disadvantages
    - May spend more time to update script when more ticker is added. For example, if one ticker is added, the following scripts will need to be updated
        - Ticker dimension count
        - Output array count (total 3 arrays)
        - For "tickerIndex" loop count

