# Using Refactored VBA Code to Analyze Stocks
---
## **Overview of Project**
### Purpose
The purpose of this project is to refactor an existing VBA code in order to make the code more efficient by reducing the total execution time
### Current State
The original VBA code was written to analyze 12 stock performances for the year 2017 and 2018. Both analyses took an average of 3.76 seconds to execute. This comes to an average of 0.31 seconds per stock per year. There are currently 2,600 active stocks being traded on the New York Stock Exchange (NYSE) and if an analyst is asked to analyze all stocks on a given year, it will take 832 seconds or 13.9 minutes to complete the execution. Therefore, it will be beneficial to refactor the existing VBA code to make it more efficient.
#### *Execution Time for both 2017 and 2018 based on Original VBA Code*
![Current_State_2017](https://user-images.githubusercontent.com/70525492/93630787-2b0c2d00-f9b0-11ea-97f4-95db4ee4930b.png)
![Current_State_2018](https://user-images.githubusercontent.com/70525492/93630788-2b0c2d00-f9b0-11ea-9b6f-287184d4a6e4.png)
## **Results**
The refactoring work was done by eliminating the reference to the ticker within the nested loop in the original VBA code. By eliminating reference to the ticker, it reduced an extra step in the looping process.
#### *Comparison between the Orignal Nested loop vs. Refactored Nested Loop*
|   Year   | Runtime (Current) | Runtime (Refactored) |
| -------- | ----------------- | -------------------- |
|   2017   |       3.36        |        0.46s         |
|   2018   |       4.15        |        0.55s         |
|  Average |       3.76        |        0.51s         |

![VBA Original_Nested](https://user-images.githubusercontent.com/70525492/93631484-49265d00-f9b1-11ea-9455-6ceefb8944ae.png)
![Refactored_VBA](https://user-images.githubusercontent.com/70525492/93631821-dff31980-f9b1-11ea-98a7-a1b897398260.png)

The average run time on the refactored code was reduced to an average of 0.51 seconds or 0.042 seconds per stock, which is equivalent to an 87% improvement. Assume an analyst needing to run all 2,600 active stocks on NYSE, the total run time is now reduced to 109 seconds, or 1.8 minutes! To put in perspective, if an analyst need to run all stock performances for the past 10 years, it will take the analyst 18 minutes to run with the refactored code, compare to 139 minutes in the old one!
#### *Execution Time for both 2017 and 2018 based on Refactored VBA Code*
![VBA_Challenge_2017](https://user-images.githubusercontent.com/70525492/93630422-aa4d3100-f9af-11ea-8b0c-8750dd9290e4.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/70525492/93630426-aae5c780-f9af-11ea-8a7b-6c74314b1ab0.png)

