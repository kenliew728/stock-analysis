# Using Refactored VBA Code to Analyze Stocks
---
## **Overview of Project**
### Purpose
The purpose of this project is to refactor an existing VBA code in order to make the code more efficient by increasing its execution time
### Current State
The original VBA code was written to analyze 12 stock performances for the year 2017 and 2018. Both analyses took an average of 3.47 seconds to execute. This comes to an average of 0.29 seconds per stock per year. There are currently 2,600 active stocks being traded on the New York Stock Exchange and if we are to analyze all stocks, it will take 752 seconds or 12.5 minutes per year of analysis to complete the execution. Therefore, it will be beneficial to refactor the existing VBA code to make it more efficient.
#### *Execution Time for both 2017 and 2018 based on Original VBA Code*
![Current State 2017](https://user-images.githubusercontent.com/70525492/93521891-49155700-f8f6-11ea-9bb9-53f53d97f367.png)
![Current State 2018](https://user-images.githubusercontent.com/70525492/93521896-4a468400-f8f6-11ea-8f2c-a42771067b8d.png)
## **Results**
The refactoring work was done by eliminating the reference to the ticker using the "If" loop in the original VBA code. 
#### *Reference to the Orignal Nested loop Stock Analysis VBA code
![image](https://user-images.githubusercontent.com/70525492/93629866-99e88680-f9ae-11ea-98d5-12f11a5b748f.png)

The average run time on the refactored code was reduced to an average of 0.5 seconds, which is 0.042 seconds per stock, including formatting. This is equivalent to an 86% improvement. Assuming we running all 2,600 active stocks on NYSE, the total run time is now reduced to 108 seconds, or 1.8 minutes! To put in perspective, if an analyst need to run all stock performances for the past 10 years, it will take the analyst 18 minutes to run, compare to 125 minutes!
#### *Execution Time for both 2017 and 2018 based on Refactored VBA Code*
![VBA_Challenge_2017](https://user-images.githubusercontent.com/70525492/93630422-aa4d3100-f9af-11ea-8b0c-8750dd9290e4.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/70525492/93630426-aae5c780-f9af-11ea-8a7b-6c74314b1ab0.png)

