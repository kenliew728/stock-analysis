# Using Refactored VBA Code to Analyze Stocks
---
## **Overview of Project**
### Purpose
The purpose of this project is to refactor an existing VBA code in order to make the code more efficient by increasing its execution time
### Current State
The original VBA code was written to analyze 12 stock performances for the year 2017 and 2018. Both analyses took an average of 3.2 seconds to execute. This comes to an average of 0.26 seconds per stock per year. There are currently 2,600 active stocks being traded on the New York Stock Exchange and if we are to analyze all stocks, it will take 676 seconds or 11.3 minutes per analysis year to complete the execution. Therefore, it will be beneficial to refactor the existing VBA code to make it more efficient.
#### *Execution Time for both 2017 and 2018 based on Original VBA Code*
![Current State 2017](https://user-images.githubusercontent.com/70525492/93514687-1dda3a00-f8ed-11ea-87ff-ab3e4661dde2.png)
![Current State 2018](https://user-images.githubusercontent.com/70525492/93514857-4cf0ab80-f8ed-11ea-8530-cfb9ed982f31.png)
## **Results**
The refactoring work was done by reducing the "For" loop and 'If" statement reference to the ticker. 
The average run time on the refactored code was reduced to an average of 0.5 seconds, which is 0.042 seconds per stock, including formatting. This is equivalent to an 84% improvement. Assuming we running all 2,600 active stocks on NYSE, the total run time is now reduced to 108 seconds, or 1.8 minutes! 
#### *Execution Time for both 2017 and 2018 based on Refactored VBA Code*
![VBA_Challenge_2017](https://user-images.githubusercontent.com/70525492/93521064-13bc3980-f8f5-11ea-8935-935d0bae2060.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/70525492/93521066-16b72a00-f8f5-11ea-9926-525c1192a081.png)
