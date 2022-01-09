# **Stock Analysis with VBA & Excel**

## **Overview of Project**

### Purpose
In this project, we are building off of code we put together that provided our friend Steve with a worksheet that enabled him with the click of a button, to analyze a list of 12 tickers for either 2017 or 2018. Steve was not content with his findings and has asked us to expand our dataset to include the entire stock market over the last few years. In theory, we could not expect Steve to use this same code to analyze thousands of tickers because it could run very slowly and inefficiently. There is a better way!
In this project we are going to update the previous script we wrote for Steve using refactoring. We want to know for certain if this new method works more efficiently so we are going to calculate the time it takes for our script to output the analysis.
The reason we are going to refactor the code is because this process creates a more efficient script, taking fewer steps, using less memory or improving the logic of the code to make it easier for the future users to use.



## **Results**

### Deliverable 1

>Create tickerIndex

>Create 3 output arrays

>Create for loop to initialize tickerVolumes to 0

>Create a for loop that loops over all the rows in the same spreadsheet

>Inside for loop in 2b, write script that increases current tickerVolumes and adds the ticker volume for the current stock ticker, using the variable tickerIndex as the index.

>Write an if-then statement to check if current row is first row

>Write an if-then statement to check if the current row is the last row

>Write script that increases the tickerIndex if the next row's ticker doesn't match previous row

>Use a for loop to loop through our arrays: tickers, tickerVolumes, tickerStartingPrices and tickerEndingPrices

>Run the stock analysis and confirm outputs for 2017 and 2018 match our results from the AllStockAnalysis, to output the "Ticker, Total Daily Volume and Return" columns.


## **Summary**

### 1. What are the advantages or disadvantages of using refactoring code?

### 2. How do these pros and cons apply to refactoring the original VBA script?