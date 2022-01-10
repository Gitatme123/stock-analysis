# **Stock Analysis with VBA & Excel**

## **Overview of Project**


In this project, we are building off of code we put together that provided our friend Steve with a worksheet that enabled him with the click of a button, to analyze a list of 12 tickers for either 2017 or 2018. Steve was not content with his findings and has asked us to expand our dataset to include the entire stock market over the last few years. In theory, we could not expect Steve to use this same code to analyze thousands of tickers because it could run very slowly and inefficiently. There is a better way!

### Purpose
In this project we are going to update the previous script we wrote for Steve using refactoring. We want to know for certain if this new method works more efficiently so we are going to calculate the time it takes for our script to output the analysis.

The reason we are going to refactor the code is because this process creates a more efficient script, taking fewer steps, using less memory or improving the logic of the code to make it easier for the future users to use.



## **Results**

### Deliverable 1 - Explanation of code and what it is doing for us

> 1. The tickerIndex is set equal to zero before looping over the rows.

> 2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

> 3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.

> 4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

> 5. Code for formatting the cells in the spreadsheet is working.

> 6. There are comments to explain the purpose of the code.

> 7. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module.

> 8. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png.

### Analysis

#### The script does not take into account that there could be more than 12 tickers in the data in 2017 or 2018. Our script does not take that into account and is not able to pull in another ticker if one exists in the list. Therefore I believe our script failed Steve.  



## **Summary**

### 1. What are the advantages or disadvantages of using refactoring code?

### 2. How do these pros and cons apply to refactoring the original VBA script?