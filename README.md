# **Stock Analysis with VBA & Excel**

## **Overview of Project**


In this project, we are building off of code we put together that provided our friend Steve with a worksheet that enabled him with the click of a button, to analyze a list of 12 tickers for either 2017 or 2018. Steve was not content with his findings and has asked us to expand our dataset to include the entire stock market over the last few years. In theory, we could not expect Steve to use this same code to analyze thousands of tickers because it could run very slowly and inefficiently. There is a better way!

### Purpose
In this project we are going to update the previous script we wrote for Steve using refactoring. We want to know for certain if this new method works more efficiently so we are going to calculate the time it takes for our script to output the analysis.

The reason we are going to refactor the code is because this process creates a more efficient script, taking fewer steps, using less memory or improving the logic of the code to make it easier for the future users to use.



## **Results**

### Deliverable 1 - Explanation of code and what it is doing for us

> The code we were told to use is not the code that I put together for my previous exercise. The code we were told to use caused more confusion than good due to the comments using different syntax to explain what was needed, or just clearly stating different requirements.

> 1. The tickerIndex is set equal to zero before looping over the rows.
/Users/davidwiers/Desktop/Screen Shot 2022-01-10 at 1.41.20 PM.png
> 2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
/Users/davidwiers/Desktop/Screen Shot 2022-01-10 at 1.44.21 PM.png
> 3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.
/Users/davidwiers/Desktop/Screen Shot 2022-01-10 at 1.41.58 PM.png
> 4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
/Users/davidwiers/Desktop/Screen Shot 2022-01-10 at 1.42.06 PM.png
> 5. Code for formatting the cells in the spreadsheet is working.
/var/folders/ty/xt16j9x96x1ddw2r4j2jrzf00000gn/T/TemporaryItems/NSIRD_screencaptureui_QmcheQ/Screen Shot 2022-01-10 at 1.42.38 PM.png
> 6. There are comments to explain the purpose of the code.
/var/folders/ty/xt16j9x96x1ddw2r4j2jrzf00000gn/T/TemporaryItems/NSIRD_screencaptureui_l1WQdv/Screen Shot 2022-01-10 at 1.45.03 PM.png
> 7. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module.
/var/folders/ty/xt16j9x96x1ddw2r4j2jrzf00000gn/T/TemporaryItems/NSIRD_screencaptureui_5Ci8AV/Screen Shot 2022-01-10 at 1.45.38 PM.png
/var/folders/ty/xt16j9x96x1ddw2r4j2jrzf00000gn/T/TemporaryItems/NSIRD_screencaptureui_cVBoHI/Screen Shot 2022-01-10 at 1.46.06 PM.png
> 8. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png.
/var/folders/ty/xt16j9x96x1ddw2r4j2jrzf00000gn/T/TemporaryItems/NSIRD_screencaptureui_DSg9XR/Screen Shot 2022-01-10 at 1.46.41 PM.png

### Analysis
-The script does not take into account that there could be more than 12 tickers in the data in 2017 or 2018. Our script does not take that into account and is not able to pull in another ticker if one exists in the list due to us assigning a constant to our arrays.


## **Summary**

### 1. What are the advantages or disadvantages of using refactoring code?
The advantages of using refactored code are that:
- it improves the design of the code
- it makes the code easier to understand
- it allows you to find bugs
- it enables you to write code more quickly
Source - https://methodpoet.com/benefits-of-refactoring/

The disadvantages of using refactored code are that:
- if it is expected, then poor or sloppy code is sometimes written first with the expectation of refactoring later on.
- the output should still be the same as before, therefore resources (time & money) could be wasted if refactoring creates more work than it reduces through running.
- if the code is written in outdated language then refactoring would not help with the way the code was constructed
Source - https://rotate.cc/should-you-refactor-or-rewrite-your-code/

### 2. How do these pros and cons apply to refactoring the original VBA script?
The pros of the refactored the original VBA script are that:
- it obviously runs much quicker given our timer clocked around a 300% decrease in run time to produce our analysis

The cons of refactoring the original VBA code are that:
- it took a lot of time which was valuable for me for learning purposes, but in a real world situation Steve shouldn't care how long the script takes to run especially if its less than a few seconds.
- 