# **Stock Analysis with VBA & Excel**

## **Overview of Project**


In this project, we are building off of code we put together for our good friend Steve. The code provided him with a worksheet that enabled him with the click of a button to analyze data from 12 stock tickers from either 2017 or 2018. Steve loved the file, but was not content with his findings and has asked us to expand our dataset to include the entire stock market over the last few years. In theory, we could not expect Steve to use this same code to analyze thousands of tickers because it could take forever to run. There must be a better way!

### Purpose
In this project we are going to update the previous script we wrote for Steve using refactoring. We want to know for certain if this new method works more efficiently. Therefore we will add a feature that calculates the time it takes for our script to output the analysis.

The reason we are going to refactor the code is because the refactoring process creates a more efficient script, takes fewer steps, uses less memory or improves the logic of the code to make it easier for the future users to use or understand.



## **Results**

### Deliverable 1 - Explanation of code and what it is doing for us

- The code we were told to use is not the code that I put together for my previous exercise. The code we were told to use caused more confusion than good due to the comments using different syntax to explain what was needed, or just clearly stating different requirements. That being said I was able to piece everything together to achieve a more efficient run time. The details of my code are included below!

> 1. The tickerIndex is set equal to zero before looping over the rows.
<img width="794" alt="Screen Shot 2022-01-10 at 1 41 20 PM" src="https://user-images.githubusercontent.com/95602006/148866628-5a7dc63e-6f77-4b93-af11-f186d1ead54f.png">

> 2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
<img width="752" alt="Screen Shot 2022-01-10 at 1 41 42 PM" src="https://user-images.githubusercontent.com/95602006/148866762-e0259e47-e8ad-46fd-8fc8-b0f3a130659c.png">

> 3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.
<img width="877" alt="Screen Shot 2022-01-10 at 1 41 58 PM" src="https://user-images.githubusercontent.com/95602006/148866831-85d9374b-f74a-4a57-b31e-960b8bdcf875.png">

> 4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
<img width="933" alt="Screen Shot 2022-01-10 at 1 42 06 PM" src="https://user-images.githubusercontent.com/95602006/148866856-25a1702b-239e-43e7-962a-7e443e0c1359.png">

> 5. Code for formatting the cells in the spreadsheet is working.
-See item number 7 for output format png

> 6. There are comments to explain the purpose of the code.
<img width="736" alt="Screen Shot 2022-01-10 at 7 51 02 PM" src="https://user-images.githubusercontent.com/95602006/148867550-33a58247-6f19-4c19-98ab-261772bcc3c6.png">

> 7. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module.
<img width="249" alt="Screen Shot 2022-01-10 at 7 50 15 PM" src="https://user-images.githubusercontent.com/95602006/148867567-6316900d-31e8-44a8-b740-ff3f5451f9a2.png">
<img width="233" alt="Screen Shot 2022-01-10 at 7 50 36 PM" src="https://user-images.githubusercontent.com/95602006/148867576-f795a8c4-a4a8-4ae7-932d-1de1cba23e21.png">

> 8. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png.
<img width="600" alt="Screen Shot 2022-01-10 at 7 53 48 PM" src="https://user-images.githubusercontent.com/95602006/148869463-85dbfec8-f030-4875-a5c2-9f712cc5befa.png">

### Analysis

a. The script does not take into account that there could be more than 12 tickers in the data in 2017 or 2018. Our script does not take that into account and is not able to pull in another ticker if one exists in the list due to us assigning a constant to our arrays.
b. The script ran significantly quicker, it required 300% less time - which could provide an advantage if there is a vast amount of data being analyzed. 
c. The script does not create actions that need to be completed within a small time frame and therefore does not add much value by being efficient. I'd imagine Steve would be okay if the script took up to 10 seconds, as long as we added a cool loading graphic that would convince Steve the computer and script are working very hard to return his analysis!
d. I still don't think this is enough information for Steve to make solid picks.  A tickers return for any fixed window of time can be misleading. There are quite a few more stats I would need to provide Steve in order for him to confidently pick some winning tickers. 52 wk high vs 52 wk low for example. 
e. We also don't really know what Steve is looking for. If his parameters for a tickers success are the tickers total daily volume and return, then we have succeeded. Does he want small cap or large cap, dividend paying or not? I have many questions for Steve in order to complete this analysis for him.

## **Summary**

### 1. What are the advantages or disadvantages of using refactoring code?
The advantages of using refactored code are that it:
- improves the design of the code.
- makes the code easier to understand.
- allows you to find bugs.
- enables you to write code more quickly.
> Source - https://methodpoet.com/benefits-of-refactoring/

The disadvantages of using refactored code are that:
- can promote using poor or sloppy code at first because they know that the code will be refactored by themselves or someone else in the future.
- the output is still be the same as before you refactored, therefore resources (time & money for example) could be wasted if refactoring creates more work the refactored code is worth.
- your code could be written in different languages or written in outdated syntax which would make any attempt for a successful refactoring impossible.
> Source - https://rotate.cc/should-you-refactor-or-rewrite-your-code/

### 2. How do these pros and cons apply to refactoring the original VBA script?
##### Pros
- An advantage of our refactored code is that we were able to decrease the amount of time it took to run by approximately 300%.
- I don't see any other benefits considering the script completes a simple task with only a small amount of data. I would have to update the script if I wanted to include any other tickers in our analysis and update the formatting code.
- I had to seek out resources to gain a good enough understanding of the topic to complete the challenge, major learning pro!

##### Cons
- It took a lot of time to refactor and there is no financial benefit to Steve or myself for creating the refactored code.
- If there is poor documentation then the script is difficult to refactor. The notes from the vba file I downloaded for the challenge were not helpful and did not follow the challenges instructions verbatim.
