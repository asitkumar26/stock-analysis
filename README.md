# stock-analysis
## Overview & Purpose of Project

Refactoring the 'All Stocks Analysis' script to make it perform better in terms of Run time so that Steve can use the modified/improvised script to do the analysis of all Stocks of the entire stock market. 

## Analysis and challenges

### Analysis during refactoring

During the analysis to refactor, following steps were involved.
    1. The way source data in '2017' and '2018' worksheets are arranged for all stickers.
        - The tickers that are used for look up in this code have been arranged in ascending order (A to Z). This fact may result in reducing one of the external loop of the nested loop of original code.
    
    Add the link for  Ticker-Values
     
    2. The way script is written in original code
        - Because of the nested loop, the inner loop that contains three calculations ('total violume', 'ticker starting price' and 'ticker ending price') will be executed 'no of tickers' * no of rows of the corresponding year worksheet' times which could be time consuming. 

    Add the link for  Loop-original code
    
### Challenges to refactor for improvement

The real challenge was to avoid nested loop for calculations so that the script can come up with the desired result faster than the original code.

### Techniques applied 

Since all tickers in source worksheets ('2017' and '2018' worksheets) are arranged in sorted order (A to Z), the ticker look up logic has been modified where ticker value changes as and when all rows of the corresponcing ticker gets processed completely through the loop. The ticker value is controlled through ticker index in the script.

Add the link for  Ticker-Index-modifiedcode

### Efforts spent 

Amount of effort was more during analysis step where data analysis was done during refactoring and also it took time to apply the logic of changing the ticker value as soon as all corresponding rows are processed.


## Results

The refactored code with modified logic is much faster than the original code. 

Link for VBA_Challenge_2017
Link for VBA_Challenge_2018


### Advantage of refactored code

Now refactored script can scale to much larger datasets since it performs much better compared to original code. It may be able to analyze all tickers of the Stock Market as intended.

### Dis-advantage of the refactored code

Though more time was spent in analysis and figuring out a logic to pmprove the runtime, but it was worth spending that much time and effort to come up with a better performing script. 



