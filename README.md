# Refactoring VBA for performance gains

## Overview of Project
The existing stock ticker code performance is terrible. This project aims to refactor the current code to process stock ticker data with a single pass through the spreadsheet.

## Results

### Analysis of single loop method vs. n(ticker) loops.
The single loop method is %24600 faster than executing a nested loop. When performing the nested loop of an outer loop for each stock ticker and an inner loop of the stock data for that year, the machine becomes unresponsive due to the additional CPU resource demand.

VBA cannot take advantage of multiple cores, so the result of non-optimized code is an inefficient use of resources.

![CPU LOAD](https://raw.githubusercontent.com/skanab/stock-analysis/main/Resources/CPU%20Load.PNG)

![Comparison](https://raw.githubusercontent.com/skanab/stock-analysis/main/Resources/Compare.png)


## Summary

### What are the advantages and disadvantages of refactoring code?

The most common reasons for refactoring code is to improve performance, reduce complexity, and ensure good code hygiene. Refactoring assures code remains readable and serviceable. Refactoring projects can reduce labor costs if they are managed correctly.

Starting a refactoring project without a clear understanding of the codebase could introduce more bugs and complexity. Fixing one section of code could lead to problems dependent segments. Refactoring projects can quickly become unmanageable if the scope creeps.

### What are the advantages and disadvantages of refactoring code for this project?

The single loop method was able to process data much more efficiently. The user can now explore other stock tickers without incurring additional wait time. Any time a user has to wait, they will context switch into another task or distraction, which will lead to user time wasted time and the wasted processing time.

I see no advantage of the original code. 
