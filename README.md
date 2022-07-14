# Stock Analysis using Excel VBA

## Overview of Project
Steve has asked that the original green stock analysis code be refactored to analyze the entire dataset of stocks. The script needs to be able to recieve input and analyze the correct year of stocks. Then it should provide the resulting formatted summary.

## Resouces
  - Challenge File: VBA_Challenge.xlsm
  - Excel Version 2206 (Microsoft 365), Excel VBA

## Results

### Stock Performance for 2017 and 2018
When considering the summary analysis of 2017 and 2018 results, there was a significant difference between the returns in 2017 and 2018. As seen in image 1.1 below, the yearly returns for 2017 show a positive return for all but TERP stock. The higest return was DQ with 199.4% which would have been the best investment for 2017. 2018 results show in image 1.2 below. It appears that there was a large downturn in green stocks for nearly every stock ticker listed except for ENPH and RUN. DQ experienced this downturn as well with a -62.6% return for 2018. RUN and ENPH represent the two highest and ony positive returns for 2018 with 84.0% and 81.9% respectively.

#### Visuals
| Image 1.1: Stock Performance 2017 | Image 1.2 Stock Performance 2018 |
| ----------------------------- | ----------------------------- |
| <img src="/Resources/VBA_Challenge_2017_Stock_Results.png" width="400"> | <img src="/Resources/VBA_Challenge_2018_Stock_Results.png" width="400"> |


### Execution Times for the Original and Refactored Code
The purpose of this code refactor was to allow the script to analyze all of the company returns for any year as well as to attempt to increase efficency. In regards to the effeciancy change, the refactored code running for 2017 ran for 0.106 seconds whereas the original ran for 0.7344. This represented a 85.6% reduction in the execution time. For 2018, the efficency increase was basically the same. The refactored code running for 2018 also ran for 0.106 seconds whereas the original ran for 0.7617. This represented a 86.1% reduction in the execution time. Refactoring this code has succeded in the goal to increase effeciency. (Please see the visuals found below for a table of results and the screenshots of each time the script was run)

#### Visuals


| Code Version  | Year | Execution Time |
| ------------- | ------------- | -------------- |
| Original  | 2017 | 0.7344 seconds|
| Refactored  | 2017 | 0.106 seconds |
| Original  | 2018 | 0.7617 seconds |
| Refactored  | 2018 | 0.106 seconds |

 | Image 2.1: Original Code 2017 | Image 2.2: Refactored Code 2017 |
 | ----------------------------- | ----------------------------- |
 | <img src="/Resources/VBA_Challenge_2017_Original.png" width="400"> | <img src="/Resources/VBA_Challenge_2017_Refactored.png" width="400"> |
  
 | Image 2.3: Original Code 2018 | Image 2.4: Refactored Code 2018 |
 | ----------------------------- | ----------------------------- |
 | <img src="/Resources/VBA_Challenge_2018_Original.png" width="400"> | <img src="/Resources/VBA_Challenge_2018_Refactored.png" width="400"> |
 

## Summary

### Refactoring Code
 - The advantages of refactoring code are:
  1. Refactoring code allows the coder to add additional features and functions to the code.
  2. Refactoring code can also be used to improve the efficiency of the code.
  3. Refactoring code can also be used when debugging to deal with new errors that have been discovered by the end users.
 - The disadvantages of refactoring code are:
  1. Refactoring code can add new bugs into the code
  2. Refactoring code may not always improve efficiency
  3. Refactoring code that the coder did not write requires understanding the code first before it can be adjusted.

### Original VBA Script
 - The advantages of the orginal VBA script are:
  1. The original code already works to provide the results needed for a single stock
  2. The original code has been tested and used by the end user so it functions in the real world environment
    
 - The disadvantages of the orginal VBA script are:
  1. Missing features requested by the end users
  2. Less efficient than the refactored code

### The Refactored VBA Script
  - The advantages of the refactored VBA script are:
  1. The code was updated to add additional requested features.
  2. The code was altered to improve efficiency.
  3. The code can be used in the future with additional data sets for earlier or later years (as long as the data structure does not change).

 - The disadvantages of the refactored VBA script are:
  1. The code required additional time and resources to add features and increase efficiency.
  2. The code will work with the stocks already listed, but if new stocks are added, they will be skipped by this script.
