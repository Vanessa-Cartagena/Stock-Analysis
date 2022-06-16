# VBA Stock Analysis Challenge 

## Overview 

Steve's parents requested that he research a variety of stocks for them to consider investing in. Steve then asked for assistance with his analysis. We generated a code to run in an Excel workbook using Visual Basics for Applications, often known as VBA. He can analyze an entire dataset with a single click of a button. He now wants to expand the dataset to encompass the entire stock market over the last five years. Thousands of stocks are included in this extension, as opposed to the dozen or so that we worked on previously. This code may or may not work with this large dataset, or it may take a long time to run. In this challenge, we will refactor the code to loop through all the data once to collect the same data that we originally collected. 

## Results
###Original Code 

The original code contained a nest for loop that ultimately lead to a longer run-time by about 0.5 seconds. This code works completely fine as is but can be factored more efficiently. 
<img width="512" alt="All_Stocks_Analysis_Code" src="https://user-images.githubusercontent.com/105958160/173977679-82fef342-cfa7-40c2-92f1-76fb44bdb906.png">

<img width="264" alt="VBA_Challenge_2017-2" src="https://user-images.githubusercontent.com/105958160/173977726-3c5f82ac-fd70-49bb-b93b-2179b9488799.png">

<img width="260" alt="VBA_Challenge_2018-2" src="https://user-images.githubusercontent.com/105958160/173977744-84a342b1-b376-49fc-845e-490532dc786d.png">

###Refactored Code 

By refactoring the original code, we were able to loop through all the data once to collect the same data that we originally collected. This resulted in a shorter run-time. I restructured the for loop where I didn't require the nest loop for a more efficient and cleaner code. The ticker variable was assigned to the tickers array, and the other three arrays, tickerVolumes, tickerStartingPrices, and tickerEndingPrices, were each allocated to a tickerIndex. The data analysis was able to run significantly faster because it was assigned to a variable, the tickerIndex.
<img width="867" alt="All_Stocks_Analysis_Refactored" src="https://user-images.githubusercontent.com/105958160/173977807-6f0c897e-3c6f-492e-9336-89fc3d23c887.png">

<img width="272" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/105958160/173977823-88570558-e2ff-4638-80b9-cf8787974d9e.png">

<img width="260" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/105958160/173977833-2467c19c-6a6c-46c3-ad36-f62404c8e763.png">

## Summary 

The benefits of restructuring code are that it makes it cleaner, with faster execution and shorter procedures. It increases the code's readability and ease of understanding for other programmers since it provides a cleaner code. It can help with error analysis, maintainability of the software, and remove redundancies and duplications. Overall, it can help with performance. Refactoring can have a few drawbacks, including the introduction of new bugs and errors in the code, the time commitment, and the potential to influence testing results.
