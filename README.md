## Overview of Project
Steve’s parents are looking forward to invest in an energy company. There haven’t done much research and are on the verge of investing all of their money in DAQO New Energy Group (DQ). Steve with his background is willing to look into the stocks of DQ but he is interested in diversifying his parents funds. He wants to analyze are green energy company stocks so he has a stock file and will want us to help him analyze these stocks. To achieve this, I employed the use of Visual Basic Applications (VBA) to interact with  Excel.  VBA allows me to make calculations, use complex logics to perform analysis and most importantly using code with automated analysis allowing Steve to use it for any stock.
Once I completed the worksheet which allowed Steve to analyze the entire dataset, he preferred expanding  the data set to include the entre stock market over the last few years. To achieve this, I had to edit (refactor) the solution code to loop through all the data one time to collect the same information. The purpose of refactoring was to ensure that the code was more efficient. 

## Results

### Analysis
To determine whether refactoring the code was more efficient, I began addressing the first module (challenge) Steve proposed which was finding out if DQ was the best option I started by analysis DQ’s stock and realized it might not be the best stocks for his parents to invest in. 

<img width="420" alt="Screen Shot 2021-03-04 at 11 03 09 PM" src="https://user-images.githubusercontent.com/78401776/110071068-58277e80-7d49-11eb-8b90-a77bd8ca1a57.png">

Once I completed that task, I analyzed multiple stocks to find some better choices for his parents which I termed the “All Stocks Analysis” . To run analysis on all the stocks, I created a program flow that looped through all of the tickers. The results for this analysis showed that the code takes **0.4960938** seconds to ran for year 2017 and **0.484375** seconds to ran in 2018. 

<img width="290" alt="Screen Shot 2021-03-04 at 11 21 43 PM" src="https://user-images.githubusercontent.com/78401776/110071084-65dd0400-7d49-11eb-8a34-cefdd6507629.png">

After performing this code, I went a step further to refactor it to determine if the code was sufficient enough and to loop through the entire data at the same time. I also add formatting changes to show which stocks made great returns. I create three output arrays for this code (tickerVolumes, tickerStartingPrices and tickerEnding Prices).  Formatted the cells (colors) to in order to have a better and more efficient reading of the data. Green being returns that were good and red the opposite. 

The stock analysis outputs for the “All Stocks Analysis Refactored” were the same as that of the “All Stocks Analysis”.  
The stock performances in 2017 and 2018 shows that 2017 was a better year for most of the companies. They had a better return as compared to 2018. Per the analysis below, TERP stock was the only company that had a negative return in 2017.  

<img width="312" alt="Screen Shot 2021-03-04 at 11 48 33 PM" src="https://user-images.githubusercontent.com/78401776/110071241-ad639000-7d49-11eb-9383-bf6b07b57dfb.png">

In 2018, there were only two (2) tickers that had good returns. These tickers include “RUN” and “ENPH”. 

<img width="314" alt="Screen Shot 2021-03-04 at 11 48 11 PM" src="https://user-images.githubusercontent.com/78401776/110071212-a50b5500-7d49-11eb-8b5b-9e20b981a113.png">

With regards to the execution times, the refactored script had a lesser execution time for both years than the original script. It means that, the refactored script is more efficient. 

<img width="890" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/78401776/110071276-c0766000-7d49-11eb-9911-7946ac648e8b.png">

<img width="885" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/78401776/110071301-c66c4100-7d49-11eb-874b-ba11bab50f3b.png">

## Summary

### General

#### Advantages
- It helps in a faster execution making the code more efficient.
- It makes the software easier to understand and improves the design of the software.
- It makes it easier for other users to read.
- Helps the developer to improve the logic of the code through expanding it.

#### Disadvantages
- It is expensive and may introduce bugs that the test may not catch. 
- It poses a greater risk as compared to an original script

### Advantages and Disadvantages of Original and Refactored VBA Script
VBA is a great tool to help perform automated analysis and solve complex logics. It helped Steve advise his parents on which stocks to invest in. Performing the refactored script gave me a deeper understanding of VBA and how to use it in performing such analysis by thinking broadly. It was more efficient than the original script. The original script was easier to use even though both scripts provided the same outputs.  


