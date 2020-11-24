# Stock Analysis

## Overview of Project
The goal of this project is to utilzie stock market data to help analyze what stocks would be best for Steve to invest in. To do this I wrote a Macro in VBA that would analyze different aspects of the stock data to see how each stock performed in 2017 and 2018. Another big part of this project was refactoring the previously written code to make the VBA script run faster anbd more efficiently than before. This is important because it was used to analyze thousands of stocks.  

## Results: 
### Values of Stocks Decreased Over Time 
The majority of the stocks decreased in value from 2017 to 2018. The only two stocks that performed better in 2018 than 2017 were TERP and RUN. You can see the different amounts of volume and return percentages for all of the stocks [here](https://github.com/jmerenstein/stock-analysis/blob/main/VBA_Challenge.xlsm). To view the results for the different years go into the VBA editor and run the bottom script. This will prompt you to enter which year you would like the results for. Once you enter either 2017 or 2018 you can see the specific results for each stock yourslef. 

### Refactoring Code Reults
The original script did run quicker than the refactored script, but the different was negliglible. For example the orginal script ran the analysis for 2017 in about .047 seconds. The refactored code for 2017 ran in about .168 seconds. The execution time of the refactored code can be seen [here](https://github.com/jmerenstein/stock-analysis/blob/main/VBA_Challenge_2017.png). So, even though the refactored code was slow, it was still very fast and less than a second. 

## Summary
- The advantages of refactoring a code are making it more efficient and making it easier to read. By going through the code again, you find places where the code could have possibly be written in a more efficient or "clean" manner so when it is run again it will take less steps and be easy to follow for yourslef or anyone else reading it.

- The disadvantages of refactoring a code could be taking extra time to have a slight imporvement over something that already worked. If it takes a long time to refactor the code and the original code did work successfully, I could see refactring as a possible waste of a person's time.

- The main adavatage of refactoring the stock analysis code was that it put all of the different Macros we performed in the Module into just one Macro for the Challenge. For example during the Module we had different Macros for formatting the cells, running the analysis, and getting the timer execution up. For the Challenge we were able to combine all of those steps together in just one Macro. This did result in the Macro running slightly slower, but not slow enough that it was not worth refactoring. Overall, the advantage of having all of the steps in one Macro instead of multiple outweighed the disadvanatage of the Macro running at a slightly slower speed. 


