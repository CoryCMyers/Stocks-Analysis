# <div align="center">Stocks-Analysis</div>

Repository for the Stocks Analysis Project 

# <div align = "center">Analysis Summary</div>

## <div align = "center">Overview of Project</div>

  The purpose of this analysis is to summarize and then contrast the performance of the listed Tickers over the course of a year, as well as displaying the overall return on profit for that stock over the course of that year. This Project was made in a way that so long as the Data entered into the Workbook in the same manner and the same naming conventions followed on the name of individual sheet names it will continue to function moving forward.
  
## <div align = "center">Overview of Results</div>

#### <div align = "center">2017</div>
When looking at the results of the 2017 Data it can be seen that overall for the year the stocks selected to be followed for this exercise had an overall positive trend overall. From the information that you provided for the basic creation of the spreadsheet these stocks would seem to be a solid choice of investment with only one TERP showing a neagative return for the year as we see here. 

[2017 Stocks Analysis Results](https://github.com/CoryCMyers/Stocks-Analysis/blob/main/VBA_Challenge_2017.png) ![2017 Stocks Analysis Results](https://github.com/CoryCMyers/Stocks-Analysis/blob/main/VBA_Challenge_2017.png) 

#### <div align = "center">2018</div>

However, when further information is provided and the results for the year of 2018 are also added this story changes. 

[2018 Stocks Analysis Results](https://github.com/CoryCMyers/Stocks-Analysis/blob/main/VBA_Challenge_2018.png) ![2018 Stocks Analysis Results](https://github.com/CoryCMyers/Stocks-Analysis/blob/main/VBA_Challenge_2018.png)

When we consult the information from the analysis of the 2018 stock ticker history the situation seems to be entirely different. From all of the stocks that were covered as profitable in the previous year only one of them is still profitable in the most recent data set provided. While this stock did increase exponentially in value over the course of the year, the other stock that remained profitable also took a loss from the previous year while the remained all took a sharp decline.

### <div align = "center">Conclusions</div>

Without access to further years of information to show any sort of trend to see if there is an ebb and flow to these tickers to determine if the current downturn is a normal part of the stocks price cycle, making it a good time to invest none of these tickers have shown to be educated choices for investment at the time of this analysis.

## <div align = "center">Runtime Comparisons</div>

#### <div align = "center">2017</div>

The runtime for the original and refactored Macros can be found below.[^1]

2017 Original | 2017 Refactored 
:-------------------------:|:-------------------------:
![2017 Runtime Original](https://github.com/CoryCMyers/Stocks-Analysis/blob/CoryCMyers-patch-1-workingOnReadme/2017Base.PNG)  |  ![2017 Runtime Refactored](https://github.com/CoryCMyers/Stocks-Analysis/blob/main/VBA_Challenge_2017.png)

#### <div align = "center">2018</div>

2018 Original             |  2018 Refactored
:-------------------------:|:-------------------------:
![2018 Runtime Original](https://github.com/CoryCMyers/Stocks-Analysis/blob/CoryCMyers-patch-1-workingOnReadme/2018Base.PNG)  |  ![2018 Runtime Refactored](https://github.com/CoryCMyers/Stocks-Analysis/blob/CoryCMyers-patch-1-workingOnReadme/VBA_Challenge_2018.png)

#### <div align = "center">Code Comparisons</div>

Original Code | Refactored Code 
:-------------------------:|:-------------------------:
![Original Code](https://github.com/CoryCMyers/Stocks-Analysis/blob/main/Analysis_Code_Original.PNG)  | ![Refactored Code](https://github.com/CoryCMyers/Stocks-Analysis/blob/main/Analysis_Code_Refactored.PNG)


# <div align = "center">Summary</div>

The primary change between these two codes, and the difference in their runtimes can be traced to the change in code from using

```
'When the code is written like this, then each time this code loop is run it must verify both cells being referenced for each value each loop.
  If Cells(iteratorNumber - 1, columnNumber).Value <> ticker And Cells(iteratorNumber, columnNumber).Value = ticker Then
    startingPrice = Cells(iteratorNumber, columnNumber).Value
```

The code when written like this for every iteration through the code it needs to check the cell at address Cells(i - 1, columnNumber) value does not equal the ticker currently being checked, if it does it then checks if Cells(i, columnNumber).Vale does equal the ticker, then it assigns the current startingPrice value plus the value of the cell currently targeted whose ticker does equal the ticker value to the variable starting price. It must run this full loop for every iteration.

However, when the code has been refactored to run more effeciently that same code looks like this

```
'Writing the code like this uses variables already being held within the function rather than having to stop and check other multiple cells
If Cells(iteratorNumber - 1, columnNumber).Value <> tickers(tickerIndex) Then
  tickerStartingPrices(tickerIndex) = Cells(iteratorNumber, columnNumber).Value
```

This way of writing however instead of refrencing specific cells that have to be sequentially checked and then stopping to check the value of another cell it instead uses values already stored in variables from previously in the function. This allows the function to run faster as it reduces the overall number of calculations that need to be run on each iteration.


## <div align = "center">Refactored Code Pro/Con</div>

Refactored Code Pro       | 
:-------------------------:

- Able to begin from a shared startpoint if working as a group
- Code that has proven functional can be used as a control when testing functionality
- Refactoring code you wrote previously reinforces conceptual connections
- Refactoring your previous code in GitHub shows a performance standard
- Allows for code refinement to enhance code performance

Refactored Code Con       |
:-------------------------:

- Working from anothers code does not teach foundational concepts
- Working from anothers code does not teach creativity
- Can take longer to understand if code is refactored at a higher conceptual level than the original writers
- Variables can be unintentionally broken much easier
- Can add additional time to submission if variables are missed causing extra read throughs of the code
- If data set is not verified to match original code then additional time is needed to correct original code

### <div align = "center">How do these apply to the Code?</div>
  
When looking at the ways that the specific code examples seen here there are advantages to both of them. For the original writing of the code then the information was laid out clearly. There was no requirement of understanding for how the different arrays functioned and the iteration loops were simple and sequential. Additionally in the original code there was no understanding of how a function will hold a variable and then the availability of that information to be used for additional purposes.
This allows the original code to be written very simply and like it is being driven from point A to point B with a minimum of understanding needed of just how complicated the function can begin to be.

When you look at the refactored code however the overall code is less bulky, and has a cleaner flow to it allowing it to run more swiftly. However to be able to write and understand code that uses this style rather than the original style requires an understanding of the fact that you can use previously established variables in new ways to temporarily hold additional information. Doing things in this manner does not require as many values to be initialized by the code but requires an understanding of how and why it does to be able to be used and written.


##### <div align = "center"> Footnotes </div>
[^1]: Formatting for side by side found at [Stack Overflow](https://stackoverflow.com/questions/24319505/how-can-one-display-images-side-by-side-in-a-github-readme-md)
