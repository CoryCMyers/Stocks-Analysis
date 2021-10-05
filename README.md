# Stocks-Analysis
Repository for the Stocks Analysis Project 

# Analysis Summary

## Overview of Project

  The purpose of this analysis is to summarize and then contrast the performance of the listed Tickers over the course of a year, as well as displaying the overall return on profit for that stock over the course of that year. This Project was made in a way that so long as the Data entered into the Workbook in the same manner and the same naming conventions followed on the name of individual sheet names it will continue to function moving forward.
  
## Overview of Results

#### 2017
When looking at the results of the 2017 Data it can be seen that overall for the year the stocks selected to be followed for this exercise had an overall positive trend overall. From the information that you provided for the basic creation of the spreadsheet these stocks would seem to be a solid choice of investment with only one TERP showing a neagative return for the year as we see here. 

[2017 Stocks Analysis Results](https://github.com/CoryCMyers/Stocks-Analysis/blob/main/VBA_Challenge_2017.png) ![2017 Stocks Analysis Results](https://github.com/CoryCMyers/Stocks-Analysis/blob/main/VBA_Challenge_2017.png) 

#### 2018

However, when further information is provided and the results for the year of 2018 are also added this story changes. 

[2018 Stocks Analysis Results](https://github.com/CoryCMyers/Stocks-Analysis/blob/main/VBA_Challenge_2018.png) ![2018 Stocks Analysis Results](https://github.com/CoryCMyers/Stocks-Analysis/blob/main/VBA_Challenge_2018.png)

When we consult the information from the analysis of the 2018 stock ticker history the situation seems to be entirely different. From all of the stocks that were covered as profitable in the previous year only one of them is still profitable in the most recent data set provided. While this stock did increase exponentially in value over the course of the year, the other stock that remained profitable also took a loss from the previous year while the remained all took a sharp decline.

### Conclusions

Without access to further years of information to show any sort of trend to see if there is an ebb and flow to these tickers to determine if the current downturn is a normal part of the stocks price cycle, making it a good time to invest none of these tickers have shown to be educated choices for investment at the time of this analysis.

## Execution Times Comparisons

#### 2017

The runtime for the original and refactored Macros can be found below.

'Formatting found at [Stack Overflow](https://stackoverflow.com/questions/24319505/how-can-one-display-images-side-by-side-in-a-github-readme-md)

2017 Original | 2017 Refactored 
:-------------------------:|:-------------------------:
![2017 Runtime Original](https://github.com/CoryCMyers/Stocks-Analysis/blob/CoryCMyers-patch-1-workingOnReadme/2017Base.PNG)  |  ![2017 Runtime Refactored](https://github.com/CoryCMyers/Stocks-Analysis/blob/main/VBA_Challenge_2017.png)

2017 Original Code | 2017 Refactored Code 
:-------------------------:|:-------------------------:
  Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer
       
   '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
   Dim tickers(11) As String
   tickers(0) = "AY"
   tickers(1) = "CSIQ"
   tickers(2) = "DQ"
   tickers(3) = "ENPH"
   tickers(4) = "FSLR"
   tickers(5) = "HASI"
   tickers(6) = "JKS"
   tickers(7) = "RUN"
   tickers(8) = "SEDG"
   tickers(9) = "SPWR"
   tickers(10) = "TERP"
   tickers(11) = "VSLR"
   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b) Activate data worksheet
   Worksheets(yearValue).Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i

   dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub | Sub AllStocksAnalysisRefactored()
   Dim startTime As Single
   Dim endTime  As Singl
   yearValue = InputBox("What year would you like to run the analysis on?")
   startTime = Timer
   
   'Format the output sheet on All Stocks Analysis worksheet
   Sheets("All Stocks Analysis").Activate
   
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"
   'Initialize array of all tickers
   Dim tickers(12) As String
   
   tickers(0) = "AY"
   tickers(1) = "CSIQ"
   tickers(2) = "DQ"
   tickers(3) = "ENPH"
   tickers(4) = "FSLR"
   tickers(5) = "HASI"
   tickers(6) = "JKS"
   tickers(7) = "RUN"
   tickers(8) = "SEDG"
   tickers(9) = "SPWR"
   tickers(10) = "TERP"
   tickers(11) = "VSLR"
   
   'Activate data worksheet
   Sheets(yearValue).Activate
   
   'Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row
   
   '1a) Create a ticker Index
   tickerIndex = 0
   '1b) Create three output arrays
   Dim tickerVolumes(12) As Long
   Dim tickerStartingPrices(12) As Single
   Dim tickerEndingPrices(12) As Single
   ''2a) Create a for loop to initialize the tickerVolumes to zero.
   For i = 0 To 11
       tickerVolumes(i) = 0
   Next i
       
   '2b) Loop over all the rows in the spreadsheet.
   For i = 2 To RowCount
   
       '3a) Increase volume for current ticker
       tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
       
       '3b) Check if the current row is the first row with the selected tickerIndex.
       If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
           
           tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
           
       End If
       
       '3c) check if the current row is the last row with the selected ticker
       'If the next row’s ticker doesn’t match, increase the tickerIndex.
       If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
           
           tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
           '3d Increase the tickerIndex.
           tickerIndex = tickerIndex + 1
           
       End If
   
   Next i
   
   '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
   For i = 0 To 11
       
       Sheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
       Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
       
   Next i
   
   'Formatting
   Sheets("All Stocks Analysis").Activate
   Range("A3:C3").Font.FontStyle = "Bold"
   Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
   Range("B4:B15").NumberFormat = "#,##0"
   Range("C4:C15").NumberFormat = "0.0%"
   Columns("B").AutoFit
   dataRowStart = 4
   dataRowEnd = 15
   For i = dataRowStart To dataRowEnd
       
       If Cells(i, 3) > 0 Then
           
           Cells(i, 3).Interior.Color = vbGreen
           
       Else
       
           Cells(i, 3).Interior.Color = vbRed
           
       End If
       
   Next i
 
   endTime = Timer
   MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
End Sub


#### 2018

2018 Original             |  2018 Refactored
:-------------------------:|:-------------------------:
[2018 Runtime Original](https://github.com/CoryCMyers/Stocks-Analysis/blob/CoryCMyers-patch-1-workingOnReadme/2018Base.PNG)  |  ![2018 Runtime Refactored](https://github.com/CoryCMyers/Stocks-Analysis/blob/CoryCMyers-patch-1-workingOnReadme/VBA_Challenge_2018.png)

#Summary

##Refactored Code Pro/Con


###How do these Apply to this Code?
  
  

