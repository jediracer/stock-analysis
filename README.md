#stock-analysis
Performing analysis on stock data from 2017 and 2018
## Overview of Project
### Purpose
- Create a VBA subroutine that allows the user to enter the year of data to be analyzed.  
Then collect and summarize stock data based on user input.  Format the data and use conditionals to color code results making it easier for 
the user to read. Finally, add a timer and display the run time of the subroutine. 
- Refactor the subroutine using different methods to archive the same results.  Also, calculate and display the run time of the refactored subroutine 
for comparison with the first subroutine.
## Results
### Stock Performance
- The overall stock performance was better in 2017 vs. 2018. Only 1 of the tickers in 2017 did not have a positive return.
Of the 12 tickers only 2 had a positive return in 2018.  The over all best performing ticker was ENPH in over the 2 year period. 
### Script Comparison
- The original subroutine used a single array (only to store the tickers), nested for loops and displayed the data as it was collected.  
```
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
        
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    Worksheets("2018").Activate
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
        Worksheets("2018").Activate
        For j = 2 To RowCount
        
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If

            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        Next j
```
- The refactored subroutine utilized 4 arrays and 3 separate for loops to initialize, collect and finally display the data after all 
of the data was collected.  
```
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
    
    Worksheets(yearValue).Activate
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim tickerIndex As Integer
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    For j = 0 To 11
        tickerVolumes(j) = 0
    Next j
    
    tickerIndex = 0
    
    For i = 2 To RowCount
        ticker = tickers(tickerIndex)
        
        If Cells(i, 1).Value = ticker Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
        If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If
        
    Next i

    For tickerIndex = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        
        Cells(tickerIndex + 4, 1).Value = tickers(tickerIndex)
        Cells(tickerIndex + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(tickerIndex + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next tickerIndex
```
- The refactored subroutine ran significantly faster for both the 2017 and the 2018 data (see the results below).
  - ![2017 Results and Run Time](https://github.com/jediracer/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)
  - ![2018 Results and Run Time](https://github.com/jediracer/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)
## Results
### Refactoring Code
- Refactoring code can be advantageous as it allows you to cleanup the code and make the script run more efficiently.  Cleaning up the code
will make it easier for yourself and others to read in the event it needs to be revisited in the future.  Refactored scripts also run
more efficiently, improving performance, and saving time.
- The disadvantage of refactoring code is the time it takes to complete the task.  You are spending time reworking an already functioning script.
### Refactoring of the AllStocksAnalysis vba subroutine
- The refactoring of this script made the process run more quickly, which is a great advantage. Using multiple arrays to store 
the data during the process and then displaying it at the end, made the whole process run more efficiently.
- The disadvantages of refactoring this subroutine were; 1) the time it took to complete the refactoring, and 2) the over length
of the script is slightly longer.