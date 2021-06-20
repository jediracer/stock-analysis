#stock-analysis
Performing analysis on stock data from 2017 and 2018
## Overview of Project
### Purpose
- Create a VBA subroutine that allows the user to enter the year of data to be analyzed.  
Then collect and summarize stock data based on user input.  Format the data to make it easier to read and use 
conditionals to color code results. Finally, add a timer and display the run time of the subroutine. 
- Refactor the subroutine performing the same tasks and display the run time of the refactored subroutine for 
comparison with the first subroutine.
## Results
### Stock Performance
- The over stock performance was better in 2017 vs. 2018. Only 1 of the tickers in 2017 did not have a positive return.
Of the 12 tickers only 2 had a positve return in 2018.  The over all best performing ticker is ENPH in 2017 and 2018 combined. 
### Script Comparison
- The original subroutine used a single array and nested for loops.  
```
'2) Initialize array of all tickers
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
        
'3a) Initialize variables for starting price and ending price
    Dim startingPrice As Double
    Dim endingPrice As Double
    
'3b) Activate data worksheet
    Worksheets("2018").Activate
    
'3c) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
'4) Loop through tickers
    For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
'5) loop through rows in the data
        Worksheets("2018").Activate
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
```
- The refactored subroutine utilized 4 arrays and 3 for loops to intialize, collect and display the data.
'''
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
    
    'Activate data worksheet using the year collected from prompt to select the corresponding worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11
        tickerVolumes(tickerIndex) = 0
    Next j
    
    'Start tickerIndex with first position of array
    tickerIndex = 0
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        'set ticker as string name based on array position
        ticker = tickers(tickerIndex)
        
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = ticker Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next row ticker doesn't match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
        
    Next i

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For outRow = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        
        Cells(outRow + 4, 1).Value = tickers(outRow)
        Cells(outRow + 4, 2).Value = tickerVolumes(outRow)
        Cells(outRow + 4, 3).Value = tickerEndingPrices(outRow) / tickerStartingPrices(outRow) - 1
        
    Next outRow
'''
- The refactored subroutine ran significantly faster for both the 2017 and the 2018 data
  - ![2017 Results and Run Time]


The refactored subroutine used