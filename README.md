# Analyzing Stocks with VBA
## Project Overview

### Purpose

The purpose of this project was to refactor a macro used for analyzing stocks (in Excel) in order to make it more efficient (faster) and to able to use it with any year (given the data is added to the sheet for that specific year). 

### Results

Using the refactored code, I was able to cut down the time needed to run the macro through the worksheet:

Orignal code time:

- OG TIME 2017
- OG TIME 2018

Refactored Code time:

- NEW TIME 2017
- NEW TIME 2018

Refactored code:

```
'Create variable, set to zero:
    
    tickerIndex = 0
    
    'Output arrays:
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    'Loop to initialize the tickerVolumes equal zero:
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    'Looping over all rows:
    For i = 2 To RowCount
    
        'Increases volume for current ticker (taken from Module 2 Challenge page):
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        'Checks if the current row is the first row using the current ticker value:
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'Checks if the current row is the last row using the current ticker value.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
       End If

            'Increases tickerIndex by one.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
    
    'Loops arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
    
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```

#### Analysis:

### Summary

#### Advantages and Disadvantages of Refactoring the Code

#### 
