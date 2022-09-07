# Analyzing Stocks with VBA
## Project Overview

### Purpose

The purpose of this project was to refactor our Excel VBA code used for analyzing stocks in order to make it more efficient (faster) and to be able to use it with any year (given the data is added to the sheet for that specific year). Using this newly refactored code, anyone should be able to analyze an array of stock tickers to get a visual on how well (or badly) that list of tickers performed for the chosen year. 

### Background

We originally created a macro that was hard coded with the year to be run, but then added an input box that asked what year the user wanted to see analyzed: 

#### Original code:

```

Sub AllStocksAnalysis()


   'Format output sheet on All Stocks Analysis worksheet:
   Worksheets("All Stocks Analysis").Activate
   
   Dim startTime As Single
   
   Dim endTime As Single
   
       
   yearValue = InputBox("What year would you like to run the analysis on?")
   
   
        startTime = Timer
   
   
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   
   
   'Create headers:
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   'Create array of tickers:
   
   Dim tickers(11) As String
   
   'Initialize:
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


   'Create var for starting and ending price:
   Dim startingPrice As Single
   Dim endingPrice As Single

   'Activate Worksheet:

   Worksheets(yearValue).Activate

   'Sheets(yearValue).Activate


   'Count number of rows to loop:
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   'Loop tickers:

   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0

       'loop data rows:
       Worksheets(yearValue).Activate
       'Sheets(yearValue).Activate

       For j = 2 To RowCount

           'Grab total volume for looped ticker:

           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If

           'Grab starting price for looped ticker:

           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           'Grab ending price for looped ticker:
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If

       Next j


       'Output data for looped ticker:

       Worksheets("All Stocks Analysis").Activate

         Cells(4 + i, 1).Value = ticker
         Cells(4 + i, 2).Value = totalVolume
         Cells(4 + i, 3).Value = endingPrice / startingPrice - 1


   Next i

   endTime = Timer

   MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

```



#### Refactored code:

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

### Analysis:

#### Results

Using the refactored code, I was able to cut down the time needed to run the macro through the worksheet:

Refactored code times:

<img width="270" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/58227052/188755397-2b1fe974-5560-4043-a0cf-47f3bc1e39ff.png">

<img width="281" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/58227052/188755425-4f29795f-c969-41fc-9f94-68934326f73e.png">

Original code times:

<img width="268" alt="VBA_Challenge_OG_2017" src="https://user-images.githubusercontent.com/58227052/188756035-af775d6d-1752-4add-8757-d395b9818c76.png">

<img width="271" alt="VBA_Challenge_OG_2018" src="https://user-images.githubusercontent.com/58227052/188756095-9e8a557c-6305-4952-b996-b7cec93a2f50.png">

From this we can see we saved roughly half a second (~0.5) from refactoring the code. I did this by creating tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays and using a new variable tickerIndex to loop through tickers.

### Summary

#### Advantages and Disadvantages of Refactoring Code

Generally, advantages of refactoring code include cleaner, easier to read code for others to modify or read through, and also allows for faster debugging. We also showed here that we can make some changes that allows the code to run much faster and efficient than it did previously. However, sometimes the amount of code is just simply too large to refactor in a reasonable amount of time. Refactoring bits and pieces of the code may be feasible, but can also lead to syntax and other errors, ultimately making the code unusable until corrected.

#### Advantages and Disadvantages of the Original vs Refactoring Code

When comparing the original code with the newly refactored code, the original has a few advantages: It was easier to understand in a step-by-step way, and it was already written. Refactoring the code itself took extra time to accomplish the same task, just in a faster way. Of course, taking that extra time allowed us to streamline the process and make it faster and will ultimately save time in the future, which is always a plus.

###### Disclaimer:

<sub><sup>I used a variety of resources in order to get this code to work properly and cut down on the time it ran. Some of the lines are from my own researches, but there are also some hints taken from other sites and sources that I refactored and used here as well.</sub></sup>
    
 <sub><sup>Yes I made this text tiny just to see if I could.</sub></sup>

