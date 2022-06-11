# stock_analysis

# Links






# Overview of Project

## Purpose and Background

	We have been helping Steve in his efforts to analyze stock performance for his parents. Throughout Module 2, we created analyses for one stock, and then multiple stocks. Now, Steve wants us to analyze stock performance for thousands of stocks in 2017 and 2018. Given this increase in data that needs to be processed, we needed to refactor the original code so it can run faster, and more efficiently. Refactoring is critical when attempting to uncover the best way to achieve your goal. 

# Results

## Analysis is well described with code
  
### Code for yearValueAnalysis subroutine

```
Sub yearValueAnalysis()

Dim startTime As Single
Dim endTime As Single

yearValue = InputBox("What year would you like to run the analysis on?")

startTime = Timer

   '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   
   
   'Range("A1").Value = "All Stocks (2018)" - REPLACE WITH yearValue for analyses on any year
   
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
   'Worksheets("2018").Activate
   
   Sheets(yearValue).Activate
   
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       'Worksheets("2018").Activate
       
       Sheets(yearValue).Activate
       
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
   
endTime = Timer
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

### Code for AllStocksAnalysisRefactored subroutine

```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
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
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'setting the values in the array to zeroes
    
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        'use the tickerIndex variable as reference for tickerVolumes
        
        
       tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'want to make sure the previous row is not equal to the current row - if they are not equal, we are on the right row
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            'tickerStartingPrices uses the Close column because the closing price of the previous day is the starting price of the next day
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = (tickerIndex + 1)
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        'the division in this line generates the rate of change, subtracting 1 gives us the percentage change
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
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
```

## Analysis is well described with screenshots

	The project goal for Steve was accomplished. One of our goals was to make sure the new subroutine ran faster, please see the screenshots below for the run-time:
	
### 2017 Analysis Before Refactoring

![performance of 2017 analysis before refactoring](https://github.com/jmalauss/stock_analysis/blob/c1fe8ab9093fb8504474b8c461e104cfb402768a/Sub_yearValueAnalysis_2017.png)

### 2018 Analysis Before Refactoring

![performance of 2018 analysis before refactoring](https://github.com/jmalauss/stock_analysis/blob/c1fe8ab9093fb8504474b8c461e104cfb402768a/Sub_yearValueAnalysis_2018.png)

### 2017 Analysis After Refactoring

![performance of 2017 analysis after refactoring](https://github.com/jmalauss/stock_analysis/blob/c1fe8ab9093fb8504474b8c461e104cfb402768a/Sub_AllStocksAnalysis_2017.png)

### 2018 Analysis After Refactoring

![performance of 2018 analysis after refactoring](https://github.com/jmalauss/stock_analysis/blob/c1fe8ab9093fb8504474b8c461e104cfb402768a/Sub_AllStocksAnalysis_2018.png)

# Summary

## Advantages and Disadvantages of refactoring code in general

## How do these pros and cons apply to refactoring the original VBA script?
