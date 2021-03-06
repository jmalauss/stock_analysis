# stock_analysis

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
         'If the next row???s ticker doesn???t match, increase the tickerIndex.
        
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

## Explanation of the code above

The refactored code differs from the original code in a few ways: Mainly, the refactored code creates three arrays that correspond with the ticker array from the first code. These additional arrays and variables (tickerIndex) allow us to adjust the nested loop in the initial code so it does more, with less lines of code than we would have needed to accomplish the same analysis with the original code structure. In other words, we created a "tickerIndex" variable used as an input for the tickerVolumes array. After setting the tickerIndex to 0, we were able to ensure that tickerIndex could be used as an input to yield the corresponding tickerVolumes. 

## Analysis is well described with screenshots

The project goal for Steve was accomplished. One of our goals was to make sure the new subroutine ran faster, please see the screenshots below for the run-time:
	
### 2017 Analysis Before Refactoring

![performance of 2017 analysis before refactoring](https://github.com/jmalauss/stock_analysis/blob/c1fe8ab9093fb8504474b8c461e104cfb402768a/Sub_yearValueAnalysis_2017.png)

### 2018 Analysis Before Refactoring

![performance of 2018 analysis before refactoring](https://github.com/jmalauss/stock_analysis/blob/c1fe8ab9093fb8504474b8c461e104cfb402768a/Sub_yearValueAnalysis_2018.png)

### 2017 Analysis After Refactoring

![performance of 2017 analysis after refactoring](https://github.com/jmalauss/stock_analysis/blob/c1fe8ab9093fb8504474b8c461e104cfb402768a/Sub_AllStocksAnalysis_2017.png)

**Refactored code was faster by 0.925781 seconds!**

### 2018 Analysis After Refactoring

![performance of 2018 analysis after refactoring](https://github.com/jmalauss/stock_analysis/blob/c1fe8ab9093fb8504474b8c461e104cfb402768a/Sub_AllStocksAnalysis_2018.png)

**Refactored code was faster by 0.9843755 seconds!**

# Summary

[In order to best answer the following questions, I needed to do some research, here is where I found the following information](https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software#:~:text=Advantages%3A%201.%20Refactoring%20is%20a%20really%20good%20weapon,It%27s%20risky%20when%20the%20application%20is%20big%202.)

## Advantages and Disadvantages of refactoring code in general

### General Advantages

The main advantages to refactoring code is that it can result in a simplified code, which will make the code easier to adjust in the future, and easier for others to understand. Additionally, it will be easier to identify bugs in your code and will allow the code to run faster, and use less memory. 

### General Disadvantages

The main disadvantages based on my research, arise when working on the code is dependent on limited resources. In other words, refactoring may require additional time or budget, so you may not be able to refactor if you do not have the time and budget to allocate to refactoring. Additionally, refactoring may jeopardize the functionality of the code if the code you are refactoring is something you don't understand completely. You can get around this last disadvantage by saving copies of code, and refactoring the copies so you can always revert back to the original, functional code.

## How do these pros and cons apply to refactoring the original VBA script?

Application of pros and cons to Module 2 Challenge:

Mainly, the refactored code ran faster and more efficiently. We know this because the refactored subrountine is able to process more data, in less time than the original subrountine. Another application relates to the disadvantages of refactoring code: I found myself saving and copying code that was functional. In other words, after running code and seeing that it successfully accomplished the task, I copy and pasted that code into NotePad, or commented out the code that I knew worked. This allowed me to revert any changes I made that jeopardized the functionality of the subroutine. Additionally, when I first ran my refactored code, there was an error message. It took me to the line of code where I could make the adjustment. It was very clear to notice that "tickerVolume" should have been "tickerVolumes" - the refactoring of the code made it easy for me to identify the error in my code. 

