# stocks-analysis

## Project Overview

### Background

Steve has created an Excel spreadsheet with a variety of renewable energy stocks detailing their prices at certain dates, daily volumes, opening and closing prices, as well as their highs and lows. The data are split into two different worksheets by the years 2017 and 2018. The data relating to the return on each stock will give Steve a much better idea of which stocks may yield the highest returns. 

### Purpose

In this project, we are tasked with creating a VBA program for Steve so that he can decide on which green energy stocks are worth investing in for his parents.While it is definitely possible to go through the endless rows of data and manually calculate the return for each individual ticker, it is much more efficient to use Microsoft’s Visual Basic for Applications language, or VBA, to automate the tasks with macros. Through the use of VBA, our goal is to create a subroutine to extract and present the information we need in the blink of an eye.

## Analysis

### Results

After copying the necessary data from the VBS file onto the Visual Basic prompt, I got to work defining the variables and initializing the arrays that I needed for the program. One of the most important components was the ticker index that contained the tickers of all the stocks being analyzed. The next steps involved creating a series of for loops and conditionals that would scan the worksheets for the desired data depending on the year. Unlike the previous code, which required a button to run the program, this refactored code would analyze the data automatically after the year value was inputted and return the results much faster. Attached below is the subroutine used in this program.

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
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
       If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        'End If
        End If
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        'End If
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
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

## Summary

### Pros and cons of refactoring in general

In general, refactored code runs much more smoothly and efficiently, allowing one to process information faster with fewer errors. In this sense, refactoring is like taking a direct flight to a destination as opposed to going through multiple connections. Not only is it a much more efficient process, but it also has much fewer points of failure. The disadvantages of refactoring may be that it may be difficult and time consuming to rebug long lines of code and this process may require more advanced knowledge of VBA.

### Advantages and disadvantages of the refactored script

Whereas the original code that required the use of a button had a run time of approximately 0.6 seconds, the refactored code ran in about .12 seconds, which is about 5 times faster. Although this time difference may seem insignificant to us, if Steve were analyzing a much larger amount of data, refactoring his code can result in a major increase in efficiency, leading to higher customer satisfaction for his company. The disadvantage of this is that it requires high levels of knowledge to properly utilize the tools required to refactor the code without introducing more bugs. Steve may have to hire a more advanced programmer. If the time difference is only about a fraction of a second, this would likely not be worth the cost.
