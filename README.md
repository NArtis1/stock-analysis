# Green Stock Analysis
## Overview

Steve and his parents asked for help analyzing a green stock for his parents to invest in. To help them out , I used VBA in excel to find the annual return percentage and the total daily volume. Steve parents wanted to know how the stock I was analysis performance compared to the other 11 stocks listed. The information I recieved allowed me to help them make an informed decision on which stock was the best option to invest in.

Purpose

The purpose of the project was to figure out a way to efficently look at a variety of stocks using VBA within Excel. After completing my first analysis, I refactored my code to make it more effiecent.

## Results

I combed through the data for a years worth of stock data from 2017 and 2018. The data within the spreadsheet contained important information on the daily performance of 12 stocks. 

### Orignal Code
In order to get the neccessary data points needed to get the total daily volume and annual return, I used an array to look through the stock spreadsheat and get the neccessary information for each stock. 

  Sub AllStocksAnalysis()

      Dim startTime As Single
      Dim endTime As Single

  'Format the output sheet on All Stocks Analysis worksheet

  Sheets("All Stocks Analysis").Activate

  yearValue = InputBox("What year would you like to run the analysis on?")

  startTime = Timer

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

  'Initialize varaibles for starting priceand ending price

  Dim startinPrice As Single
  Dim endingPrice As Single

  'Activate data worksheet

  Sheets(yearValue).Activate

  'Get the number of rows to loop over

  RowCount = Cells(Rows.Count, "A").End(xlUp).Row

  'Loop through tickers

   For i = 0 To 11

      ticker = tickers(i)
      totalVolume = 0

      'Loop through rows in the data
      Sheets(yearValue).Activate

          For j = 2 To RowCount

              'Get total volume rows in the data
              If Cells(j, 1).Value = ticker Then

                   totalVolume = totalVolume + Cells(j, 8).Value

              End If
              'get starting price for current ticker
              If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                  startingPrice = Cells(j, 6).Value

              End If
              'get ending price for current ticker
              If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                  endingPrice = Cells(j, 6).Value

              End If

          Next j
          'output data for current ticker
          Worksheets("All Stocks Analysis").Activate
          Cells(4 + i, 1).Value = ticker
          Cells(4 + i, 2).Value = totalVolume
          Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

  Next i

  endTime = Timer
  MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

  End Sub
  
  
  
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

### Refactored Code

### Orignal Code


## Summary
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?



