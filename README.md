# Module 2- VBA - All Stock Analysis
Link to the main file is available here: [Module-2 VBA Challenge - Stock Analysis File](./VBA_Challenge.xlsm)

## Overview of Project
### Purpose
The main purpose of this project was to refactor a VBA code that was already created as part of Module 2 solution. The code was already collecting and analyzing 2017 and 2018 stocks data and summarizing the volume and % return for each of the 12 tickers.  The goal of refactoring VBA code is to make the code efficient and run the script faster. 

## Results
### Analysis
 Below is the main code which was refactored as part of this project.  In this code, three output arrays were created. The new variable "tickerIndex" was created and was used to access the correct index across the four different arrays.  Comments in the code further explain reason for adding/editing each line. 

    
    '1a) Create a ticker Index
    
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long  'Creating Array to store 12 different volumes - each per one ticker
    Dim tickerStartingPrices(12) As String 'Creating Array to store 12 different Start Price - each per one ticker
    Dim tickerEndingPrices(12) As String 'Creating Array to store 12 different End Price - each per one ticker
    
        
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
    
    tickerVolumes(i) = 0
    
    Next i
            
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        
        '3c) check if the current row is the last row with the selected ticker
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
  
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
       
        Worksheets("All Stocks Analysis").Activate  ' print the values of results for all 3 fields on All Stock tab
        
        Cells(4 + i, 1).Value = tickers(i) 'this will print value for tikcers(0), (1).... till (11)
        Cells(4 + i, 2).Value = tickerVolumes(i) 'this will print value for tickerVolumes(0), (1).... till (11)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1 'this will print the % return for each ticker
            
    Next i

## Summary
### Advantages and disadvantages of refactoring code in general
•	1.Simplified support and code updates. Clean code is much easier to debug and update.
•	2.Reduced complexity for easier to understand and read by other users. 
•	3.Saved time and money in the future.
•	4.Easier to maintain and increase scalability

### The Advantages of Refactoring Stock Analysis
The run time for macro went down significantly. The original code took around 0.93 seconds and 0.98 seconds for 2017 and 2018 respectively. While after refactoring, it went down to 0.17 and 0.18 seconds for 2017 and 2018 respectively. 

Below are the screenshots which show the run time for 2017 & 2018 analysis.

![2017 Screenshot](./VBA_Challenge_2017.PNG)

![2018 Screenshot](./VBA_Challenge_2018.PNG)
