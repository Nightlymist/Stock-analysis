# Stock-analysis
## Overview of Project
### Background
My friend Steve recently graduated with a financial degree. His parents are very proud of him to the point where he becomes their financial advisor. The parents are really into green companies and want to invest in some. Curently they are looking into DAQO New Energy Corporation. Which is a company that makes the silicon wafers for solar panels. The parents want to only invest money into DAQO New Corporation but Steve want to diversify the portfolio. 

### Purpose

The purpose of this assignment is to determine whether the refactory code was faster than the original code. The original code was to analyze the change in stock over 2017 and 2018 from the start of the year to the end of the year.  

## Results

### Original Code Results

![2017](https://user-images.githubusercontent.com/110945895/188242256-76613e09-3735-4fd8-8fec-f2c7ddf15fc8.png)

Figure 1. 

![2017 (Refined)](https://user-images.githubusercontent.com/110945895/188242270-90f14400-b9f5-4152-ba92-94901ec5a539.png)

Figure 2. 

### Refactory Code Results

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
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, "H").Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row‚Äôs ticker doesn‚Äôt match, increase the tickerIndex.
        'If  Then
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
            

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        
        End If
        'End If
    
    Next i
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i
![2018](https://user-images.githubusercontent.com/110945895/188242280-29290c05-5b2c-49d8-b853-85eefe1df911.png)

Figure 3.

![2018 (Refined)](https://user-images.githubusercontent.com/110945895/188242297-49c3c676-f98b-49bd-97a7-c3696033b24d.png)

Figure 4. 

## Summary

The advantages of refactory code is taking existing code and making it more efficient. This can be either done by using less code which then uses less memory in the computer or by improving the logic of the code to make it easier for other people reading the code to understand the code. The advantage of using the refactory code is using less time to get the results of the code by mili-seconds. 

