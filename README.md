# Stock-analysis
## Overview of Project
### Background
My friend Steve recently graduated with a financial degree. His parents are very proud of him to the point where he becomes their financial advisor. The parents are really into green companies and want to invest in some. Curently they are looking into DAQO New Energy Corporation. Which is a company that makes the silicon wafers for solar panels. The parents want to only invest money into DAQO New Corporation but Steve want to diversify the portfolio. 

### Purpose
The purpose of this assignment is to determine whether the refactory code was faster than the original code. The original code was to analyze the change in stock over 2017 and 2018 from the start of the year to the end of the year.  

## Results
### Original Code Results
As you can see below, this is the original code of renewable energy companies including DAQO New Energy corporation analysis. In the file, you can see a button to create analysis for 2017 and 2018 from the prompt entering the year that you want the analysis for the selected companies. For 2017, only one company that decreased which is TERP from the beginning of the year to the end of the year. As shown below in figure 1, the time it took to run the code below was 2.74 seconds. For 2018, only two companies increased from the start of the year to the end of the year which are ENPH and RUN. As shown below in figure 2, the code shown below took 1 second. 
   '1) Format the output sheet on All Stocks Analysis worksheet
   
    Sheets("All Stocks Analysis").Activate
    
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
   
   Worksheets("2018").Activate
   
   '3c) Get the number of rows to loop over
   
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   
   For i = 0 To 11
   
       ticker = tickers(i)
       
       totalVolume = 0
       
   '5) loop through rows in the data
   
       Sheets("2018").Activate
       
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
   
![2017](https://user-images.githubusercontent.com/110945895/188242256-76613e09-3735-4fd8-8fec-f2c7ddf15fc8.png)

Figure 1. Shown above is the time taken to calculate the results of 12 companies that is of interested for stock investment. The year that was selected for data calculation is 2017. The original code was used to do analysis. 

![2018](https://user-images.githubusercontent.com/110945895/188242280-29290c05-5b2c-49d8-b853-85eefe1df911.png)

Figure 2. The time shown in figure is the time taken to calculate the results of 12 companies different in stock from beginning of year to the end of the year. The year that was selected for the calculation is 2018. The code used for calculation is the original code. 

### Refactory Code
As shown below, the refactory code of the code above trying to improve the time of code to show results from the original code. The code was originized more so that it is easy for the readers to understand the coding. The ticker index was added to the original code. The two figure below are the refined of both years of results 2017 and 2018. 

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

![2017 (Refined)](https://user-images.githubusercontent.com/110945895/188242270-90f14400-b9f5-4152-ba92-94901ec5a539.png)

Figure 3.

![2018 (Refined)](https://user-images.githubusercontent.com/110945895/188242297-49c3c676-f98b-49bd-97a7-c3696033b24d.png)

Figure 4. 

## Summary
The advantages of refactory code is taking existing code and making it more efficient. This can be either done by using less code which then uses less memory in the computer or by improving the logic of the code to make it easier for other people reading the code to understand the code. The advantage of using the refactory code is using less time to get the results of the code by mili-seconds. 
