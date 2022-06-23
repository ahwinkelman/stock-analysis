# stock-analysis
# Stock Analysis With Excel VBA

## Overview of Project
### Purpose
The purpose of the project was to refactor or clean up code for 2 years worth of data for 12 different stocks.  By cleaning up the code from the original workbook, we were able to run a larger analysis to display a more accurate performance of a companies stock performance by increase the sample size.  The main goal of this challenge was to increase the speed of the code by utilizing nexted 'For Loops'.

## Results
#### 2017 Stock Performance
![VBA_Challenge_2017_performance](https://user-images.githubusercontent.com/107078763/175383764-4c87d8f6-08c1-4660-821d-2d63648fa5cc.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/107078763/175384078-247ef79b-35e3-41c9-bef5-ecb7baac5b1e.png)
#### 2018 Stock Performance
![VBA_Challenge_2018_performance](https://user-images.githubusercontent.com/107078763/175384269-2bd11ecd-e7ec-4415-9644-e98ea739c562.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/107078763/175384387-5ded9087-5fd5-474e-b1e8-49ff271d4777.png)
### Analysis
The analysis is well described with screenshots and code (4 pt).
in 2017, ~92% of companies that were analyzed had positive returns on their stocks performance. However, when we looked at the following year, 2018, ~17 of companies we able to have back to back positive returns. 

For a user to perform this yearly analysis, we needed to copy the starter code to provide the input box, ticker array, chart headers, and to activate the worksheet based on what year the user selected.   Below is the code I refactored based on the comment structure provide.

  '1a) Create a ticker Index
    tickerIndex = 0
    

    '1b) Create three output arrays - 12 used because want the loop to run 12 times for number of unique tickers (0 to 11)
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'outside loop run after inside loop, completed, once completed it will increase i +1 to next tickers array
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
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
        
            
            
            
            'End If
        
            '3c) check if the current row is the last row with the selected ticker
             'If the next row’s ticker doesn’t match, increase the tickerIndex.
            'If  Then
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
             tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
             End If
            
            

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
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

## Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
### Pros and Cons of Refactoring Code

### The Advantages of Refactoring Stock Analysis
