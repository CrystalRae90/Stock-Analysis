# Stock-Analysis

## Overview
The purpose of this project was to take the code written in the module and refactor it to allow for more efficient use. The data provided was for differnt tickers over the time, showing the volume traded and price data for a given day. Using the code, we were attempting to anlayze each ticker over the course of a year, looking at the performance as well as the starting/ending prices. 

## Results: Using images and examples of your code, compare the stock performance between 2018 and 2018, as well as the execution times of the original script and the refactored script.

While the code written in the module was sufficient for the exercise, it would not have been efficient if the data size had increased. In order to look at the difference between the two run times, it is important to deciher how each code was run. To be able to differentiate the two codes, the module code will be referred to "2018 Script" an the refactored code will be referred to as "2018 Refactored Script."

### 2018 Script 
The 2018 Script used a nested for loop to loop through both the tickers, and then each row to determine the values for Total Volume, Starting Price and Endinging Price. 

The initial for loop shown below shows that we want to loop through the all 12 tickers, starting 0 as all Tickers for i had previously been initialized for the tickers array. In addition, we set the Total Volume to zero each time i looped through. 

        For i = 0 To 11

        ticker = Tickers(i)
    
        TotalVolume = 0

 We then nested another for loop within to loop through the rows where j is the row number. For each ticker, the goal was to loop through the rows and determine if that row belonged to that ticker based on value i. If the ticker was the same then, the total volume would be added to the totalvolume variable. For the prices, it determined it based on the previous and next rows in addition to the current rows to set prices.         


            For j = 2 To rowend
    
                If Cells(j, 1).Value = ticker Then
            
                    TotalVolume = TotalVolume + Cells(j, 8).Value
                
                End If
                
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                
                    StartingPrice = Cells(j, 6).Value
                      
                End If
                
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                
                    EndingPrice = Cells(j, 6).Value
            
            End If
       
            Next j

When running this method and setting the values once the loops were complete, we were able to achieve the results to show the below. Total time to run the 2018 Script was 2.1875 seconds. 

![All_Stocks_Results_2018.png](Resources/All_Stocks_Results_2018.png)                           

<kbd>![VBA_Challenge_2018.png](Resources/VBA_Challenge_2018.PNG)<kbd>

## 2018 Refactored Script

Unlike the 2018 Script, the 2018 Refactored Script uses multiple arrays to increase the efficiency of the code. The Ticker Volumes, Starting and Ending Prices were previous set as indivual variables. In this new script, we have set them to also be arrays similar to tickers. 

      Dim TickerVolumes(12) As Long
      Dim TickerStartingPrices(12) As Single
      Dim TickerEndingPrices(12) As Single
      
Similar to 2018 Script, a for loop was written 

      For i = 0 To 11

          TickerVolumes(i) = 0

      Next i



        For i = 2 To RowCount
    
'3a) Increase volume for current ticker
'Take current TickerVolume for that tickerindex and increase it by new value

     TickerVolumes(TickerIndex) = TickerVolumes(TickerIndex) + Cells(i, 8).Value

'3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then

'If Cell in row above is not equal to current ticker based on current ticker index, then remember starting price
        
           
        If Cells(i - 1, 1).Value <> Tickers(TickerIndex) Then
        
        TickerStartingPrices(TickerIndex) = Cells(i, 6).Value
        
        'Debug.Print (TickerStartingPrices(TickerIndex))
        
        End If
        
'3c) check if the current row is the last row with the selected ticker
'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
'If  Then
        
'If Cell in row below is not equal to current ticker based on current ticker index, then remember ending price

         If Cells(i + 1, 1).Value <> Tickers(TickerIndex) Then
        
         TickerEndingPrices(TickerIndex) = Cells(i, 6).Value
         
         'Debug.Print ("Ending If Loop")

 '3d Increase the tickerIndex.
        
         TickerIndex = TickerIndex + 1
                 
         End If 

## Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?

