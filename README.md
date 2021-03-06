# Stock-Analysis

### 1. Overview of Project:
#### Taking a previously prepared stocks analysis file that summarizes tickers by daily volume and return and refactoring the code to run faster and more efficiently. We have stock data for 12 tickers from 2017 and 2018.

### 2. Results:
#### After refactoring the code the analysis for all stock 2017 and 2018 displayed the exact same results. 
![All-Stocks_2017](https://github.com/maldonado91/Stock-Analysis/blob/main/Resources/All_Stocks_2017.png) ![All-Stocks_2018](https://github.com/maldonado91/Stock-Analysis/blob/main/Resources/All_Stocks_2018.png)
#### However, the run time was much different in both instances. We saw much faster times, therefore, acheiving our goals of enhancing code performance.
#### Here is 2017
Before ![Run_Time2017_Before](https://github.com/maldonado91/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017_Before.PNG) After ![Run2017_Time_After](https://github.com/maldonado91/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.PNG)
#### Here is 2018
Before ![Run_Time2018_Before](https://github.com/maldonado91/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018_Before.PNG) After ![Run_Time2018_After](https://github.com/maldonado91/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

#### Changing the code to run throught the data once was extremely useful. See below 'for loop' used in macro:
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
         If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
         
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
         
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
         End If
    
    Next i
#### Before we ran through the two separate loops to acheive the same output. See below 'for loop' used in macro:
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
        'Loop through the data
        Worksheets(yearValue).Activate
        For j = 2 To rowEnd
            
            If Cells(j, 1).Value = ticker Then
            
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If
            
            'Find the starting price for the current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                
                startingPrice = Cells(j, 6).Value
                
            End If
            
            'Find ending price for the current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                
                endingPrice = Cells(j, 6).Value
                
            End If
            
         Next j
         
         'Outout the data for the current ticker
         Worksheets("All Stock Analysis").Activate
         Cells(4 + i, 1).Value = ticker
         Cells(4 + i, 2).Value = totalVolume
         Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
         
    Next i
#### You can find final project VBA code [here.](https://github.com/maldonado91/Stock-Analysis/blob/main/VBA_Challenge_Complete.vbs)

### 3. Summary:
#### What are the advantages or disadvantages of refactoring code?
The advantages are certainly improved code. We manageed to shrink the amount of code used and made the macro more readable.

#### How do these pros and cons apply to refactoring the original VBA script?
In this particular example we established quicker run times which will allow us the potential to analyze more information and additional years. We were able to leverage arrays to provide our output and we only had to run through the rows of tickers one time. Like the code above shows, we specifically used a for loop to run through our data one time as opposed to once for every ticker in the data set. 

