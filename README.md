# Stock-Analysis

## Purpose of Stock Analysis
The purpose of this Stock Analysis using VBA/Excel is to analyze and refactor the code to loop through the data. This is to make the code is more efficient due to refactoring or if it does not need to be refactored. This includes analyzing to measure performance to see if refactoring causes the data to process faster. 

### Analysis and Challenges
During this module, most of the code work has already been completed. The completed items are listed below:

- Created a variable with single or long data. 
- Wrote if-then statements.
- Used design patterns.
- Used logical and comparison operators.
- Used an index to access the data in an array. 
- Nested loops.
- Reused code.
- Debugged and commented on code.
- Used visual and numeric formatting. 
- Used conditional formatting.
- Measured code performance. 

Understanding what has already been done to the code will help analyze the advantages and disadvantages of refactoring the code for efficiency. 

**<ins>Challenges</ins>**

Measuring code performance and refactoring were the most challenging parts. Understanding what code to use required a lot of research and examples. The good thing is most of the code was already done and everything did not have to be re-written. 


**<ins>Deliverables</ins>**
The following steps were taken to refactor the data. 
   '1a) Create a ticker Index
    For i = 0 To 11
    tickerIndex = tickers(i)

    '1b) Create three output arrays
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single, tickerEndingPrices As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
Worksheets(yearValue).Activate
    tickerVolumes = 0
        
    ''2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount
    
  ' If the next row’s ticker doesn’t match, increase the tickerIndex.
           If Cells(j, 1).Value = tickerIndex Then
           
              '3a) Increase volume for current ticker
              tickerVolumes = tickerVolumes + Cells(j, 8).Value
        
           End If
           
           
        '3b) Write an If-Then statement to check if the current row is the first row with the selected ticker. 
        'If  Then
           If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               tickerStartingPrices = Cells(j, 6).Value
               
          'End If
           End If

        '3c) Write an If-Then statement to check if the current row is the last row with the selected ticker.
        'If  Then
           If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               tickerEndingPrices = Cells(j, 6).Value
               
          'End If
           End If
           
       Next j
       
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

           Worksheets("All Stocks Analysis").Activate
           
           Cells(4 + i, 1).Value = tickerIndex
           Cells(4 + i, 2).Value = tickerVolumes
           Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
    
            'With Range("C4:C15")
                        '.NumberFormat = "0.0%"
                        '.Value = .Value
            'End With
            

   Next i
 
   '5) Added formatting and Ran the code. 
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

2017 Outcome from AllStocksAnalysis
![2017](https://user-images.githubusercontent.com/30300621/175452908-cdbe305c-3aef-491f-b776-b740babd2087.png)

2017 Outcome from VBA Challenge
![VBA_Challenge_2017](https://user-images.githubusercontent.com/30300621/175452536-30415fb5-4032-4fab-b8b0-964be5744744.png)

2018 Outcome from AllStocksAnalysis
![2018](https://user-images.githubusercontent.com/30300621/175452977-e872a227-fc03-4b43-b6da-8898a93b5a5d.png)

2018 Outcome
![VBA_Challenge_2018](https://user-images.githubusercontent.com/30300621/175452599-2e97f1af-2685-47b0-9dfc-4f88ff6b3721.png)


    '6) Comments are added.   
 
 ![image](https://user-images.githubusercontent.com/30300621/175452686-d866f79c-5201-467e-8844-2b478322cd76.png)

## SUMMARY

### Overall Outcome
Refactoring the codes allows the user to view each year without having to search within larger data set. There are some advantages as well as disadvantages of refactoring. Refactoring can be considered as quality check to improve the code and configure the analysis to be user friendly. However, refactoring can be time consuming if someone needs to complete a task quickly. With this analysis, the refactored code did not run quicker than the original code. It is safe to say, it may or may not run quicker.  

    
    
    
    
    
 

  
