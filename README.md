# Green-Stocks-Analysis-

## PURPOSE:
This project takes a deeper approach on the various ways we edit and refactor codes using the Stock Market Dataset with VBA. The purpose is to loop through all the data at one go in order to collect and analyze the necessary datasets. By doing this, we get to see if the process of refactoring makes the VBA script run faster or not. The goal is to make the code more efficient and easier for future users to read. !

## Results:
I started the coding process by creating 3 new arrays to store performance data for each stock for the “for” loop analysis. These arrays include: 
i.	tickerVolumes(11) to hold volume 
ii.	tickerStartingPrices(11) to hold starting prices 
iii.	tickerEndingPrices(11) to hold ending prices. 
I then create a “tickerIndex” variable to match the performance arrays with the listed ticker() arrays. Once those are created the “Nested For Loop” to complete the analysis. 

Sub AllStocksAnalysisRefactored1()
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
    
    'Activate data worksheet
    
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    tickerIndex = 0
        
    '1b) Create three output arrays
        
        Dim tickerVolumes(11) As Long
        Dim tickerStartingPrices(11) As Single
        Dim tickerEndingPrices(11) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
    
    tickerVolumes(i) = 0
        
    Next i
     
    ''2b) Loop over all the rows in the spreadsheet.
    
    For j = 2 To RowCount

    
        '3a) Increase volume for current ticker
        
        If Cells(j, 1).Value = tickers(tickerIndex) Then
        
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
        If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
        
        tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        
      End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        
        If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
        
        tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
        
        End If
        
        
            '3d Increase the tickerIndex.
          If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
          
            tickerIndex = tickerIndex + 1
    
        End If
    
     Next j
     
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

## 2017 VS 2018:

Green stocks seemed to be more prosperous in 2017 as opposed to 2018, and I would classify this as a decline. Only 2 of the 12 stocks, ENPH & RUN produced a positive return in both years. Although the total volume increases with an additional years from 2017, the returns still report to be low and that could be as a result of the prices. 
<img width="285" alt="All Stocks Analysis 2017" src="https://user-images.githubusercontent.com/104735724/167329370-a4749e9a-3f3b-4637-95be-274869b89904.png">

<img width="286" alt="All Stocks Analysis 2018" src="https://user-images.githubusercontent.com/104735724/167329377-4de88c9d-6fb9-4abd-b3b7-369cd668d458.png">

## EXECUTION TIME: 

The code proved to be very time efficient. The execution time for 2017 improved from 0.1484375 seconds to 0.0859375 seconds, showing a 42% increase. 
Likewise, the execution time for 2018 improved from 0.1171875 seconds to 0.0859375 seconds, showing a 26.67% increase. 

2017:
<img width="297" alt="Original VBA Count 2017" src="https://user-images.githubusercontent.com/104735724/167329511-3d231549-853b-4691-bec4-0ddfc98d344b.png">

<img width="286" alt="VBA_Challenge_2017 png " src="https://user-images.githubusercontent.com/104735724/167329521-ce7a8e23-5d8d-4317-9c5f-98bd0bcc47a6.png">
 
2018:
<img width="286" alt="Original VBA Count 2018" src="https://user-images.githubusercontent.com/104735724/167329556-2c8bea88-3e0d-4846-81ad-89094a94332a.png">

<img width="286" alt="VBA_Challenge_2018 png " src="https://user-images.githubusercontent.com/104735724/167329567-3e781dad-0f03-4b17-bb1d-6357d4769d1d.png">

## ADVANTAGES AND DISADVANTAGES:
Adv: 
Refactoring helps in cleaning up codes and makes it easier for other developers to read. One of the major benefits is that it increases the quality of codes making it easier to conform to coding standards. It makes it more accessible in the future and makes code easier to understand. This process allows us to eliminate complex features, remove blocks of code that may not be necessary anymore and breaks out larger classes and methods into smaller ones. Refactoring also helps us in identifying bugs and errors in codes. This improves the quality of the code and enables the script to run smoother and faster. 

Dis:
A main disadvantage of refactoring is that you are prone to make errors that may disrupt an already working code. This was something I struggled with especially when you forget to save your original code and any other changes you make. Another disadvantage is that it can be time consuming when for new developers who aren’t fully familiar with coding. 




