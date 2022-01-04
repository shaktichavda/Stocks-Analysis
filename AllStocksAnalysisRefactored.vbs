Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("E1").Value = "All Stocks (" + yearValue + ") - Refactored"
    
    'Create a header row
    Cells(3, 5).Value = "Ticker"
    Cells(3, 6).Value = "Total Daily Volume"
    Cells(3, 7).Value = "Return"

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
    Start = 2
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For init = 0 To 11
        tickerVolumes(init) = 0
        tickerStartingPrices(init) = 0
        tickerEndingPrices(init) = 0
    Next init

    ''2b) Loop over all the rows in the spreadsheet.
    For m = Start To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(m, 1).Value = tickers(tickerIndex) Then
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(m, 8).Value
        
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'A2 = AY            and       A1 <> AY         and A2 = AY           =  tickers(0) = AY
        If Cells(m, 1).Value <> Cells(m - 1, 1).Value And Cells(m, 1).Value = tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(m, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(m, 1).Value <> Cells(m + 1, 1).Value And Cells(m, 1).Value = tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(m, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
        
    Next m
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        
    For o = 0 To 11
        
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + o, 5).Value = tickers(o)
    Cells(4 + o, 6).Value = tickerVolumes(o)
    Cells(4 + o, 7).Value = tickerEndingPrices(o) / tickerStartingPrices(o) - 1
    
    Next o
 
    endTime = Timer
    MsgBox "AllStocksAnalysisRefactored code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("E3:G3").Font.Color = vbBlue
    Range("E3:G3").Font.FontStyle = "Bold"
    Range("E3:G3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("F4:F15").NumberFormat = "#,##0"
    Range("G4:G15").NumberFormat = "0.00%"
    Columns("F").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 7) > 0 Then
            
            Cells(i, 7).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 7).Interior.Color = vbRed
            
        End If
        
    Next i

End Sub
