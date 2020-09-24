Attribute VB_Name = "Module7"
Sub AllStocksAnalysisRefactored()
    
'Get the Start and End Time

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    
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
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    Dim tickerVolumes(12) As Long
        
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
        
    For tickerIndex = 0 To 11
        tickerVolumes(tickerIndex) = 0

    ' If the next row's ticker doesn't match, increase the tickerIndex.
        If tickers(tickerIndex + 1) <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1

        End If
       
    Next tickerIndex

  
  
    '2b) Loop over all the rows in the spreadsheet.

    Worksheets(yearValue).Activate
    For tickerIndex = 0 To 11
        '3a) Increase volume for current ticker
        For i = 2 To RowCount
            If Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
            End If
        
              
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        'End If
            End If
        
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            
        'End If
            End If
        
        Next i
    
    Next tickerIndex

    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    Worksheets("AllStocksAnalysisRefactored").Activate
    Dim Yearcolumn As Integer
    
    For tickerIndex = 0 To 11
        If yearValue = 2017 Then
            Yearcolumn = 1
        Else
            Yearcolumn = 10
        End If
        
    '5.Format the output sheet on All Stocks Analysis worksheet
    Worksheets("AllStocksAnalysisRefactored").Activate
    
    Cells(1, Yearcolumn).Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, Yearcolumn).Value = "Ticker"
    Cells(3, Yearcolumn + 1).Value = "Total Daily Volume"
    Cells(3, Yearcolumn + 2).Value = "Return"
        
        
        
        Worksheets("AllStocksAnalysisRefactored").Activate
        Cells(4 + tickerIndex, Yearcolumn).Value = tickers(tickerIndex)
        Cells(4 + tickerIndex, Yearcolumn + 1).Value = tickerVolumes(tickerIndex)
        Cells(4 + tickerIndex, Yearcolumn + 2).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
     
    'Formatting
    Worksheets("AllStocksAnalysisRefactored").Activate
    Range(Cells(3, Yearcolumn), Cells(3, Yearcolumn + 2)).Font.FontStyle = "Bold Italic"
    Range(Cells(3, Yearcolumn), Cells(3, Yearcolumn + 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range(Cells(4, Yearcolumn + 1), Cells(15, Yearcolumn + 1)).NumberFormat = "#,##0"
    Range(Cells(4, Yearcolumn + 2), Cells(15, Yearcolumn + 2)).NumberFormat = "0.0%"
    Columns(Yearcolumn + 1).EntireColumn.AutoFit
  
  Next tickerIndex
  
  
  
    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, Yearcolumn + 2) > 0 Then
            
            Cells(i, Yearcolumn + 2).Interior.Color = vbGreen
            
        Else
        
            Cells(i, Yearcolumn + 2).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
