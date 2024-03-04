Attribute VB_Name = "Module1"
Sub Test()

For Each ws In Worksheets

    Dim Worksheet As String
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
     ' Add a Columns for the retrievals
    ws.Range("I2:L2").EntireColumn.Insert
    
    Dim TotalVolume As Double
    TotalVolume = 0

    Dim TableRow As Integer
    TableRow = 2
    
     ' Add the word Column Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    

     ' Variable designations for row values
        
  For i = 2 To LastRow
    

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ' Set the Tickers
    Tickerletters = ws.Cells(i, 1).Value
    
    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
    ' Print retieved Data
    ws.Range("I" & TableRow).Value = Tickerletters
    ws.Range("L" & TableRow).Value = TotalVolume
    
    TableRow = TableRow + 1
    
    TotalVolume = 0
    
    Else
    
    ' Sum of volume
    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
    End If
    
  Next i
  
Next ws

End Sub
