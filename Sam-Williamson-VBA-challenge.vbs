Sub stockticker()

' loop through wosksheets
Dim i As Long
Dim ws As Worksheet

  For Each ws In ThisWorkbook.Worksheets

' Stock Ticker holder
  Dim stock_ticker As String

  ' Quarterly Change Variables
  Dim Quarterly_Open As Double
  Quarterly_Open = Cells(2, 3).Value
  Dim Quarterly_Close As Double
  Quarterly_Close = 0
  Dim Quarterly_Change As Double
  Quarterly_Change = 0
  
  ' Percent Change Variable
  Dim Percent_Change As Double
  Percent_Change = 0
  
  ' Trade Volume Variable
  Dim Trade_Volume As LongLong
  Trade_Volume = 0

  ' Keep track of the location for each stock ticker in the summary table
  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2
  
  ' Loop through all Stock Tickers
  Dim lastRow As Long
  lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

  For i = 2 To lastRow

    ' Check if we are still within the same credit card brand, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the stock ticker name
      stock_ticker = ws.Cells(i, 1).Value

      ' Calculate Quarterly Change
      Quarterly_Close = ws.Cells(i, 6).Value
      Quarterly_Change = Quarterly_Close - Quarterly_Open
        
      ' Calculate Percent Change
      Percent_Change = (Quarterly_Close / Quarterly_Open) - 1
      
      ' Calculate Trade Volume
      Trade_Volume = Trade_Volume + Cells(i, 7).Value

      ' Print the Stock Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = stock_ticker

      ' Print the Quarterly Change to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
      
      ' conditional formatting for Quarterly Change
      If ws.Range("J" & Summary_Table_Row).Value > 0 Then
             ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        ElseIf Quarterly_Change < 0 Then
             ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      End If
      
      ' Print the Percent Change to the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Percent_Change
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
      ' print greatest percent increase and decreaase with ticker in other summary table
        If ws.Range("K" & Summary_Table_Row).Value > ws.Cells(2, 16).Value Then
            ws.Cells(2, 16).Value = ws.Range("K" & Summary_Table_Row).Value
            ws.Cells(2, 16).NumberFormat = "0.00%"
            ws.Cells(2, 15).Value = ws.Range("I" & Summary_Table_Row).Value
                
        ElseIf ws.Range("K" & Summary_Table_Row).Value < ws.Cells(3, 16).Value Then
            ws.Cells(3, 16).Value = ws.Range("K" & Summary_Table_Row).Value
            ws.Cells(3, 16).NumberFormat = "0.00%"
            ws.Cells(3, 15).Value = ws.Range("I" & Summary_Table_Row).Value
            
        End If
      ' Print the Trade Volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Trade_Volume
      
      ' print greatest trade volume with ticker in other summary table
        If ws.Range("L" & Summary_Table_Row).Value > ws.Cells(4, 16).Value Then
            ws.Cells(4, 16).Value = ws.Range("L" & Summary_Table_Row).Value
            ws.Cells(4, 15).Value = ws.Range("I" & Summary_Table_Row).Value
        End If
        
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

      ' Reset the Variables
      Quarterly_Open = ws.Cells(i + 1, 3).Value
      Quarterly_Close = 0
      Quarterly_Change = 0
      Percent_Change = 0
      Trade_Volume = 0

    Else
      
      ' Add to the Trade Volume
      Trade_Volume = Trade_Volume + ws.Cells(i, 7).Value

    End If

  Next i
  
  Next ws

End Sub

