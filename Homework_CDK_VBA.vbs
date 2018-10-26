Sub StockTicker()

' initial variable for the ticker and active worksheets
  Dim Ticker As String
  Dim lastrow As Long
  Dim WS_Count As Integer
  Dim R As Integer
  
' initial variable for total volume (double or integer)
  Dim price_open As Double
  price_open = 0
  Dim price_close As Double
  price_close = 0
  Dim total_volume As Double
  total_volume = 0
  Dim vol_change_tot As Double
  price_close = 0
  Dim percentage_chg As Double
  percentage_chg = 0
  Columns("K:K").NumberFormat = "0.00%"

' Keep track of the location for ticker and where to find the last row
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  lastrow = Cells(Rows.Count, "A").End(xlUp).Row
 ' WS_Count = ActiveWorkbook.Worksheets.Count
  
  On Error Resume Next

' Loop through all of the active worksheets (???)

 
  ' Loop through all data
  For i = 2 To lastrow
    
    ' Check if we are still within the same ticker symbol, and if it's not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the first ticker symbol
      Ticker = Cells(i, 1).Value

      ' Add to the opening stock price
      price_open = price_open + Cells(i, 3).Value

      ' Add to the closing stock price
      price_close = price_close + Cells(i, 6).Value

      ' Add to the total volume
      total_volume = total_volume + Cells(i, 7).Value

      ' Figure the percentage change
      percentage_chg = ((vol_change_tot / price_open) * 100)
      
      ' Add to the difference in stock open and close
      vol_change_tot = (price_close - price_open)

      
      

      ' Print the ticker symbol in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the difference between open and close
      Range("J" & Summary_Table_Row).Value = vol_change_tot

      ' Print the difference between open and close
      Range("K" & Summary_Table_Row).Value = percentage_chg

      'Print the total volume
      Range("L" & Summary_Table_Row).Value = total_volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Totals
      price_open = 0
      price_close = 0
      total_volume = 0
      vol_change_tot = 0
      percentage_chg = 0
      
    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the opening price
     price_open = price_open + Cells(i, 3).Value

      ' Add to the closing price
     price_close = price_close + Cells(i, 6).Value

      'Add to the total volume
     total_volume = total_volume + Cells(i, 7).Value
     
      ' Add to the difference
     vol_change_tot = (price_close - price_open)

      ' Add tp the percentage change
     percentage_chg = ((vol_change_tot / price_open) * 100)
    
 
    End If
   
  Next i
  
  On Error GoTo 0

End Sub

Sub Pick()

MaxInc = Application.WorksheetFunction.Max(Range("K:K"))
MaxDec = Application.WorksheetFunction.Min(Range("K:K"))
MaxTot = Application.WorksheetFunction.Max(Range("L:L"))

lastrow = Cells(Rows.Count, "L").End(xlUp).Row
lastcol = Cells(1, Columns.Count).End(xlToLeft).Columns

For i = 2 To lastcol

    For j = 1 To lastrow

        If Cells(i, j).Value = MaxInc Then
        MaxInc = Cells(2, 15).Value

        End If
    Next j
Next i
    

End Sub

 


