Sub stock_volume_Easy():

Dim WS As Worksheet

For Each WS In ActiveWorkbook.Worksheets
  WS.Activate
    
      LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

      Cells(1, "I").Value = "Ticker"
      Cells(1, "J").Value = "Total Stock Volume"

  
  Dim Ticker_Name As String
  Dim Ticker_Volume_Total As Double
  Ticker_Volume_Total = 0

 
  Dim Summary_Table_Row As Double
  Summary_Table_Row = 2

  Dim i As Long

  For i = 2 To LastRow
    if Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
         Ticker_Name = Cells(i, 1).Value

         Ticker_Volume_Total = Ticker_Volume_Total + Cells(i, 7).Value
        
         Range("I" & Summary_Table_Row).Value = Ticker_Name
         Range("J" & Summary_Table_Row).Value = Ticker_Volume_Total

          Summary_Table_Row = Summary_Table_Row + 1

          Ticker_Volume_Total = 0


      Else

          Ticker_Volume_Total = Ticker_Volume_Total + Cells(i, 7).Value

      End If
  Next i
Next WS

End Sub