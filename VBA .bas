


Sub MultipleYearData()

  'Worksheet for loop enables data to populate for all sheets in one go (or run)

  For Each ws In Worksheets

      Dim ticker_symbol As String
      Dim open_price As Double
      Dim close_price As Double
      Dim total_volume As Double
      Dim Summary_Row_Table As Integer
      Dim yearly_change As Double
      Dim percent_change As Double
      Dim percent_change1 As Double

      total_volume = 0
      Summary_Row_Table = 2
      LastRow = ws.Range("A:A").SpecialCells(xlCellTypeLastCell).Row
      open_price = ws.Cells(2, 3).Value


   'I loop below condenses the original data into aggregated values for one ticker at a time

   For i = 2 To LastRow

     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     ticker_symbol = ws.Cells(i, 1).Value
     total_volume = total_volume + ws.Cells(i, 7).Value
     close_price = ws.Cells(i, 6).Value


     yearly_change = close_price - open_price
     percent_change = yearly_change / open_price
     ws.Range("I" & Summary_Row_Table).Value = ticker_symbol
     ws.Range("L" & Summary_Row_Table).Value = total_volume
     ws.Range("J" & Summary_Row_Table).Value = yearly_change
     open_price = ws.Cells(i + 1, 3).Value


     ws.Range("k" & Summary_Row_Table).Value = percent_change
     ws.Range("k" & Summary_Row_Table).NumberFormat = "0.00%"
     Summary_Row_Table = Summary_Row_Table + 1
     total_volume = 0


    Else
        total_volume = total_volume + ws.Cells(i, 7).Value


   End If

   Next i


'J loop sets conditional formatting for yearly change and percent change and stops at the last row of the summary table.


   For j = 2 To 3001

       If ws.Cells(j, 10).Value < 0 Then
       ws.Cells(j, 10).Interior.ColorIndex = 3
       End If

       If ws.Cells(j, 11).Value < 0 Then
       ws.Cells(j, 11).Interior.ColorIndex = 3
       End If

       If ws.Cells(j, 10).Value >= 0 Then
       ws.Cells(j, 10).Interior.ColorIndex = 4
       End if
       
       If ws.Cells(j, 11).Value >= 0 Then
       ws.Cells(j, 11).Interior.ColorIndex = 4
       End If

    End If
    
    Next j


    'Find min and max and apply percentage to produce greatest increase, decrease, and total volume values  (source:https://www.wallstreetmojo.com/vba-max/; this source was very helpful in identifying the WorksheetFunction.Max/Min function I needed to produce the values)

            ws.Cells(2, 16).Value = WorksheetFunction.Max(ws.Range("k2:k3002"))


            ws.Cells(3, 16).Value = WorksheetFunction.Min(ws.Range("k2:k3002"))


            ws.Cells(4, 16).Value = WorksheetFunction.Max(ws.Range("l2:l3002"))


            LastRow2 = ws.Range("K:K").SpecialCells(xlCellTypeLastCell).Row


    'Generate g loop for greatest % increase, decrease, and greatest total volume using logic to draw ticker of corresponding percent data

   
    For g = 2 To LastRow2
    
        If ws.Cells(2, 16).Value = ws.Cells(g, 11).Value Then
            ws.Cells(2, 15).Value = ws.Cells(g, 9).Value
        

        ElseIf ws.Cells(3, 16).Value = ws.Cells(g, 11).Value Then
         ws.Cells(3, 15).Value = ws.Cells(g, 9).Value
           
            

        ElseIf ws.Cells(4, 16).Value = ws.Cells(g, 12).Value Then
        ws.Cells(4, 15).Value = ws.Cells(g, 9).Value
        
        
        End If
  

    Next g
    

Next ws


End Sub







