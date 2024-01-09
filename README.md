'starting from the start
Sub multi_year()

    Dim ws As Worksheet
    For Each ws In Worksheets
    
    Dim lr As Long
    lr = Cells(Rows.Count, 1).End(xlUp).Row
    
'HEADING'S - THIS WORKS
    ws.Cells(1, 9).Value = "Ticker Code"
    ws.Cells(1, 12).Value = "Stock Volume"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage"

'IDENTIFYING CODES
    Dim Ticker As String
'GET DATA FOR VOLUME
    Dim Volume As Double
    Volume = 0
'GET DATA FOR OPENING
    Dim Opening As Double
    Opening = 0
'GET DATA FOR CLOSING
    Dim Closing As Double
    Closing = 0
'GET DATA FOR PERCENTAGE
    Dim Percentage As Double
    Percentage = 0
    
'SUMMARY - THIS WORKS
    Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
'LOOP - to works
  For i = 2 To lr
  
      
'GROUP TICKER CODES - THIS WORKS
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker = ws.Cells(i, 1).Value
    Volume = Volume + ws.Cells(i, 7).Value
    Closing = ws.Cells(i, 6).Value
    
'YEARLY CHANGE CACULATION
    yearly_change = Closing - Opening
    
'PERCENTAGE CACULATION
    Percentage = yearly_change / Opening

'SUMMARY
    ws.Range("I" & Summary_Table_Row).Value = Ticker
    ws.Range("L" & Summary_Table_Row).Value = Volume
    ws.Range("j" & Summary_Table_Row).Value = yearly_change
    ws.Range("k" & Summary_Table_Row).Value = Percentage
    Summary_Table_Row = Summary_Table_Row + 1
    Volume = 0
    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    Opening = ws.Cells(i, 3).Value
    Else
      Volume = Volume + ws.Cells(i, 7).Value
       
    End If

'CHANGE BACKGROUND TO GREEN OR RED yearly change
    If ws.Cells(i, 10).Value >= 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3
     
     ElseIf ws.Cells(i, 10).Value <= 0 Then
     ws.Cells(i, 10).Interior.ColorIndex = 4
     Else
     ws.Cells(i, 10).Interior.ColorIndex = 1
    
     End If
     
'CHANGE BACKGROUND TO GREEN OR RED percentage
   If ws.Cells(i, 11).Value >= 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 3
    
    ElseIf ws.Cells(i, 11).Value <= 0 Then
     ws.Cells(i, 11).Interior.ColorIndex = 4
     Else
     ws.Cells(i, 11).Interior.ColorIndex = 1
        
     End If
     
'Percentage format
    ws.Cells(i, 11).NumberFormat = "0.00%"
    
'LOOP THROUGH WORSK SHEETS
   Next i

   Next ws
    
End Sub


