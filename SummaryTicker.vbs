Attribute VB_Name = "Module1"

Sub Summary_TickerName()

Dim ws As Worksheet


For Each ws In Worksheets

    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    ' Add the word Year to the First Column Header

    ws.Range("M" & 1).Value = "Ticker"
     ' Keep track of the location for each ticker name in the summary table

     Dim Summary_Table_Row As Integer
     
     Summary_Table_Row = 2


    ' Loop through all credit card purchases

    For i = 2 To LastRow


    ' Check if we are still within the same ticker name, if it is not...
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    

      ' Set the ticker name
      Ticker_Name = ws.Cells(i, 1).Value

    
      ' Print the Ticker name in the Summary Table
      ws.Range("M" & Summary_Table_Row).Value = Ticker_Name

     
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      
    ' If the cell immediately following a row is the same brand...
    End If
    
    

Next i
Next ws
End Sub



