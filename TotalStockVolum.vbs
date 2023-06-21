Attribute VB_Name = "Module1"
Sub TotalStockVolum()

Dim ws As Worksheet


For Each ws In Worksheets

    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    ' Add the total stock volum to the First Column Header

    ws.Range("P" & 1).Value = "Total Stock Volum"
     ' Keep track of the location for each ticker name in the summary table

     Dim Summary_Table_Row As Integer
     
     Summary_Table_Row = 2
     
    Dim Total_Volum As Double
     
Total_Volum = 0

    ' Loop through all credit card purchases

    For i = 2 To LastRow


    ' Check if we are still within the same ticker name, if it is not...
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    

      ' Set the Total_Volum
      
      Total_Volum = Total_Volum + ws.Cells(i, 7).Value

    
      ' Print the Ticker name in the Summary Table
      ws.Range("P" & Summary_Table_Row).Value = Total_Volum

     
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
    'Reset Total_Volum
     Total_Volum = 0
    
    ' If the cell immediately following a row is the same brand...
    
    Else
    
     Total_Volum = Total_Volum + ws.Cells(i, 7).Value
    
    End If
    
    

Next i
Next ws
End Sub


