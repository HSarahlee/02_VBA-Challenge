Sub Ticker()

For Each ws In Worksheets

Dim Ticker As String
Dim Total_vol As Double
Dim Summary_table_row As Integer
Dim openprice As Double
Dim closeprice As Double
Dim y_change As Double
Dim p_change As Double
Dim maxticker As String
Dim maxpercent As Double
Dim minticker As String
Dim minpercent As Double
Dim maxvolticker As String
Dim maxvol As Long

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stocke Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Total_vol = 0
Summary_table_row = 2
openprice = ws.Cells(2, 3).Value
last_row_table = ws.Cells(Rows.Count, 9).End(xlUp).Row


For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker = ws.Cells(i, 1).Value
    ws.Cells(Summary_table_row, 9).Value = Ticker
    
    Total_vol = Total_vol + ws.Cells(i, 7).Value
    ws.Cells(Summary_table_row, 12).Value = Total_vol
    
    closeprice = ws.Cells(i, 6).Value
    y_change = closeprice - openprice
    ws.Cells(Summary_table_row, 10).Value = y_change
    
     If openprice = 0 Then
            p_change = 0

        Else
            p_change = y_change / openprice
        
        
        End If
        
        ws.Cells(Summary_table_row, 11).Value = p_change
        ws.Cells(Summary_table_row, 11).NumberFormat = "0.00%"
        
        If ws.Cells(Summary_table_row, 10).Value > 0 Then
            ws.Cells(Summary_table_row, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(Summary_table_row, 10).Interior.ColorIndex = 3
        End If
        
        
        Summary_table_row = Summary_table_row + 1
        openprice = ws.Cells(i + 1, 3).Value
        Total_vol = 0

Else

    Total_vol = Total_vol + ws.Cells(i, 7).Value

End If
Next i

For i = 2 To last_row_table
    If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) Then
        ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    
    ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) Then
        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow)) Then
        ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value

    End If


Next i
Next ws

End Sub
    




