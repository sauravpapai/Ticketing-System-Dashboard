Sub CheckOverdueTickets()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Ticket Data")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 6).Value < Date And ws.Cells(i, 4).Value <> "Closed" Then
            ws.Cells(i, 4).Value = "Overdue"
            ws.Cells(i, 4).Font.Color = vbRed
        End If
    Next i
    MsgBox "Overdue tickets have been updated!"
End Sub
