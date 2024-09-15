Sub OpenTicketForm()
    TicketForm.Show
End Sub

'Code inside the Ticket Form module
Private Sub SubmitButton_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Ticket Data")
    
    'Find the next available row
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    'Store form data in the next available row
    ws.Cells(nextRow, 1).Value = TicketIDTextBox.Value
    ws.Cells(nextRow, 2).Value = DateCreatedTextBox.Value
    ws.Cells(nextRow, 3).Value = AssignedToTextBox.Value
    ws.Cells(nextRow, 4).Value = StatusTextBox.Value
    ws.Cells(nextRow, 5).Value = PriorityTextBox.Value
    ws.Cells(nextRow, 6).Value = DueDateTextBox.Value
    
    'Clear the form
    TicketIDTextBox.Value = ""
    DateCreatedTextBox.Value = ""
    AssignedToTextBox.Value = ""
    StatusTextBox.Value = ""
    PriorityTextBox.Value = ""
    DueDateTextBox.Value = ""
    
    MsgBox "Ticket added successfully!"
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub
