Sub Macro1()
'Check the number of rows in the spreadsheet
Dim last As Long, counter As Integer
last = Cells(Rows.Count, "A").End(xlUp).Row
For counter = 2 To last
    If Cells(counter, 1) = Cells(2, 10) And Not Cells(counter, 2) = "Yes" Then
        Cells(counter, 2) = "Yes"
        Cells(counter, 4) = Now()
        Cells(counter, 5) = Application.WorksheetFunction.VLookup(Range("K2").Value, StudentID.Range("A2:E93"), 3, False)
        'Updates the checked out field, writes a timestamp, and prints the student name
        Exit For
    ElseIf Cells(counter, 1) = Cells(2, 10) And Cells(counter, 2) = "Yes" Then
        MsgBox "This instrument is already checked out!"
        Exit For
    ElseIf counter = last Then
        MsgBox "Error: Instrument not found. Are you sure you scanned it correctly?"
        
    End If
Next counter
'
End Sub
