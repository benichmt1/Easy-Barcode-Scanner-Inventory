Private Sub Workbook_Open()
Do While True
    Dim myDate As String
    
    Do
        myId = InputBox("Please scan your student ID")
        Range("K2").Value = myId
        If IsNumeric(myId) Then Exit Do
        MsgBox "No, a student ID please"
        
    Loop
    
    Do
        myInstrument = InputBox("Please scan your instrument's bar code")
        Range("J2").Value = myInstrument
        If IsNumeric(myInstrument) Then Exit Do
        MsgBox "No, the instrument bar code"
        
    Loop
    
    ' K2 has a VLOOKUP query built in. Add this later for modularity
    
    
    Dim last As Long, counter As Integer
    Dim a As Variant
    a = Cells(4, 11).Value
    last = Cells(Rows.Count, "A").End(xlUp).Row
    For counter = 2 To last
        If Cells(counter, 1) = Cells(2, 10) And Not Cells(counter, 2) = "Yes" Then
            Cells(counter, 2) = "Yes"
            Cells(counter, 4) = Now()
            Cells(counter, 5).Value = a
            Exit For
        ElseIf Cells(counter, 1) = Cells(2, 10) And Cells(counter, 2) = "Yes" Then
            MsgBox "This instrument is already checked out!"
            Exit For
        ElseIf counter = last Then
            MsgBox "Error: Instrument not found. Are you sure you scanned it correctly?"
            
        End If
    Next counter
Loop
End Sub
