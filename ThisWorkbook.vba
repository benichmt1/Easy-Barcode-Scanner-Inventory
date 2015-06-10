Private Sub Workbook_Open()

Do
    myId = InputBox("Please scan your student ID")
    If IsNumeric(myId) Then Exit Do
    MsgBox "No, a student ID please"
Loop

Do
    myInstrument = InputBox("Please scan your instrument's bar code")
    If IsNumeric(myInstrument) Then Exit Do
    MsgBox "No, the instrument bar code"
Loop

' Put these values in the spreadsheet
Range("J2").Value = myInstrument
Range("K2").Value = myId
End Sub
