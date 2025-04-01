Private Sub UserForm_Initialize()
    CalendarFormSubmitted = False ' Reset flag when form opens
End Sub

Private Sub btnOK_Click()
    ' Store values in public variables
    CalendarSubject = txtReminderSubject.Value
    CalendarDate = txtReminderDate.Value
    
    CalendarFormSubmitted = True
    
    ' Close the form
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    ' Only reset CalendarFormSubmitted if user clicks the close button (X)
    If CloseMode = vbFormControlMenu Then
        CalendarFormSubmitted = False
    End If

End Sub
