
Private Sub UserForm_Initialize()
    ' Set the default values
    txtDate.Value = Format(Date + 1, "yyyy-mm-dd") ' Prefill with the current date
    txtDays.Value = 5 ' Change this to your default number
    txtMinutes.Value = 30 ' Default duration in minutes
End Sub

Private Sub btnOK_Click()
    ' Store values in public variables
    SelectedDate = txtDate.Value
    SelectedDays = CInt(txtDays.Value) ' Convert to Integer
    SelectedMinutes = CInt(txtMinutes.Value) ' Convert to Integer
    
    ' Close the form
    Unload Me
End Sub
