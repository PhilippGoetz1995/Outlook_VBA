
'Speciallity from VBA is that the module Name can not be the same as the sub Name
Sub DebugMacro()

    MsgBox "Test Message"
    
    Debug.Print "test"
    
    Dim testSubject As String
    
    testSubject = GetSelectedMailSubject
    
    Debug.Print testSubject
    
End Sub
