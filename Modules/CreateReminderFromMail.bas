Public CalendarSubject As String
Public CalendarDate As Date

Sub CreateReminderEntryFromMail()

    Dim OutlookApp As Object
    Dim Explorer As Object
    Dim Selection As Object
    Dim MailItem As Object
    Dim Folder As Object
    
    ' Get the active explorer and the selected email
    Set OutlookApp = Application
    Set Explorer = OutlookApp.ActiveExplorer
    Set Selection = Explorer.Selection
    Set NameSpace = OutlookApp.Session

    ' Make sure there is at least one email selected
    If Selection.Count > 0 Then
        Set MailItem = Selection.Item(1) ' Get the first selected email

        ' Show the custom form
        CreateReminderForm.txtReminderSubject.Value = MailItem.Subject
        CreateReminderForm.txtReminderDate.Value = Format(Date + 2, "yyyy-mm-dd") ' Default to today's date
        CreateReminderForm.Show
    
        editedSubject = "[REMINDER] " & CalendarSubject

        ' Move the email to a dedicated folder
        Set Folder = NameSpace.GetDefaultFolder(6).Folders("02_FOLLOWUP")
        MailItem.Move Folder
        
        ' Create a new calendar entry
        Set objAppointment = Outlook.Application.CreateItem(olAppointmentItem)
        With objAppointment
            .Subject = editedSubject
            .Start = CalendarDate
            .AllDayEvent = True
            .BusyStatus = olFree
            .Categories = "REMINDER"
            .Save
        End With
        

    Else
        MsgBox "No email selected!"
    End If



End Sub
