Public CalendarSubject As String
Public CalendarDate As Date
Public CalendarFormSubmitted As Boolean

Sub CreateReminderMail()
    
    Dim MailSubjectSelectedMail As String
    
    'Get the Subject of the selected Mail
    MailSubjectSelectedMail = GetSelectedMailSubject
    
    ' Make sure there is at least one email selected
    If MailSubjectSelectedMail <> "0" Then

        ' Show the custom form
        CreateReminderForm.txtReminderSubject.Value = MailSubjectSelectedMail
        CreateReminderForm.txtReminderDate.Value = Format(Date + 2, "yyyy-mm-dd") ' Default to today's date
        CreateReminderForm.Show

        ' If the user closed the form without submitting, exit the function
        If CalendarFormSubmitted Then

            editedSubject = "[REMINDER] " & CalendarSubject

            ' Move the selected Mail OR Conversation to the correct Folder with HELPER FUNCTION
            MoveMailOrConversation "02_FOLLOWUP"

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
            ' If the User clicked exit button
            Exit Sub

        End If
        
        MsgBox "Mail moved and Reminder created", vbInformation, "Success"

    Else
        MsgBox "No email selected!"
    End If

End Sub

