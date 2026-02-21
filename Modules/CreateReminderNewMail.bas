Option Explicit


' --- Create reminder from an open NEW mail window (no moving) ---
Public Sub CreateReminderNewMail()

    Dim m As Outlook.MailItem
    Dim subj As String
    Dim appt As Outlook.AppointmentItem

    CalendarFormSubmitted = False

    Set m = GetComposeMailItem()
    If m Is Nothing Then
        MsgBox "No open mail compose window found.", vbExclamation, "No Compose Window"
        Exit Sub
    End If

    ' Workaround: Outlook often updates .Subject only after you change focus/recipients
    subj = Trim$(CStr(m.Subject))
    If subj = "" Then
        MsgBox "Subject not detected yet." & vbCrLf & _
               "Quick workaround: click once into the To/CC field (or change recipients) and run again.", _
               vbExclamation, "Subject not committed"
        Exit Sub
    End If

    CreateReminderForm.txtReminderSubject.Value = subj
    CreateReminderForm.txtReminderDate.Value = Format$(Date + 2, "yyyy-mm-dd")
    CreateReminderForm.Show

    If Not CalendarFormSubmitted Then Exit Sub

    Set appt = Outlook.Application.CreateItem(olAppointmentItem)
    With appt
        .Subject = "[REMINDER] " & CalendarSubject
        .Start = CalendarDate
        .AllDayEvent = True
        .BusyStatus = olFree
        .Categories = "REMINDER"
        .Save
    End With

    MsgBox "Reminder created.", vbInformation, "Success"
End Sub

' --- Get MailItem from active/newest Inspector, or inline response ---
Private Function GetComposeMailItem() As Outlook.MailItem
    On Error GoTo CleanFail

    Dim olApp As Outlook.Application
    Dim insp As Outlook.Inspector
    Dim itm As Object
    Dim i As Long

    Set olApp = Outlook.Application

    Set insp = olApp.ActiveInspector
    If Not insp Is Nothing Then
        Set itm = insp.CurrentItem
        If TypeOf itm Is Outlook.MailItem Then Set GetComposeMailItem = itm: Exit Function
    End If

    For i = olApp.Inspectors.Count To 1 Step -1
        Set insp = olApp.Inspectors.Item(i)
        Set itm = insp.CurrentItem
        If TypeOf itm Is Outlook.MailItem Then Set GetComposeMailItem = itm: Exit Function
    Next i

    If Not olApp.ActiveExplorer Is Nothing Then
        If olApp.ActiveExplorer.ActiveInlineResponse Then
            Set itm = olApp.ActiveExplorer.ActiveInlineResponse
            If TypeOf itm Is Outlook.MailItem Then Set GetComposeMailItem = itm: Exit Function
        End If
    End If

CleanFail:
    Set GetComposeMailItem = Nothing
End Function

