
' Helper function for EN weekdays
Function GetEnglishWeekday(dayNumber As Integer) As String
    Select Case dayNumber
        Case 1: GetEnglishWeekday = "Su.,"
        Case 2: GetEnglishWeekday = "Mo.,"
        Case 3: GetEnglishWeekday = "Tu.,"
        Case 4: GetEnglishWeekday = "We.,"
        Case 5: GetEnglishWeekday = "Th.,"
        Case 6: GetEnglishWeekday = "Fr.,"
        Case 7: GetEnglishWeekday = "Sa.,"
    End Select
End Function


Sub FindFreeSlotsAndInsertToNewMessage()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItems As Outlook.Items
    Dim olAppointment As Outlook.AppointmentItem
    Dim olRestrict As Outlook.Items
    Dim StartDate As Date, EndDate As Date
    Dim startTime As Date, EndTime As Date
    Dim StartDateString As String
    Dim EndDateString As String
    Dim dt As Date
    Dim freeSlots As String
    Dim BusySlots As Object
    Dim FreeSlotStart As Date
    Dim FreeSlotEnd As Date
    Dim objSelection As Object
    Dim i As Integer
    
    ' Initialize Outlook objects
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olFolder = olNS.GetDefaultFolder(olFolderCalendar)
    Set olItems = olFolder.Items
    
    ' Ask user for the number of days
    DaysAhead = InputBox("Geben Sie die Anzahl der zu überprüfenden Tage ein (nur Werktage):", "Zeitraum auswählen", "7")
    
    ' Validate input
    If Not IsNumeric(DaysAhead) Or DaysAhead < 1 Then
        MsgBox "Bitte geben Sie eine gültige Zahl größer als 0 ein.", vbExclamation, "Ungültige Eingabe"
        Exit Sub
    End If
    
    ' Set date range
    StartDate = DateAdd("d", 1, Date)
    EndDate = StartDate + CInt(DaysAhead) ' User-defined number of days ahead
    
    StartDateString = Format(StartDate, "yyyy-mm-dd hh:mm AMPM") & "AM"
    EndDateString = Format(EndDate, "yyyy-mm-dd hh:mm AMPM") & "AM"
    
    ' Filter calendar items for the given period
    olItems.Sort "[Start]", False
    olItems.IncludeRecurrences = True
    
    Dim filter As String
    filter = "[Start] >= '" & StartDateString & "' AND [Start] <= '" & EndDateString & "'"
    
    Set olRestrict = olItems.Restrict(filter)

    ' Loop through each day in the range
    freeSlots = ""
    
    EndDate = DateAdd("d", -1, EndDate)
    
    For dt = StartDate To EndDate
        If Weekday(dt, vbMonday) <= 5 Then ' Only Monday to Friday
            startTime = dt + TimeValue("07:00:00") ' 07:00 AM
            EndTime = dt + TimeValue("19:00:00")   ' 07:00 PM
            
            ' Set the Free Slot Start Time to 7 AM
            FreeSlotStart = startTime
            
            ' Get all appointments for the day
            Set BusySlots = olRestrict
            For Each olAppointment In BusySlots
            
                If olAppointment.BusyStatus = olBusy Or olAppointment.BusyStatus = olOutOfOffice Then
                
                    ' Event need to be within working hours
                    If olAppointment.Start >= startTime And olAppointment.Start < EndTime Then
                        
                        ' If the event is parallel to the previous event => Only if the End of the Event is behind the start of the Free Slot => Otherwise just go to the next one
                        If olAppointment.End >= FreeSlotStart Then
                            
                            ' New End of Free Slot is the Start of the Next Event
                            FreeSlotEnd = olAppointment.Start
                    
                            If DateDiff("n", FreeSlotStart, FreeSlotEnd) >= 30 Then
                                freeSlots = freeSlots & GetEnglishWeekday(Weekday(FreeSlotStart)) & " " & Format(FreeSlotStart, "dd.mm. h:mm AM/PM") & "  - " & Format(FreeSlotEnd, "h:mm AM/PM") & vbCrLf
                            End If
                            
                            FreeSlotStart = olAppointment.End
                        End If
                    End If
                End If
            Next
            
            ' Check for free time between last meeting and end of work hours
            If DateDiff("n", FreeSlotStart, EndTime) >= 30 Then
                freeSlots = freeSlots & GetEnglishWeekday(Weekday(FreeSlotStart)) & " " & Format(FreeSlotStart, "dd.mm. h:mm AM/PM") & "  - " & Format(EndTime, "h:mm AM/PM") & vbCrLf
            End If
        End If
    Next dt
    
    ' Debug.Print freeSlots
    
    ' Get the currently active Outlook application
    Set olApp = Outlook.Application
    
    ' Check if there is an active inspector (open email window)
    If olApp.ActiveInspector Is Nothing Then
        MsgBox "No open email found. Please open an email and place the cursor where you want to insert the text.", vbExclamation, "No Active Email"
        Exit Sub
    End If
    
    ' Get the active mail item
    Set objInspector = olApp.ActiveInspector
    Set objMail = objInspector.CurrentItem
    
    ' Ensure it's an email
    If objMail.Class <> olMail Then
        MsgBox "The active window is not an email.", vbExclamation, "Invalid Window"
        Exit Sub
    End If
    

    ' Get the Word editor of the email body
    Set objWord = objInspector.WordEditor
    Set objSelection = objWord.Application.Selection
    
    ' Insert free slots at the cursor position
    objSelection.TypeText freeSlots
    
    ' Clean up
    Set olApp = Nothing
    Set olNS = Nothing
    Set olFolder = Nothing
    Set olItems = Nothing
    Set objSelection = Nothing
End Sub

