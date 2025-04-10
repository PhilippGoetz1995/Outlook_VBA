' Declare public variables to hold the form values
Public selectedDate As String
Public selectedDays As Integer
Public selectedMinutes As Integer


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

    ' GENERAL NOTES
    ' # Get Start and End Date from form
    ' # Format the dates and times correct
    ' # Go trough every working day and start with the beginning of the day and check for busy meetings
    ' # add it to the current mail

    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olCalendar As Outlook.MAPIFolder
    Dim olItems As Outlook.Items
    Dim filteredItems As Outlook.Items
    Dim olItem As Object
    
    Dim startDate As Date, endDate As Date
    Dim startTime As Date, endTime As Date
    Dim filter As String
    
    Dim freeSlots As String
    
    Dim FreeSlotStart As Date
    Dim FreeSlotEnd As Date
    
    FindSlotForm.Show
    
    ' Validate input
    If Not IsNumeric(selectedDays) Or selectedDays < 1 Then
        MsgBox "Bitte geben Sie eine gültige Zahl größer als 0 ein.", vbExclamation, "Ungültige Eingabe"
        Exit Sub
    End If
    
    ' Set date range
    startDate = CDate(selectedDate)
    endDate = startDate + CInt(selectedDays) ' User-defined number of days ahead
    
    startDate = startDate + TimeValue("07:00:00")
    endDate = endDate + TimeValue("18:00:00")


    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olCalendar = olNS.GetDefaultFolder(olFolderCalendar)
    Set olItems = olCalendar.Items

    ' Important: Sort by start date and include recurring items
    olItems.Sort "[Start]"
    olItems.IncludeRecurrences = True

    ' Define filter in the following syntax dd.mm.yyyy HH:nn
    filter = "[Start] >= '" & Format(startDate, "dd.mm.yyyy HH:nn") & "' AND [End] <= '" & Format(endDate, "dd.mm.yyyy HH:nn") & "'"

    Set filteredItems = olItems.Restrict(filter)

    ' DEBUGGING Loop through filtered items
'    For Each olItem In filteredItems
'        Debug.Print "Subject: " & olItem.Subject
'        Debug.Print "Start: " & olItem.Start
'        Debug.Print "End: " & olItem.End
'        Debug.Print "----------------------------"
'    Next
'
'    Debug.Print filteredItems.Count & " calendar items found between " & startDate & " and " & endDate
    
    freeSlots = ""
    
    For dt = startDate To endDate
        If Weekday(dt, vbMonday) <= 5 Then ' Only Monday to Friday

            ' Set every iteration the start of the day
            startTime = DateValue(dt) + TimeSerial(7, 0, 0)
            endTime = DateValue(dt) + TimeSerial(18, 0, 0)

            ' Set the Free Slot Start Time to 7 AM
            FreeSlotStart = startTime
            
            ' Iterate trough each calender item
            For Each olItem In filteredItems
            
                If olItem.BusyStatus = olBusy Or olItem.BusyStatus = olOutOfOffice Then
                
                    ' Event need to be within working hours
                    If olItem.Start >= startTime And olItem.Start < endTime Then
                        
                        ' If the event is parallel to the previous event => Only if the End of the Event is behind the start of the Free Slot => Otherwise just go to the next one
                        If olItem.End >= FreeSlotStart Then
                            
                            ' New End of Free Slot is the Start of the Next Event
                            FreeSlotEnd = olItem.Start
                    
                            If DateDiff("n", FreeSlotStart, FreeSlotEnd) >= selectedMinutes Then
                                freeSlots = freeSlots & GetEnglishWeekday(Weekday(FreeSlotStart)) & " " & Format(FreeSlotStart, "dd.mm. h:mm AM/PM") & "  - " & Format(FreeSlotEnd, "h:mm AM/PM") & vbCrLf
                            End If
                            
                            FreeSlotStart = olItem.End
                        End If
                    End If
                End If
            Next
            
            ' Check for free time between last meeting and end of work hours
            If DateDiff("n", FreeSlotStart, endTime) >= 30 Then
                freeSlots = freeSlots & GetEnglishWeekday(Weekday(FreeSlotStart)) & " " & Format(FreeSlotStart, "dd.mm. h:mm AM/PM") & "  - " & Format(endTime, "h:mm AM/PM") & vbCrLf
            End If
        End If
    Next dt
    
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

End Sub
