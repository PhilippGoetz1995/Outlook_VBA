Sub AddMailToTrello()

    Dim URL As String
    Dim HttpRequest As Object
    Dim MailSubjectSelectedMail As String
    
    'Get the Subject of the selected Mail
    MailSubjectSelectedMail = GetSelectedMailSubject
    
    ' Make sure there is at least one email selected
    If MailSubjectSelectedMail <> "0" Then
        
        ' Define the external URL to send the subject to
        URL = "https://p-goetz.app.n8n.cloud/webhook/a68b3a99-5b82-4330-b520-d5adc6a45ba4"
        
    
        ' Show an input box for the user to edit the subject
        MailSubjectSelectedMail = InputBox("Title of Card:", "Send", MailSubjectSelectedMail)
        
        ' Check if the user pressed Cancel, ESC, or closed the box
        If MailSubjectSelectedMail = "" Then
            Exit Sub ' Stop further execution
        End If

        ' Create an HTTP request to send the subject to the external URL
        Set HttpRequest = CreateObject("MSXML2.XMLHTTP")
        HttpRequest.Open "POST", URL, False
        HttpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        HttpRequest.Send "subject=" & MailSubjectSelectedMail

        ' Check for success (you can customize this part as needed)
        If HttpRequest.Status = 200 Then
            ' MsgBox "Subject sent successfully!"
        Else
            MsgBox "Error sending subject: " & HttpRequest.Status & " - " & HttpRequest.statusText
        End If
        
        ' Move the selected Mail OR Conversation to the correct Folder with HELPER FUNCTION
        MoveMailOrConversation "01_TODO"
        
        MsgBox "Mail moved and Trello Task created", vbInformation, "Success"

    Else
        MsgBox "No email selected!"
    End If
End Sub
