Sub AddMailToTrello()
    Dim OutlookApp As Object
    Dim Explorer As Object
    Dim Selection As Object
    Dim MailItem As Object
    Dim Subject As String
    Dim URL As String
    Dim HttpRequest As Object
    Dim Folder As Object

    ' Get the active explorer and the selected email
    Set OutlookApp = Application
    Set Explorer = OutlookApp.ActiveExplorer
    Set Selection = Explorer.Selection
    Set NameSpace = OutlookApp.Session

    ' Make sure there is at least one email selected
    If Selection.Count > 0 Then
        Set MailItem = Selection.Item(1) ' Get the first selected email
        Subject = MailItem.Subject ' Get the subject of the email
        
        ' Define the external URL to send the subject to
        URL = URL = "PLACEHOLDER N8N URL for sendToTrello webhook"
        
        ' STAGE URL
        ' URL = "https://p-goetz.app.n8n.cloud/webhook-test/b000f3e6-5b73-4dd6-9337-944d566c9930"
    
        ' Show an input box for the user to edit the subject
        Subject = InputBox("Title of Card:", "Send", Subject)

        ' Create an HTTP request to send the subject to the external URL
        Set HttpRequest = CreateObject("MSXML2.XMLHTTP")
        HttpRequest.Open "POST", URL, False
        HttpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        HttpRequest.Send "subject=" & Subject

        ' Check for success (you can customize this part as needed)
        If HttpRequest.Status = 200 Then
            ' MsgBox "Subject sent successfully!"
        Else
            MsgBox "Error sending subject: " & HttpRequest.Status & " - " & HttpRequest.statusText
        End If

        ' Move the email to a dedicated folder
        Set Folder = NameSpace.GetDefaultFolder(6).Folders("01_TODO")
        MailItem.Move Folder

    Else
        MsgBox "No email selected!"
    End If
End Sub
