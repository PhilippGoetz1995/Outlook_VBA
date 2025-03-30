Option Explicit

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Function GetEmailText() As String
    Dim objMail As Outlook.MailItem
    Dim emailBody As String
    Dim signatureStart As Long
    Dim replyPosition As Long
    
    ' Check if a mail item is selected
    If TypeName(Application.ActiveWindow) = "Inspector" Then
        Set objMail = Application.ActiveInspector.CurrentItem
        ' Get the HTML body of the mail
        emailBody = objMail.body
        
    Else
        MsgBox "No email is currently open."
    End If
    
    ' Look for "Philipp Goetz" as the start of the signature
    signatureStart = InStr(emailBody, "Philipp Goetz")
    
    ' Or look for div divider if it is a reply Mail
    replyPosition = InStr(emailBody, "From:")
    
    ' If the signature is found, remove everything after it
    If replyPosition > 0 Then
        emailBody = Left(emailBody, replyPosition - 1)
    ' If the signature is found, remove everything after it
    ElseIf signatureStart > 0 Then
        emailBody = Left(emailBody, signatureStart - 1)
    End If
    
    ' Remove only trailing carriage returns (\r at the end)
    Do While Right(emailBody, 1) = vbCr Or Right(emailBody, 1) = vbLf
        emailBody = Left(emailBody, Len(emailBody) - 1)
    Loop
    
    ' Trim spaces and return cleaned text
    GetEmailText = Trim(emailBody)
     
End Function


Function SendToExternalService(originalText As String, URL As String)
    Dim xhr As Object
    Dim jsonData As String
    Dim loadingTime As Double
    Dim loadingText As String
    Dim Response As String
    Dim JSONEncoded As Object
    Dim JSONBody As Object
    Dim translatedText As String

    ' Create XMLHttpRequest object
    Set xhr = CreateObject("MSXML2.XMLHTTP")
    
    Set JSONBody = JsonConverter.ParseJson("{}") ' Create an empty JSON object
    JSONBody("text") = originalText

    ' Open the request using the POST method (async = True)
    xhr.Open "POST", URL, True

    ' Set the content type header for JSON data
    xhr.setRequestHeader "Content-Type", "application/json"

    ' Send the request with JSON data
    xhr.Send JsonConverter.ConvertToJson(JSONBody)
    
    ' Set up the initial loading message
    loadingText = "Translation is Loading... "
    loadingTime = 0
    
    ' Wait for the request to complete
    Do While xhr.readyState <> 4
        DoEvents ' Allow the app to process other events

        loadingTime = loadingTime + 0.1
        loadingText = "Translation is Loading... " & Format$(loadingTime, "0.0") & "s"

        ModifyEmailBody (loadingText)

        ' Wait for 0.1sec
        Sleep 100

        If loadingTime >= 10 Then
            ModifyEmailBody ("Translation request timed out.")
            Exit Do ' Exit loop after timeout
        End If
    Loop

    ' Get response
    If xhr.Status = 200 Then
        Response = xhr.responseText

        Set JSONEncoded = JsonConverter.ParseJson(Response)
        translatedText = JSONEncoded("output")

        ModifyEmailBody (translatedText)

    Else
        translatedText = "Error in translation => HTTP Error: " & xhr.Status

        ModifyEmailBody (translatedText)

    End If
    
End Function

Function ModifyEmailBody(textInsert As String)
    Dim objMail As Outlook.MailItem
    Dim htmlBody As String
    Dim fontFamily As String
    Dim fontSize As String
    Dim mailTextToReplace As String
    Dim insertPosition As Long
    
    
    ' Check if the active window is a MailItem
    If TypeName(Application.ActiveWindow) = "Inspector" Then
        Set objMail = Application.ActiveWindow.CurrentItem
    Else
        MsgBox "No email is open."
        Exit Function
    End If
    
    ' Get the HTML body of the email
    htmlBody = objMail.htmlBody
    
    fontFamily = "Calibri Light" ' Default font
    fontSize = "11pt" ' Default font size
    
    If InStr(htmlBody, "---- END OF INSERTED TEXT ----") > 0 Then
    
        ' Translation is Loading... 0,1s
        insertPosition = InStr(htmlBody, "Translation is Loading...")
        
        mailTextToReplace = Mid(htmlBody, insertPosition, 30)
        
        htmlBody = Replace(htmlBody, mailTextToReplace, textInsert)
        
    Else
        
        'Insert the Divider into the Mail
    
        textInsert = "<span style=""font-family:" & fontFamily & "; font-size:" & fontSize & """>" & textInsert & "</span> <br><br>---- END OF INSERTED TEXT ----<br><br>"
        
        ' It should be inserted into the Body of the Mail
        insertPosition = InStr(htmlBody, "<div class=WordSection1>")
        
        ' If found, insert the variable after it
        If insertPosition > 0 Then
        
            ' Insert the variable after the specific text
            htmlBody = Left(htmlBody, insertPosition + Len("<div class=WordSection1>") - 1) & textInsert & Mid(htmlBody, insertPosition + Len("<div class=WordSection1>"))
                       
        End If
          
    End If
    
    ' Update the email body
    objMail.htmlBody = htmlBody
    
End Function


' Translate Text With ChatGPT
Sub TranslateTextWithChatGPT()
    Dim originalText As String
    Dim URL As String
    
    ' Get selected text
    originalText = GetEmailText()
    
    If originalText = "" Then
       MsgBox "No text selected!", vbExclamation, "Translation"
       Exit Sub
    End If
    
    URL = "PLACEHOLDER N8N URL for Translation webhook"
    
    ' Translate text via n8n
    SendToExternalService originalText, URL
    
End Sub


' Translate Text With ChatGPT
Sub CheckTextWithChatGPT()
    Dim originalText As String
    Dim URL As String
    
    ' Get selected text
    originalText = GetEmailText()
    
    If originalText = "" Then
       MsgBox "No text selected!", vbExclamation, "Translation"
       Exit Sub
    End If
    
    URL = "PLACEHOLDER N8N URL for CheckText webhook"
    
    ' Translate text via n8n
    SendToExternalService originalText, URL
    
End Sub

