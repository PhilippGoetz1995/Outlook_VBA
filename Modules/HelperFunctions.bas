' Helper function to get the Subject of the currently selected Mail
Function GetSelectedMailSubject() As String
    Dim objOutlook As Outlook.Application
    Dim objSelection As Outlook.Selection
    Dim objMail As Outlook.MailItem
    
    ' Initialize Outlook application
    Set objOutlook = Application
    ' Get the selected items in the active explorer
    Set objSelection = objOutlook.ActiveExplorer.Selection
    
    ' Check if a mail item is selected
    If objSelection.Count > 0 Then
        ' Ensure the selected item is a mail item
        If TypeOf objSelection.Item(1) Is Outlook.MailItem Then
            Set objMail = objSelection.Item(1)
            GetSelectedMailSubject = objMail.Subject
        Else
            GetSelectedMailSubject = "0"
        End If
    Else
        GetSelectedMailSubject = "0"
    End If
    
End Function


' Helper function to move the selected Mail or Conversation
Sub MoveMailOrConversation(targetFolderName As String)
    Dim objApp As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objExplorer As Outlook.Explorer
    Dim objSelection As Outlook.Selection
    Dim objMail As Outlook.MailItem
    Dim objConversation As Outlook.Conversation
    Dim objFolder As Outlook.Folder
    Dim objItems As Outlook.SimpleItems
    Dim objItem As Object
    Dim targetFolder As Outlook.Folder
    
    ' Set Outlook application and namespace
    Set objApp = Outlook.Application
    Set objNamespace = objApp.GetNamespace("MAPI")
    Set objExplorer = objApp.ActiveExplorer
    Set objSelection = objExplorer.Selection

    ' Check if a mail item is selected
    If objSelection.Count = 0 Then
        MsgBox "Please select an email.", vbExclamation, "No Email Selected"
        Exit Sub
    End If

    ' Set the first selected email
    Set objMail = objSelection.Item(1)

    ' Set the target folder (Change this to your desired folder path)
    Set targetFolder = objNamespace.GetDefaultFolder(6).Folders(targetFolderName)

    ' Check if email is part of a conversation
    Set objConversation = objMail.GetConversation
    If Not objConversation Is Nothing Then
        ' Get all items in the conversation from all folders
        Set objTable = objConversation.GetTable
        
        Do Until objTable.EndOfTable
            Set objRow = objTable.GetNextRow
            Set objItem = objNamespace.GetItemFromID(objRow("EntryID"))
            
            ' Ensure it's a MailItem before moving
            If Not objItem Is Nothing Then
                If TypeOf objItem Is MailItem Then
                    objItem.Move targetFolder
                End If
            End If
        Loop
    Else
        ' Move only the selected mail
        objMail.Move targetFolder
    End If

End Sub
