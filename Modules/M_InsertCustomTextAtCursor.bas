Sub InsertFormattedTextAtCursor()
    Dim insp As Outlook.Inspector
    Dim wdDoc As Object ' Word.Document
    Dim sel As Object ' Word.Selection
    Dim todayStr As String
    Dim fullText As String
    
    todayStr = Format(Date, "dd.mm.")
    fullText = "[PG " & todayStr & "]"
    
    Set insp = Application.ActiveInspector
    If insp Is Nothing Then Exit Sub
    If Not insp.CurrentItem.Class = olMail Then Exit Sub
    If insp.EditorType <> olEditorWord Then Exit Sub
    
    Set wdDoc = insp.WordEditor
    Set sel = wdDoc.Application.Selection

    ' Insert the text at cursor
    sel.TypeText Text:=fullText

    ' Move selection back to newly inserted text
    sel.MoveLeft Unit:=1, Count:=Len(fullText), Extend:=True
    
    ' Apply green color to whole selection
    sel.Font.Color = RGB(112, 173, 71)
    sel.Font.Bold = True
    
    ' After apply the color and Bold move one position to the right
    sel.MoveRight Unit:=1, Count:=1
    sel.Font.Bold = False
    
    ' Insert the text at cursor
    sel.TypeText Text:=" "

End Sub
