Attribute VB_Name = "ForwardandRemeveAttachment"
Sub RemoveAttachmentBeforeForwarding()
 
    Dim myinspector As Outlook.Inspector
    Dim myItem As Outlook.MailItem
    Dim myattachments As Outlook.Attachments
    
    Set myinspector = Application.ActiveInspector
    
    If Not TypeName(myinspector) = "Nothing" Then
        Set myItem = myinspector.CurrentItem.Forward
        Set myattachments = myItem.Attachments
    
        While myattachments.Count > 0
            myattachments.Remove 1
        Wend
        
        myItem.Display
        myItem.Recipients.Add "m.maeyama@winckler.co.jp"
        'myItem.Send
    Else
        MsgBox "There is no active inspector."
    End If
 
End Sub
