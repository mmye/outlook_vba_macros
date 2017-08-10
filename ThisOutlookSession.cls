VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    TitlePresenceCheck Item, Cancel
    AttachmentPresenceCheck Item, Cancel
End Sub

Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    Dim objItem As Object
    Set objItem = Session.GetItemFromID(EntryIDCollection)
    If TypeName(objItem) = "MailItem" Then
        SaveAttachment.SaveAttachment objItem
    End If
End Sub

