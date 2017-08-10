Attribute VB_Name = "ItemSendEvents"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hwnd As Long, ByVal lpszOp As String, _
                 ByVal lpszFile As String, ByVal lpszParams As String, _
                 ByVal LpszDir As String, ByVal FsShowCmd As Long) _
                 As Long

Function TitlePresenceCheck(Item As Object, Cancel)
    Dim strSubject As String
    Dim strBody As String
    strSubject = Item.Subject '����
    strBody = Item.Body       '�{��
    ' �����`�F�b�N
    If Trim(strSubject) = "" Then
        If MsgBox("������Y��Ă���\��������܂��B�{���ɑ��M���܂����H", vbYesNo + vbExclamation) = vbNo Then
            Cancel = True
            Exit Function
        End If
    End If
End Function

Function AttachmentPresenceCheck(Item As Object, Cancel)
    Dim strSubject As String
    Dim strBody As String
    strSubject = Item.Subject '����
    strBody = Item.Body       '�{��
    ' �Y�t�t�@�C���`�F�b�N
    If InStr(strSubject & strBody, "�Y�t") > 0 And Item.Attachments.Count = 0 Then
        If MsgBox("�Y�t�t�@�C����Y��Ă���\��������܂��B�{���ɑ��M���܂����H", vbYesNo + vbQuestion) = vbNo Then
            Cancel = True
            Exit Function
        End If
    End If
End Function

Sub SaveAndPrintAttach(ByVal objItem As Object)
    
    objItem.PrintOut
End Sub
