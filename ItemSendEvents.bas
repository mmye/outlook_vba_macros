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
    strSubject = Item.Subject '件名
    strBody = Item.Body       '本文
    ' 件名チェック
    If Trim(strSubject) = "" Then
        If MsgBox("件名を忘れている可能性があります。本当に送信しますか？", vbYesNo + vbExclamation) = vbNo Then
            Cancel = True
            Exit Function
        End If
    End If
End Function

Function AttachmentPresenceCheck(Item As Object, Cancel)
    Dim strSubject As String
    Dim strBody As String
    strSubject = Item.Subject '件名
    strBody = Item.Body       '本文
    ' 添付ファイルチェック
    If InStr(strSubject & strBody, "添付") > 0 And Item.Attachments.Count = 0 Then
        If MsgBox("添付ファイルを忘れている可能性があります。本当に送信しますか？", vbYesNo + vbQuestion) = vbNo Then
            Cancel = True
            Exit Function
        End If
    End If
End Function

Sub SaveAndPrintAttach(ByVal objItem As Object)
    
    objItem.PrintOut
End Sub
