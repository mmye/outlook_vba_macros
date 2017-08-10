Attribute VB_Name = "SaveAttachment"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hwnd As Long, ByVal lpszOp As String, _
                 ByVal lpszFile As String, ByVal lpszParams As String, _
                 ByVal LpszDir As String, ByVal FsShowCmd As Long) _
                 As Long

Sub SaveAttachment(ByVal objItem As MailItem)
'Source:
    Dim objAttach As Attachment
    Dim strFileName As String
    Dim c As Integer
    Dim objFSO 'As FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    For Each objAttach In objItem.Attachments
        c = 1
        With objAttach
            strFileName = .FileName
            While objFSO.FileExists(ATTACH_PATH & strFileName)
                strFileName = Left(.FileName, InStrRev(.FileName, ".") - 1) _
                    & "-" & c & Mid(.FileName, InStrRev(.FileName, "."))
                c = c + 1
            Wend
            .SaveAsFile ATTACH_PATH & strFileName
        End With
    Next
End Sub

