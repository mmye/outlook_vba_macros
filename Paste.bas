Attribute VB_Name = "Paste"
Option Explicit

Sub PasteAsIs()
        
    Dim clipped As String
    With New MSForms.DataObject
        .GetFromClipboard    '�ϐ��̃f�[�^��DataObject�Ɋi�[����
        clipped = .GetText
    End With
    Application.CutCopyMode = False
    
Stop
    Dim s As Variant
    s = Split(clipped, vbCrLf)
    
    On Error GoTo Err
    Dim c As Long: c = 0
    Dim r As Range
    For Each r In Selection
        r.Value = s(c)
        c = c + 1
    Next
    Exit Sub
Err:
MsgBox "���͂ł��Ȃ������Z��������܂�"
End Sub
