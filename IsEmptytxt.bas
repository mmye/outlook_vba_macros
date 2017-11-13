Attribute VB_Name = "IsEmptytxt"
Option Explicit

Sub test()
    Const mytxt As String = "C:\Users\mokoo\Desktop\empty.txt"
    Dim l As Boolean
    l = IsEmptytxt(mytxt)
    If l Then
        MsgBox "yes, its empty"
    Else
        MsgBox "No, there's something in that file!!!"
    End If
End Sub

Function IsEmptytxt(Path) As Boolean

    Dim str As String
    On Error GoTo Err
    Open Path For Input As #1
        Line Input #1, str
        Debug.Print str
    Close #1
    
    If Len(str) > 0 Then
        IsEmptytxt = False
    Else
        IsEmptytxt = True
    End If
Exit Function
Err:
Close #1

'Debug.Print Err.Number
'Debug.Print Err.Description
If Err.Number = 62 Then IsEmptytxt = True
End Function
