Attribute VB_Name = "ConvertPDF_AsEditable"
Option Explicit

Sub ConvertPDF()
    Dim myinspector As Outlook.Inspector
    Dim myItem As Outlook.MailItem
    Dim myattachments As Outlook.Attachments
    Dim FileName As String
    
    Set myinspector = Application.ActiveInspector
    If TypeName(myinspector) = "Nothing" Then MsgBox "There is no active inspector."

    Set myItem = myinspector.CurrentItem
    Set myattachments = myItem.Attachments

    'On Error Resume Next
    Dim j
    Dim FilePath As String
    For j = 1 To myItem.Attachments.Count
        DoEvents
        FileName = myItem.Attachments.Item(j).FileName
        FilePath = Constants.Path & FileName
        myItem.Attachments.Item(j).SaveAsFile FilePath
    'On Error GoTo 0
        
        If Right$(FilePath, 4) Like ".pdf" Or Right$(FilePath, 4) Like ".PDF" Then
            
            'In case converting to pptx
            'Dim ConvertType As String
            'ConvertType = AskConvertType
            
            Dim ConvertedFilePath As String
            ConvertedFilePath = Convertpdf2txt.PDF(FilePath, False)
            If ConvertedFilePath = "" Then Exit Sub
            
            Dim IsEmpty As Boolean
            IsEmpty = IsEmptytxt.IsEmptytxt(ConvertedFilePath)
            Select Case IsEmpty
                Case True
                    ConvertedFilePath = Convertpdf2txt.PDF(FilePath, True)
                    With CreateObject("Wscript.Shell")
                        On Error GoTo CANT_OPEN
                        .Run ConvertedFilePath, 5
                        On Error GoTo 0
                    End With
                Case False
                    With CreateObject("Wscript.Shell")
                        .Run ConvertedFilePath, 5
                    End With
            End Select
        End If
    Next
    Exit Sub
CANT_OPEN:
    If Err.Number = -2147024894 Then
        MsgBox "Failed to open the converted file!"
    End If
End Sub

