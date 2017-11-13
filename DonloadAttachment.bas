Attribute VB_Name = "DonloadAttachment"
Option Explicit

Sub DownloadAttachemnts()
    Dim Path
    Dim objOutlook
    Set objOutlook = CreateObject("Outlook.Application")
    
    Dim objFolder
    Set objFolder = objOutlook.GetNamespace("MAPI")
    Set objFolder = objOutlook.Session.Folders.Item("m.maeyama@winckler.co.jp")
    Dim ret
    ret = InputBox("Please choose folders to save mail attachments... 1= INBOX 2=SENT ITEMS")
    
    Select Case ret
        Case 1
            Set objFolder = objFolder.Folders.Item(Constants.INBOX)
        Case 2
            Set objFolder = objFolder.Folders.Item(Constants.SENT)
        Case Else
            Exit Sub
    End Select
    
    Dim strToday
    strToday = Left$(CStr(Now()), 4) & _
                Mid$(CStr(Now()), 6, 2) & _
                Mid$(CStr(Now()), 9, 2)

    Dim i
    For i = 1 To objFolder.Items.Count
        Dim objItem
        Set objItem = objFolder.Items.Item(i)
    
        Dim strMailDate
        strMailDate = Left$(CStr(objItem.CreationTime), 4) & _
        Mid$(CStr(objItem.CreationTime), 6, 2) & _
        Mid$(CStr(objItem.CreationTime), 9, 2)

        If (strMailDate = strToday) Then
            Dim j
            Dim FileName As String
            On Error Resume Next
            For j = 1 To objItem.Attachments.Count
                FileName = objItem.Attachments.Item(j).DisplayName
                objItem.Attachments.Item(j).SaveAsFile _
                    Path & FileName
            Next
            On Error GoTo 0
        Set objItem = Nothing
    Next
    
    Set objFolder = Nothing
    Set objOutlook = Nothing

End Sub

Sub DownloadAllAttachemnts()
    Dim Path As String
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    
    Dim objFolder
    Set objFolder = objOutlook.GetNamespace("MAPI")
    Set objFolder = objOutlook.Session.Folders.Item("m.maeyama@winckler.co.jp")
    Set objFolder = objFolder.Folders.Item("受信トレイ")
    Dim i As Long
    For i = 1 To objFolder.Items.Count
        Dim objItem
        Set objItem = objFolder.Items.Item(i)
            Dim j As Long
            Dim FileName As String
            On Error Resume Next
            For j = 1 To objItem.Attachments.Count
                FileName = objItem.Attachments.Item(j).DisplayName
                objItem.Attachments.Item(j).SaveAsFile _
                    Path & FileName
            Next
            On Error GoTo 0
        Set objItem = Nothing
    Next
    
    Set objFolder = Nothing
    Set objOutlook = Nothing

End Sub

Sub PrintFolderNames()
    Dim Path
    Dim objOutlook
    Set objOutlook = CreateObject("Outlook.Application")
    
    Dim objFolder
    Set objFolder = objOutlook.GetNamespace("MAPI")
    Set objFolder = objOutlook.Session.Folders.Item("パブリック フォルダー - m.maeyama@winckler.co.jp")
    
    Dim i As Long
    Dim FolderName As String
    For i = 1 To objFolder.Folders.Count
        FolderName = objFolder.Folders.Item(i).Name
        Debug.Print FolderName
        If FolderName = "すべてのパブリック フォルダー" Then
            Set objFolder = objFolder.Folders.Item(i)
            Dim k As Long
            For k = i To objFolder.Folders.Count
                FolderName = objFolder.Folders.Item(k).Name
                If FolderName = "INCOMING" Then Debug.Print "INCOMINGのメール件数：" & _
                                objFolder.Folders.Item(i).Items.Count
            Next k
        End If
    Next i
    
    Set objFolder = Nothing
    Set objOutlook = Nothing

End Sub

Sub DonwloadAttachmentofAPerson()

    Dim objOutlook
    Set objOutlook = CreateObject("Outlook.Application")
    Const Sender As String = "m.kawai@winckler.co.jp"
    
    Dim objFolder
    Set objFolder = objOutlook.GetNamespace("MAPI")
    Set objFolder = objOutlook.Session.Folders.Item("パブリック フォルダー - m.maeyama@winckler.co.jp")
    
    Dim i As Long
    Dim FolderName As String
    For i = 1 To objFolder.Folders.Count
        FolderName = objFolder.Folders.Item(i).Name
        Debug.Print FolderName
        If FolderName = "すべてのパブリック フォルダー" Then
            Set objFolder = objFolder.Folders.Item(i)
            Dim k As Long
            For k = i To objFolder.Folders.Count
                FolderName = objFolder.Folders.Item(k).Name
                If FolderName = "OUTGOING" Then
                    Set objFolder = objFolder.Folders.Item(k)
                    Dim m As Long
                    For m = 1 To objFolder.Items.Count
                    DoEvents
                        Dim objItem As Object
                        Set objItem = objFolder.Items.Item(m)
'                        Debug.Print TypeName(objItem)
'                        Debug.Print objItem.Subject
'                        Debug.Print objItem.Body
                        If objItem.Subject = Sender Then
                            Call DownloadAttachemet(objItem)
                        End If
                        Set objItem = Nothing
                    Next m
                End If
            Next k
        End If
    Next i
    
    Set objFolder = Nothing
    Set objOutlook = Nothing

End Sub

Private Sub DownloadAttachemet(Item)
    Dim FileName As String
    Dim Path
    Dim j As Long
    On Error Resume Next
    For j = 1 To Item.Attachments.Count
        DoEvents
        FileName = Item.Attachments.Item(j).DisplayName
        Item.Attachments.Item(j).SaveAsFile _
            Path & FileName
    Next
    On Error GoTo 0
End Sub
