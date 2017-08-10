Attribute VB_Name = "Convertpdf2txt"
Option Explicit

Private Enum Conv
    TypeDoc = 0
    TypeDocx = 1
    TypeEps = 2
    TypeHtml = 3
    TypeJpeg = 4
    TypeJpf = 5
    typepdfA = 6
    TypePdfE = 7
    TypePdfX = 8
    TypePng = 9
    TypePs = 10
    TypeRft = 11
    TypeTiff = 12
    TypeTxtA = 13
    TypeTxtP = 14
    TypeXlsx = 15
    TypeSpreadsheet = 16
    TypeXml = 17
End Enum

Public Function PDF(ByVal myFile As String, Optional IsRetry As Boolean, Optional ConvertType As String) As String
    If myFile = "" Then Exit Function
    
    Select Case IsRetry
        Case False
            If ConvertType = "" Then ConvertType = TypeTxtA
        Case True
            ConvertType = TypeDocx
    End Select
    
    PDF = Convertpdf2txt(myFile, ConvertType, IsRetry)
End Function

Private Function Convertpdf2txt(ByVal Fullpath As String, _
                       ByVal ConvertType As Conv, IsRetry As Boolean) As String  'Return converted file name

    Dim jso As Object
    Dim ext As String
    Dim fp As String
    Dim fn As String
    Dim File As String

    'Acrobat 7,8,9,10,11 �̎�
    'Make sure "Acrobat" reference is enabled in case of error
    Dim objAcroApp     As New Acrobat.AcroApp
    Dim objAcroAVDoc As New Acrobat.AcroAVDoc
    Dim objAcroPDDoc As Acrobat.AcroPDDoc

    '�ȍ~��Acrobat�S�ċ���
    Dim lRet As Long    '�߂�l

    With CreateObject("Scripting.FileSystemObject")
        fp = AddPathSeparator(.GetParentFolderName(Fullpath))
        fn = .GetBaseName(Fullpath)
    End With
    ext = GetExtension(ConvertType)
    File = fp & fn & "." & ext

    'Acrobat�A�v���P�[�V�������N������B
    lRet = objAcroApp.Show
    lRet = objAcroAVDoc.Open(Fullpath, "")
    
    If Not lRet Then Exit Function
    
    'PDDoc�I�u�W�F�N�g���擾����B
    Set objAcroPDDoc = objAcroAVDoc.GetPDDoc()
    'JavaScript�I�u�W�F�N�g���쐬����B
    Set jso = objAcroPDDoc.GetJSObject

    If IsRetry Then
        'PDF��Word�ɕϊ�����B
        jso.SaveAs File, "com.adobe.acrobat.docx"
    Else
        'PDF���A�N�Z�X�e�L�X�g(accesstext)�ɕϊ�����B
        jso.SaveAs File, "com.adobe.acrobat.accesstext"
        'PDF���v���[���e�L�X�g(plain-text)�ɕϊ�����B
        'jso.SaveAs File, "com.adobe.acrobat.plain-text"
    End If
                     
    'PDDoc�I�u�W�F�N�g��GetFlags���\�b�h�̖߂�l��\������B
    'MsgBox "AcroPDDoc.GetFlags(1)=" & objAcroPDDoc.GetFlags
     
    'PDF�t�@�C������܂��B
    lRet = objAcroAVDoc.Close(1)
     
    'Acrobat�A�v���P�[�V�������I������B
    lRet = objAcroApp.Hide
    lRet = objAcroApp.Exit

    '�I�u�W�F�N�g�������������
    Set objAcroPDDoc = Nothing
    Set objAcroAVDoc = Nothing
    Set objAcroApp = Nothing
    
    Convertpdf2txt = File

End Function

Private Function GetConvID(ByVal ConvType As Conv) As String
  Dim v As Variant

  v = Array("com.adobe.acrobat.doc", "com.adobe.acrobat.docx", "com.adobe.acrobat.eps", _
            "com.adobe.acrobat.html", "com.adobe.acrobat.jpeg", "com.adobe.acrobat.jp2k", _
            "com.callas.preflight.pdfa", "com.callas.preflight.pdfe", "com.callas.preflight.pdfx", _
            "com.adobe.acrobat.png", "com.adobe.acrobat.ps", "com.adobe.acrobat.rtf", _
            "com.adobe.acrobat.tiff", "com.adobe.acrobat.accesstext", "com.adobe.acrobat.plain-text", _
            "com.adobe.acrobat.xlsx", "com.adobe.acrobat.spreadsheet", "com.adobe.acrobat.xml-1-00")
  GetConvID = v(ConvType)
End Function

Private Function GetExtension(ByVal ConvType As Conv) As String
  Dim v As Variant

  v = Array("doc", "docx", "eps", "html", "jpeg", "jpf", "pdf", "pdf", "pdf", "png", _
            "ps", "rft", "tiff", "txt", "txt", "xlsx", "xml", "xml")
  GetExtension = v(ConvType)
End Function

Private Function AddPathSeparator(ByVal s As String)
  If Right$(s, 1) <> ChrW(92) Then s = s & ChrW(92)
  AddPathSeparator = s
End Function
