Attribute VB_Name = "AutoCellLabeler"
Option Explicit

Const str�{�@ As String = "�{�@"
Const str1000���ȏ� As String = "1000��" & vbLf & "�ȏ�"
Const str���ϔԍ� As String = "���ϔԍ�"
Const str�q�� As String = "�q��"
Const str�q��S�� As String = "�q��S��"
Const str�q��H�� As String = "�q��H��"
Const str�Ώۋ@�� As String = "�Ώۋ@��"
Const str�@�B�ԍ� As String = "�V���A��No."
Const str���ϓ��e As String = "���ϓ��e"
Const str���ϓ� As String = "���ϓ�"
Const str�쐬�� As String = "�쐬��"
Const str�S���� As String = "�S����"
Const strQuotationNo As String = "Quotation no."
Const str�ǉ� As String = "�ǉ�"
Const str�v���W�F�N�g�i���o�[ As String = "�v���W�F�N�g�i���o�["
Const str�I�[�_�[�i���o�[ As String = "�I�[�_�[�i���o�["
Const str���͓� As String = "���͓�(YY/MM/DD)"
Const str�v���W�F�N�g���e As String = "�v���W�F�N�g���e"
Const str�󒍌��ϔԍ� As String = "�󒍌��ϔԍ�"
Const str�d���� As String = "�����d����"
Const str���l As String = "���l"
Const str���i As String = "���i"
Const str��Ɠ� As String = "��Ɠ�"
Const HeaderRow As Long = 2

Sub LoopSheets()
    Dim st As Worksheet
    For Each st In ActiveWorkbook.Sheets
        LabelColumnn st
    Next st
End Sub

Sub LabelColumnn(st As Worksheet)
    Const HeaderRow As Long = 2
    Dim c As Long
    Dim LastCol As Long
    LastCol = Cells(2, st.Columns.Count).End(xlToLeft).Column
    For c = 1 To LastCol
        Dim r As Range
        Set r = st.Cells(HeaderRow, c)

        Dim Prefix As String
'        If st.Name Like "[\(\)-&�i�j]" Then
            Prefix = RegularExpressions.RegexReplace(st.name, "\(|\)", "")
            Prefix = RegularExpressions.RegexReplace(Prefix, "�i", "")
            Prefix = RegularExpressions.RegexReplace(Prefix, "�j", "")
            Prefix = RegularExpressions.RegexReplace(Prefix, "&", "")
            Prefix = RegularExpressions.RegexReplace(Prefix, "-", "")
 '       End If
        Select Case r.Value
            Case str�{�@
                r.name = Prefix & "_" & str�{�@
            Case str1000���ȏ�
                r.name = Prefix & "_" & str1000���ȏ�
            Case str���ϔԍ�
                r.name = Prefix & "_" & str���ϔԍ�
            Case str�q��
                r.name = Prefix & "_" & str�q��
            Case strQuotationNo
                r.name = Prefix & "_" & strQuotationNo
            Case str�ǉ�
                r.name = Prefix & "_" & str�ǉ�
            Case str�v���W�F�N�g�i���o�[
                r.name = Prefix & "_" & str�v���W�F�N�g�i���o�[
            Case str�I�[�_�[�i���o�[
                r.name = Prefix & "_" & str�I�[�_�[�i���o�[
            Case str���͓�
                r.name = Prefix & "_" & str���͓�
            Case str�q��H��
                r.name = Prefix & "_" & str�q��H��
            Case str�Ώۋ@��
                r.name = Prefix & "_" & str�Ώۋ@��
            Case str�@�B�ԍ�
                r.name = Prefix & "_" & str�@�B�ԍ�
            Case str�v���W�F�N�g���e
                r.name = Prefix & "_" & str�v���W�F�N�g���e
            Case str�󒍌��ϔԍ�
                r.name = Prefix & "_" & str�󒍌��ϔԍ�
            Case str�쐬��
                r.name = Prefix & "_" & str�쐬��
            Case str�S����
                r.name = Prefix & "_" & str�S����
            Case str�d����
                r.name = Prefix & "_" & str�d����
            Case str���l
                r.name = Prefix & "_" & str���l
            Case str���i
                r.name = Prefix & "_" & str���i
            Case str��Ɠ�
                r.name = Prefix & "_" & str��Ɠ�
        End Select
    Next
End Sub





