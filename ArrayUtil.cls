VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim v As Variant
Function Compress(var As Variant) As Variant
    Dim ret
    ret = IsArrayEx(var)
    Select Case ret
        Case 1
            'OK!
        Case Else
            'Not OK!
            Exit Function
    End Select
    
    Dim Compressed
    Dim arr As Variant
    Dim i, c
    Const s As String = "|"
    On Error GoTo INDEX_LIMIT
    For i = LBound(var, 1) To UBound(var, 1)
        For c = LBound(var, 2) To UBound(var, 2)
            Dim d As Boolean
            d = IsDuplicate(var(i, c), var(i, c + 1))
            If Not d Then arr = arr & s & var(i, c)
        Next c
        If arr <> "" Then Compressed = arr & vbCrLf & Compressed
    Next i

    If Compressed <> "" Then Compress = Compressed
    Exit Function
INDEX_LIMIT:
    If Err.Number = 9 Then Resume Next
End Function

'***********************************************************
' �@�\   : �������z�񂩔��肵�A�z��̏ꍇ�͋󂩂ǂ��������肷��
' ����   : varArray  �z��
' �߂�l : ���茋�ʁi1:�z��/0:��̔z��/-1:�z�񂶂�Ȃ��j
'***********************************************************
Private Function IsArrayEx(varArray As Variant) As Long
On Error GoTo ERROR_

    If IsArray(varArray) Then
        IsArrayEx = IIf(UBound(varArray) >= 0, 1, 0)
    Else
        IsArrayEx = -1
    End If

    Exit Function

ERROR_:
    If Err.Number = 9 Then
        IsArrayEx = 0
    End If
End Function
Private Function IsDuplicate(ValueA, ValueB) As Boolean
    '�Z�����}�[�W����Ă���ƁA�����l���d������B
    '�����Ŏ󂯂��z��̂ƂȂ肠����̒l�̈�v�E�s��v���m�F���ē�����������True��Ԃ�
    If ValueA = Empty Then
        IsDuplicate = True
        Exit Function
    End If
    If ValueA = ValueB Then
        IsDuplicate = True
    Else
        IsDuplicate = False
    End If
End Function
