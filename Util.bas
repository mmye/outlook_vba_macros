Attribute VB_Name = "Util"
Option Explicit

'***********************************************************
' �@�\   : �������z�񂩔��肵�A�z��̏ꍇ�͋󂩂ǂ��������肷��
' ����   : varArray  �z��
' �߂�l : ���茋�ʁi1:�z��/0:��̔z��/-1:�z�񂶂�Ȃ��j
'***********************************************************
Public Function IsArrayEx(varArray As Variant) As Long
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

Function Yoko2Tate(List As Variant) As Variant
    Dim H As Long, W As Long
    H = UBound(List, 1) - LBound(List, 1)
    W = UBound(List, 2) - LBound(List, 2)
    
    Dim Inverted()
    ReDim Inverted(W, H)
    Dim r As Long, c As Long
    For r = LBound(List, 2) To H
        For c = LBound(List, 1) To W
            Inverted(c, r) = List(r, c)
        Next c
    Next r
    Yoko2Tate = Inverted
End Function
