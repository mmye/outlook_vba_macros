VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SmartSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim st As Worksheet
Sub Init(ThisSheet As Worksheet)
    Set st = ThisSheet
End Sub

Sub Pull()

    Dim name As String
    name = InputBox("���Ϗ��ԍ�����͂��Ă�������")
    If name = "" Then Exit Sub
    Dim rQuotationNumber As Range
    Set rQuotationNumber = st.Cells.Find(What:="No.", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, searchdirection:=xlNext, MatchCase:=False _
            , MatchByte:=False, SearchFormat:=False)
    rQuotationNumber.Value = "No. " & name

    Dim rDate As Range
    Set rDate = st.Cells.Find(What:="Date:", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
            xlPart, SearchOrder:=xlByRows, searchdirection:=xlNext, MatchCase:=False _
            , MatchByte:=False, SearchFormat:=False)
    Dim i As Long
    Dim r As Range
    For i = 0 To 10
        Set r = rDate.Offset(0, i)
        If r.Value Like "*�N*��*��" Then r.Value = GetToday
    Next
    Set rQuotationNumber = Nothing
End Sub
Private Function GetToday() As String
    Dim d As String
    d = Format(Date, "Long Date")
    Dim ret
    ret = RegularExpressions.RegexReplace(d, "(\d+)�N(\d+)��(\d+)��", "$1�N $2�� $3��")
    GetToday = ret
End Function

Sub get_tts()
    GetTTS.ShowTTS
End Sub


