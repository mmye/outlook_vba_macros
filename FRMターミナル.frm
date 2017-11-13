VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRMターミナル 
   Caption         =   "見積書を開く..."
   ClientHeight    =   1335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3855
   OleObjectBlob   =   "FRMターミナル.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FRMターミナル"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const strInstruction As String = "見積書番号を入力してEnterを押してください"

Enum RevisionStatus
    FirstRevision = 0
    HasRevisedBefore = 1
End Enum

Private Sub UserForm_Initialize()
    txtStatus.Caption = strInstruction
End Sub

Private Sub txtTerminal_Change()
    Dim buf As String
    buf = txtTerminal.Text
    txtTerminal.Text = UCase(buf)
    If txtStatus.Caption = "" Then txtStatus.Caption = strInstruction
    If txtStatus.Caption <> "" Then
        txtStatus.Caption = Empty
        txtStatus.Caption = strInstruction
    End If
End Sub

Private Sub UserForm_Click()
    txtTerminal.SetFocus
End Sub

Private Sub txtTerminal_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim Key As String
    Dim Path As String, ShName As String

    Dim ID As String
    ID = txtTerminal.Text
    Select Case KeyCode
        Case 27
            Unload Me
            Exit Sub
        Case 13
            If LenB(ID) = 0 Then Exit Sub
            Key = StrConv(ID, vbNarrow)

            Select Case Right$(Key, 2)
                Case "-R"
                    Key = Left$(Key, Len(Key) - 3)
                    Dim IsRevision As Boolean
                    IsRevision = True
                Case "-W"
                    Dim Writable As Boolean
                    Writable = True
                    Key = Left$(Key, Len(Key) - 3)
                End Select
                Key = Trim$(Key)
                Key = UCase(Key)
    End Select

    If Len(Key) = 0 Then Exit Sub

    Dim SQL As String
    SQL = "SELECT * FROM quotation_index WHERE quotation_id LIKE " & """" & ID & "%" & """"
    Debug.Print SQL

    On Error Resume Next
    Dim ret As Variant
    ret = sqlite_no_ADODB.SearchAll(SQL)
    Dim Rows As Variant
    Rows = Util.Yoko2Tate(ret)
    If IsEmpty(Rows) Then
        txtStatus.Caption = Empty
        txtStatus.Caption = "見積書が見つかりません。"
        Exit Sub
    End If

    ShName = Rows(0, 0)
    Path = Rows(0, 1)

    Dim File As String
    File = Dir(Path)
    Select Case File
        Case ""
            '見積書ブックが見つからなかった場合
        Case Else
            On Error Resume Next
            Dim bk As Workbook
            Select Case Writable
                Case True
                    Set bk = Workbooks.Open(fileName:=Path, ReadOnly:=False)
                Case False
                    Set bk = Workbooks.Open(fileName:=Path, ReadOnly:=True)
            End Select

            If IsRevision Then
                If ShName <> "" Then
                    Dim st As Worksheet
                    Application.DisplayAlerts = False
                    Set st = bk.Sheets(ShName).Copy(After:=bk.Sheets(ShName))
                    Application.DisplayAlerts = False

                    'リビジョンのシート名をつけるためごちゃごちゃやってる
                    Dim stName As String
                    stName = bk.Sheets(ShName).name
                    If (Right$(stName, 2) Like "R[0-9]+") Then
                        stName = Replace(st.name, "R", "")
                        st.name = stName & "R" & (Right$(stName, 1) + 1)
                    Else
                        'すでにある改訂番を検知する必要がある
                        'st.Name = st.Name & "R1"
                    End If
                    st.Activate
                    Range("a1").Select
                End If
            Else
                If ShName <> "" Then
                    bk.Sheets(ShName).Activate
                    Range("a1").Select
                End If
                txtStatus.Caption = Empty
                txtStatus.Caption = "見積書" & ShName & "を開きました。"
                Unload Me
            End If
    End Select
    On Error GoTo 0
    Set bk = Nothing
End Sub
