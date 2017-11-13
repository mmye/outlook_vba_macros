VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRMQueryForm 
   Caption         =   "UserForm1"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14055
   OleObjectBlob   =   "FRMQueryForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FRMQueryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function GetMyRows()
    Dim Owner
    Owner = SQLs.GetComputerName
    Dim SQL As String
    SQL = "SELECT is_ordered, main_machine, quotation_id, customer_name, machine_type, machine_id, published_at " & _
            "FROM winckler_quotations WHERE published_by=" & """" & Owner & """" & _
            "ORDER BY quotation_id"
    Dim v
    GetMyRows = sqlite_no_ADODB.SearchAll(SQL)
End Function

Private Sub cmbSearch_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmbSearch.DropDown
End Sub

Private Sub cmbSelectScope_Change()
    If cmbSelectScope.Text = "見積書" Then
        cmbSearch.List = Array()
        
        Dim SQL As String
        SQL = "SELECT DISTINCT customer_name FROM winckler_quotations ORDER BY customer_name"
        Dim v As Variant
        v = sqlite_no_ADODB.SearchAll(SQL)
        v = Util.Yoko2Tate(v)
        cmbSearch.List = v
    End If
End Sub

Private Sub cmbSelectScope_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmbSelectScope.DropDown
End Sub

Private Sub CommandButton1_Click()
    Unload Me
End Sub


Private Sub UserForm_Initialize()
    Dim Rows
    Rows = GetMyRows
    
    SET_SELECT_SCOPE
    
    Dim InvertedList
    InvertedList = Util.Yoko2Tate(Rows)
    ListBox1.ColumnCount = UBound(Rows, 2) - LBound(Rows, 2)
    ListBox1.List = InvertedList
End Sub
Private Sub cmbSearch_Change()
    Dim Rows As Variant
    If cmbSearch.Text <> "" Then
        Dim Scope As String
        Scope = cmbSelectScope.Text
        Select Case Scope
            Case "見積書"
                Scope = "quotation_index"
            Case "案件表"
                Scope = "winckler_quotations"
        
        End Select
        
        Rows = IncrementalGetRows(cmbSearch.Text, Scope)
    End If
    If Not IsEmpty(Rows) Then
        Dim v
        v = Util.Yoko2Tate(Rows)
        ListBox1.List = v
    Else
        ListBox1.List = Array()
    End If
End Sub

Private Function IncrementalGetRows(Keyword As String, TableName As String) As Variant
    Dim c As control
    Dim ret As String
    Dim WhereKey As String
    Select Case TableName
        Case "quotation_index"
            WhereKey = "quotation_fullpath"
        Case "winckler_quotations"
            For Each c In Frame1.Controls
                If c.Value Then ret = c.name
            Next
            Select Case ret
                Case "optCustomer"
                    WhereKey = "customer_name"
                Case "optMachineId"
                    WhereKey = "machine_id"
                Case "optMachineType"
                    WhereKey = "machine_type"
                Case "optMaker"
                    WhereKey = "manufacturer_name"
                Case "optPerson"
                    WhereKey = "published_by"
            End Select
    End Select
    
    Dim SQL As String
    SQL = "SELECT * FROM " & TableName & " WHERE " & WhereKey & " LIKE " & """" & "%" & Keyword & "%" & """" & " ORDER BY " & WhereKey & " DESC "
    Debug.Print "SQL: " & SQL
    Dim Rows
    IncrementalGetRows = sqlite_no_ADODB.SearchAll(SQL)
End Function

Private Sub SET_SELECT_SCOPE()
    cmbSelectScope.AddItem "見積書"
    cmbSelectScope.AddItem "案件表"
End Sub
