VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRMQuotationNumberGetter 
   Caption         =   "入力項目"
   ClientHeight    =   8730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12210
   OleObjectBlob   =   "FRMQuotationNumberGetter.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FRMQuotationNumberGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const DB_FILE_NAME As String = "winckler.sqlite3"
Dim ctrl As Controls
Private Sub btnClose_Click()
    Unload Me
End Sub
Private Sub btnEnter_Click()
Dim b As VbMsgBoxResult
b = MsgBox("入力してもいいですか", vbQuestion + vbYesNo)
If b = vbYes Then
    InsertData
End If
End Sub
Private Sub cmbシリアルNo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmbシリアルNo.DropDown
End Sub
Private Sub cmb客先_Change()
    'ResetForm
    ResetLists
    cmb客先工場.Text = ""
    cmbメーカー.Text = ""
    cmb客先担当.Text = ""
    cmb対象機器.Text = ""
    cmbシリアルNo.Text = ""
    '要素が1つのみであれば自動で入力する
    Set_CustomerFactory
    Set_CustomerPerson
    Set_MachineName
    Set_Makers
End Sub
Private Sub cmbメーカー_Change()
    'ResetForm
    If cmb客先.Text <> "" Then Set_MachineName
End Sub
Private Sub cmbメーカー_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmbメーカー.DropDown
End Sub
Private Sub cmb客先_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmb客先.DropDown
End Sub
Private Sub cmb客先工場_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmb客先工場.DropDown
End Sub
Private Sub cmb客先担当_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmb客先担当.DropDown
End Sub
Private Sub cmb対象機器_Change()
    If cmb対象機器.Text <> "" Then Set_シリアルNo
End Sub
Private Sub cmb対象機器_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmb対象機器.DropDown
End Sub
Private Sub UserForm_Initialize()
    Set_Date
    Set_Makers
    Set_Customers
    Set_PersonInCharge
End Sub
Private Sub Set_CustomerPerson()
    Dim CustomerPerson As Variant
    CustomerPerson = SQLs.Get_CustomerPerson(Me.cmb客先)
    cmbメーカー.List = Array()
    If Util.IsArrayEx(CustomerPerson) <> -1 Then
        Dim i As Long
        For i = LBound(CustomerPerson, 2) To UBound(CustomerPerson, 2)
            cmb客先担当.AddItem CustomerPerson(0, i) & " 様"
        Next
    End If
End Sub
Private Sub Set_PersonInCharge()
    Dim Owner As String
    Owner = SQLs.GetComputerName
    txt作成者.Text = Owner
End Sub
Private Sub Set_Makers()
    If cmb客先.Text = "" Then Exit Sub
    Dim Makers As Variant
    Makers = SQLs.Get_Manufacturers(Me.cmb客先)
    cmbメーカー.List = Array()
    Dim i As Long
    For i = LBound(Makers, 2) To UBound(Makers, 2)
        cmbメーカー.AddItem Makers(0, i)
    Next
    If cmbメーカー.ListCount = 1 And cmbメーカー.List(0) <> "" Then
        cmbメーカー.Text = cmbメーカー.List(0)
    Else

    End If
End Sub
Private Sub Set_Customers()
    Dim Customers As Variant
    Customers = SQLs.Get_Customers(Me.cmb客先)
    cmb客先.List = Array()
    Dim i As Long
    For i = LBound(Customers, 2) To UBound(Customers, 2)
        cmb客先.AddItem Customers(0, i)
    Next
End Sub
Private Sub Set_CustomerFactory()
    Dim Factories As Variant
    Factories = Get_CustomerFactories(Me.cmb客先)
    cmb客先工場.List = Array()
    Dim i As Long
    For i = LBound(Factories, 2) To UBound(Factories, 2)
        cmb客先工場.AddItem Factories(0, i)
    Next
    If cmb客先工場.ListCount = 1 And cmb客先工場.List(0) <> "" Then
        cmb客先工場.Enabled = True
        cmb客先工場.Text = cmb客先工場.List(0)
    Else
        cmb客先工場.Enabled = False
    End If
End Sub
Private Function Set_MachineName()
    Dim Machines As Variant
    If cmbメーカー.Text = "" Then Exit Function
    Machines = SQLs.Get_MachineNames(Me.cmb客先, cmbメーカー)
    If Util.IsArrayEx(Machines) = 0 Then
        cmb対象機器.List = Array()
        cmb対象機器.Enabled = False
        Exit Function
    Else
        cmb対象機器.Enabled = True
    End If
    On Error Resume Next
    Dim i As Long
    For i = LBound(Machines, 2) To UBound(Machines, 2)
        cmb対象機器.AddItem Machines(0, i)
    Next
    Exit Function
    On Error GoTo 0
Err:
Debug.Print "Set_MachineName: error..."
End Function
Private Function Set_シリアルNo()
    Dim SerialNos As Variant
    SerialNos = SQLs.Get_MachineId(cmb客先, cmbメーカー, cmb対象機器)
    If Util.IsArrayEx(SerialNos) = 0 Then
        cmbシリアルNo.List = Array()
        cmbシリアルNo.Enabled = False
        Exit Function
    Else
        cmbシリアルNo.Enabled = True
    End If
    cmbシリアルNo.List = Array()
    Dim i As Long
    For i = LBound(SerialNos, 2) To UBound(SerialNos, 2)
        cmbシリアルNo.AddItem SerialNos(0, i)
    Next
    If cmbシリアルNo.ListCount = 1 Then cmbシリアルNo.Text = cmbシリアルNo.List(0)
    Exit Function
Err:
Debug.Print "Set_シリアルNo: error..."
End Function

Private Sub ResetForm()
    Dim ctrl As control
    For Each ctrl In Me.Controls
        Debug.Print TypeName(ctrl)
        Select Case TypeName(ctrl)
            Case "ComboBox"
                ctrl.List = Array()
            Case "Textbox"
                ctrl.Text = ""
        End Select
    Next ctrl
End Sub
Private Sub Set_Date()
    txt見積日.Text = Format(Date, "long date")
    txt見積日.Enabled = False
End Sub
    
Sub InsertData()
'15列あるけど13項目しかないからエラーになってる
    Const TBL_NAME As String = "winckler_quotations"
    Dim SQL As String
    SQL = "INSERT INTO " & TBL_NAME & "(customer_person, customer_name, machine_type," & _
    "machine_id, published_at, description, customer_factory, published_by," & _
    "manufacturer_quotation_id, domestic_supplier," & _
    "quotation_id, main_machine, over_10mil, manufacturer_name) VALUES(" & _
    """" & cmb客先担当.Text & """" & "," & _
    """" & cmb客先.Text & """" & "," & _
    """" & cmb対象機器.Text & """" & "," & _
    """" & cmbシリアルNo.Text & """" & "," & _
    """" & txt見積日.Text & """" & "," & _
    """" & txt見積内容.Text & """" & "," & _
    """" & cmb客先工場.Text & """" & "," & _
    """" & txt作成者.Text & """" & "," & _
    """" & txtQuotationno.Text & """" & "," & _
    """" & txt国内仕入先.Text & """" & "," & _
    """" & txt見積番号.Text & """" & "," & _
    """" & chk本機.Value & """" & "," & _
    """" & chkOver10mil.Value & """" & "," & _
    """" & cmbメーカー.Text & """" & _
    ")"
    Debug.Print SQL
        sqlite_no_ADODB.Insert SQL
End Sub

Private Sub ResetLists()
    Set_CustomerFactory
    cmb客先工場.List = Array()
    cmbメーカー.List = Array()
    cmb客先担当.List = Array()
    cmb対象機器.List = Array()
    cmbシリアルNo.List = Array()
End Sub
