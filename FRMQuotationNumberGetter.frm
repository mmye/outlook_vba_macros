VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRMQuotationNumberGetter 
   Caption         =   "���͍���"
   ClientHeight    =   8730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12210
   OleObjectBlob   =   "FRMQuotationNumberGetter.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
b = MsgBox("���͂��Ă������ł���", vbQuestion + vbYesNo)
If b = vbYes Then
    InsertData
End If
End Sub
Private Sub cmb�V���A��No_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmb�V���A��No.DropDown
End Sub
Private Sub cmb�q��_Change()
    'ResetForm
    ResetLists
    cmb�q��H��.Text = ""
    cmb���[�J�[.Text = ""
    cmb�q��S��.Text = ""
    cmb�Ώۋ@��.Text = ""
    cmb�V���A��No.Text = ""
    '�v�f��1�݂̂ł���Ύ����œ��͂���
    Set_CustomerFactory
    Set_CustomerPerson
    Set_MachineName
    Set_Makers
End Sub
Private Sub cmb���[�J�[_Change()
    'ResetForm
    If cmb�q��.Text <> "" Then Set_MachineName
End Sub
Private Sub cmb���[�J�[_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmb���[�J�[.DropDown
End Sub
Private Sub cmb�q��_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmb�q��.DropDown
End Sub
Private Sub cmb�q��H��_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmb�q��H��.DropDown
End Sub
Private Sub cmb�q��S��_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmb�q��S��.DropDown
End Sub
Private Sub cmb�Ώۋ@��_Change()
    If cmb�Ώۋ@��.Text <> "" Then Set_�V���A��No
End Sub
Private Sub cmb�Ώۋ@��_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmb�Ώۋ@��.DropDown
End Sub
Private Sub UserForm_Initialize()
    Set_Date
    Set_Makers
    Set_Customers
    Set_PersonInCharge
End Sub
Private Sub Set_CustomerPerson()
    Dim CustomerPerson As Variant
    CustomerPerson = SQLs.Get_CustomerPerson(Me.cmb�q��)
    cmb���[�J�[.List = Array()
    If Util.IsArrayEx(CustomerPerson) <> -1 Then
        Dim i As Long
        For i = LBound(CustomerPerson, 2) To UBound(CustomerPerson, 2)
            cmb�q��S��.AddItem CustomerPerson(0, i) & " �l"
        Next
    End If
End Sub
Private Sub Set_PersonInCharge()
    Dim Owner As String
    Owner = SQLs.GetComputerName
    txt�쐬��.Text = Owner
End Sub
Private Sub Set_Makers()
    If cmb�q��.Text = "" Then Exit Sub
    Dim Makers As Variant
    Makers = SQLs.Get_Manufacturers(Me.cmb�q��)
    cmb���[�J�[.List = Array()
    Dim i As Long
    For i = LBound(Makers, 2) To UBound(Makers, 2)
        cmb���[�J�[.AddItem Makers(0, i)
    Next
    If cmb���[�J�[.ListCount = 1 And cmb���[�J�[.List(0) <> "" Then
        cmb���[�J�[.Text = cmb���[�J�[.List(0)
    Else

    End If
End Sub
Private Sub Set_Customers()
    Dim Customers As Variant
    Customers = SQLs.Get_Customers(Me.cmb�q��)
    cmb�q��.List = Array()
    Dim i As Long
    For i = LBound(Customers, 2) To UBound(Customers, 2)
        cmb�q��.AddItem Customers(0, i)
    Next
End Sub
Private Sub Set_CustomerFactory()
    Dim Factories As Variant
    Factories = Get_CustomerFactories(Me.cmb�q��)
    cmb�q��H��.List = Array()
    Dim i As Long
    For i = LBound(Factories, 2) To UBound(Factories, 2)
        cmb�q��H��.AddItem Factories(0, i)
    Next
    If cmb�q��H��.ListCount = 1 And cmb�q��H��.List(0) <> "" Then
        cmb�q��H��.Enabled = True
        cmb�q��H��.Text = cmb�q��H��.List(0)
    Else
        cmb�q��H��.Enabled = False
    End If
End Sub
Private Function Set_MachineName()
    Dim Machines As Variant
    If cmb���[�J�[.Text = "" Then Exit Function
    Machines = SQLs.Get_MachineNames(Me.cmb�q��, cmb���[�J�[)
    If Util.IsArrayEx(Machines) = 0 Then
        cmb�Ώۋ@��.List = Array()
        cmb�Ώۋ@��.Enabled = False
        Exit Function
    Else
        cmb�Ώۋ@��.Enabled = True
    End If
    On Error Resume Next
    Dim i As Long
    For i = LBound(Machines, 2) To UBound(Machines, 2)
        cmb�Ώۋ@��.AddItem Machines(0, i)
    Next
    Exit Function
    On Error GoTo 0
Err:
Debug.Print "Set_MachineName: error..."
End Function
Private Function Set_�V���A��No()
    Dim SerialNos As Variant
    SerialNos = SQLs.Get_MachineId(cmb�q��, cmb���[�J�[, cmb�Ώۋ@��)
    If Util.IsArrayEx(SerialNos) = 0 Then
        cmb�V���A��No.List = Array()
        cmb�V���A��No.Enabled = False
        Exit Function
    Else
        cmb�V���A��No.Enabled = True
    End If
    cmb�V���A��No.List = Array()
    Dim i As Long
    For i = LBound(SerialNos, 2) To UBound(SerialNos, 2)
        cmb�V���A��No.AddItem SerialNos(0, i)
    Next
    If cmb�V���A��No.ListCount = 1 Then cmb�V���A��No.Text = cmb�V���A��No.List(0)
    Exit Function
Err:
Debug.Print "Set_�V���A��No: error..."
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
    txt���ϓ�.Text = Format(Date, "long date")
    txt���ϓ�.Enabled = False
End Sub
    
Sub InsertData()
'15�񂠂邯��13���ڂ����Ȃ�����G���[�ɂȂ��Ă�
    Const TBL_NAME As String = "winckler_quotations"
    Dim SQL As String
    SQL = "INSERT INTO " & TBL_NAME & "(customer_person, customer_name, machine_type," & _
    "machine_id, published_at, description, customer_factory, published_by," & _
    "manufacturer_quotation_id, domestic_supplier," & _
    "quotation_id, main_machine, over_10mil, manufacturer_name) VALUES(" & _
    """" & cmb�q��S��.Text & """" & "," & _
    """" & cmb�q��.Text & """" & "," & _
    """" & cmb�Ώۋ@��.Text & """" & "," & _
    """" & cmb�V���A��No.Text & """" & "," & _
    """" & txt���ϓ�.Text & """" & "," & _
    """" & txt���ϓ��e.Text & """" & "," & _
    """" & cmb�q��H��.Text & """" & "," & _
    """" & txt�쐬��.Text & """" & "," & _
    """" & txtQuotationno.Text & """" & "," & _
    """" & txt�����d����.Text & """" & "," & _
    """" & txt���ϔԍ�.Text & """" & "," & _
    """" & chk�{�@.Value & """" & "," & _
    """" & chkOver10mil.Value & """" & "," & _
    """" & cmb���[�J�[.Text & """" & _
    ")"
    Debug.Print SQL
        sqlite_no_ADODB.Insert SQL
End Sub

Private Sub ResetLists()
    Set_CustomerFactory
    cmb�q��H��.List = Array()
    cmb���[�J�[.List = Array()
    cmb�q��S��.List = Array()
    cmb�Ώۋ@��.List = Array()
    cmb�V���A��No.List = Array()
End Sub
