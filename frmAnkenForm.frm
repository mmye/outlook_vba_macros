VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAnkenForm 
   ClientHeight    =   9630
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9720
   OleObjectBlob   =   "frmAnkenForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmAnkenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub UserForm_Initialize()
    Dim SQL As String
    
    Dim WincklerBackOffice As Variant
    WincklerBackOffice = sqlite_no_ADODB.SearchAll(SQL)
    cmbClericalPIC.List = WincklerBackOffice
End Sub
