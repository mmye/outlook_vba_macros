Attribute VB_Name = "PrintoutUtil"
Option Explicit

Sub Print_pages()
    Dim print_range As String
    print_range = InputBox("Please input which pages to print...")
    If print_range = "" Or Not print_range Like "[0-9-]" Then
        MsgBox "Please input valid page range."
    End If
    
    SendKeys "{ENTER}"
    SendKeys "%FPR"
    SendKeys "%S"
    SendKeys print_range
    SendKeys "{ENTER}"
    
End Sub
