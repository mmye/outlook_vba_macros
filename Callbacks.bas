Attribute VB_Name = "Callbacks"
Option Explicit

Sub Callback_PostIssue(ByVal control As IRibbonControl)
   FRMQuotationNumberGetter.Show False
End Sub
Sub Callback_AnkenQuery(ByVal control As IRibbonControl)
   FRMQueryForm.Show False
End Sub
Sub Callback_Terminal(ByVal control As IRibbonControl)
   FRMターミナル.Show False
End Sub


