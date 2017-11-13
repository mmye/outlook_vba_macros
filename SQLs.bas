Attribute VB_Name = "SQLs"
Option Explicit

Function Get_CustomerPerson(ctrl As control) As Variant
    Dim SQL As String
    If ctrl.Text <> "" Then
        SQL = "SELECT customer_person FROM customer_person " & _
              "WHERE customer_name=" & """" & ctrl.Text & """" & "ORDER BY customer_person"
    Else
        Exit Function
    End If
    Get_CustomerPerson = sqlite_no_ADODB.SearchAll(SQL, DB_FILE_NAME)
End Function

Function Get_Customers(ctrl As control) As Variant
    Const SQL As String = "SELECT DISTINCT customer_name FROM delivered_machines ORDER BY customer_name ASC"
    Get_Customers = sqlite_no_ADODB.SearchAll(SQL, DB_FILE_NAME)
End Function

Function Get_CustomerFactories(ctrl As control) As Variant
    Dim SQL As String
    SQL = "SELECT DISTINCT customer_factory FROM delivered_machines " & _
          " WHERE customer_name=" & """" & ctrl.Text & """" & _
          " ORDER BY customer_factory ASC"
    Get_CustomerFactories = sqlite_no_ADODB.SearchAll(SQL, DB_FILE_NAME)
End Function

Function Get_MachineNames(ctrlCustomer As control, ctrlManufacturer As control) As Variant
    Dim SQL As String
    SQL = "SELECT machine_type FROM delivered_machines WHERE customer_name = " & _
          """" & ctrlCustomer.Text & """" & " AND " & _
          " manufacturer_name = " & """" & ctrlManufacturer.Text & """" & _
          " ORDER BY machine_type ASC"
    Get_MachineNames = sqlite_no_ADODB.SearchAll(SQL, DB_FILE_NAME)
End Function

Function Get_MachineId(ctrlCustomer As control, ctrlManufacturer As control, _
                            ctrlMachineName As control) As Variant
    Dim SQL As String
    SQL = "SELECT maker_order_id FROM delivered_machines WHERE customer_name = " & _
          """" & ctrlCustomer.Text & """" & " AND " & _
          " manufacturer_name=" & """" & ctrlManufacturer.Text & """" & _
          " AND " & " machine_type=" & """" & ctrlMachineName.Text & """" & _
          " ORDER BY maker_order_id ASC"
    Debug.Print "sql: " & SQL
    Get_MachineId = sqlite_no_ADODB.SearchAll(SQL, DB_FILE_NAME)
End Function

Function Get_Manufacturers(ctrl As control) As Variant
    Dim SQL As String
    If ctrl.Text <> "" Then
        SQL = "SELECT DISTINCT manufacturer_name FROM delivered_machines " & _
                "WHERE customer_name=" & """" & ctrl.Text & """" & "ORDER BY manufacturer_name"
    Else
        SQL = "SELECT DISTINCT manufacturer_name FROM delivered_machines " & _
                "ORDER BY manufacturer_name"
    End If
    Get_Manufacturers = sqlite_no_ADODB.SearchAll(SQL, DB_FILE_NAME)
End Function

Function GetComputerName() As String
    Dim computerName As String
    computerName = Environ$("ComputerName")
    Dim SQL As String
    SQL = "SELECT owner FROM computers WHERE computer_name =" & """" & computerName & """"
    Dim v
    v = sqlite_no_ADODB.SearchAll(SQL)
    GetComputerName = CStr(v(0, 0))
End Function
