Attribute VB_Name = "sqlite_no_ADODB"
Option Explicit

Public Const DLL_DIR As String = "C:\Users\m.maeyama\Dropbox\sqlite_for_Excel"
'Public Const DLL_DIR As String = "\\LS410D760\share\◆事務\_DB\dll"
'Public Const DB_FILE_DIR As String = "C:\Users\m.maeyama\Dropbox\sqlite_for_Excel\db"
Public Const DB_FILE_DIR As String = "\\LS410D760\share\◆事務\_DB"
Public Const DB_FILE_NAME As String = "winckler.sqlite3"

Public Function SearchAll(SQL As String, Optional DB_FILE_NAME2 As String) As Variant
    Dim InitReturn As Long  'SQLiteDLL
    Dim dbFile As String    'DBファイル
    Dim RetVal As Long      'DBData

    #If Win64 Then
        Debug.Print "Excel is 64bit"
        'SQLiteDLL読み込み
        InitReturn = SQLite3Initialize(DLL_DIR + "\x64")
        'データベースハンドラ定義
        Dim myDbHandle As LongPtr
        'ステートメントハンドラ定義
        Dim myStmtHandle As LongPtr
    #Else
        Debug.Print "Excel is 32bit"
        'SQLiteDLL読み込み
        InitReturn = SQLite3Initialize(DLL_DIR + "\")
        'データベースハンドラ定義
        Dim myDbHandle As Long
        'ステートメントハンドラ定義
        Dim myStmtHandle As Long
    #End If

    'SQLiteDLL読み込み成否
    If InitReturn <> SQLITE_INIT_OK Then
        Debug.Print "Error Initializing SQLite. Error: " & Err.LastDllError
        Exit Function
    End If

    'ファイル名取得
    Dim db_file As String
    If (DB_FILE_NAME2 <> "") Then
        db_file = DB_FILE_NAME2
    Else
        db_file = DB_FILE_NAME
    End If

    dbFile = DB_FILE_DIR + "\" + db_file

    'DBファイルOpen
    RetVal = SQLite3Open(dbFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    Debug.Print myDbHandle

     'SQL statement作成
    RetVal = SQLite3PrepareV2(myDbHandle, SQL, myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal

    'SQL実行
    RetVal = SQLite3Step(myStmtHandle)
    Debug.Print "SQLite3Step returned " & RetVal

    'レコード取得
    Dim Rows As Variant
    Rows = GetRows(myStmtHandle, Rows)
    
    'statement 削除
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal

    'DBファイルClose
    RetVal = SQLite3Close(myDbHandle)
    Debug.Print "SQLite3Close returned " & RetVal
    
    SearchAll = Rows
End Function

Private Function GetRows(dbh, arr) As Variant
    Dim rowCount As Long
    Dim colCount As Long
    Dim colType As Long
    Dim Rows() As Variant
    Dim ret As Long
    Dim i As Long

    rowCount = 0
    'RecordsetオブジェクトのGetRowsメソッドの取得と同じ
    'rows = ***.GetRows()
    Do While ret <> SQLITE_DONE
        If rowCount = 0 Then
            colCount = SQLite3ColumnCount(dbh)
            If colCount = 0 Then Exit Function
            ReDim Rows(colCount - 1, rowCount)
        Else
            ReDim Preserve Rows(colCount - 1, rowCount)
        End If
        For i = 0 To colCount - 1
            colType = SQLite3ColumnType(dbh, i)
            Rows(i, rowCount) = ColumnValue(dbh, i, colType)
        Next
        ret = SQLite3Step(dbh)
        rowCount = rowCount + 1
    Loop
    GetRows = Rows
End Function

'--------------------
'Sqlite3Demo.basより
Private Function ColumnValue(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            ColumnValue = SQLite3ColumnInt32(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_FLOAT:
            ColumnValue = SQLite3ColumnDouble(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_TEXT:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_BLOB:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_NULL:
            ColumnValue = Null
    End Select
End Function

Public Sub Insert(SQL As String)

    Dim InitReturn As Long  'SQLiteDLL
    Dim dbFile As String    'DBファイル
    Dim RetVal As Long      'DBData

    #If Win64 Then
        Debug.Print "Excel is 64bit"
        'SQLiteDLL読み込み
        InitReturn = SQLite3Initialize(DLL_DIR + "\x64")
        'データベースハンドラ定義
        Dim myDbHandle As LongPtr
        'ステートメントハンドラ定義
        Dim myStmtHandle As LongPtr
    #Else
        Debug.Print "Excel is 32bit"
        'SQLiteDLL読み込み
        InitReturn = SQLite3Initialize(DLL_DIR + "\")
        'データベースハンドラ定義
        Dim myDbHandle As Long
        'ステートメントハンドラ定義
        Dim myStmtHandle As Long
    #End If

    'SQLiteDLL読み込み成否
    If InitReturn <> SQLITE_INIT_OK Then
        Debug.Print "Error Initializing SQLite. Error: " & Err.LastDllError
        Exit Sub
    End If

    'ファイル名取得
    dbFile = DB_FILE_DIR + "\" + DB_FILE_NAME

    'DBファイルOpen
    RetVal = SQLite3Open(dbFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    Debug.Print myDbHandle

    'SQLiteDLL読み込み成否
    If InitReturn <> SQLITE_INIT_OK Then
        Debug.Print "Error Initializing SQLite. Error: " & Err.LastDllError
        Exit Sub
    End If

     'SQL statement作成
    RetVal = SQLite3PrepareV2(myDbHandle, SQL, myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    'SQL実行
    RetVal = SQLite3Step(myStmtHandle)
    Debug.Print "SQLite3Step returned " & RetVal

    'statement 削除
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
    
    'DBファイルClose
    RetVal = SQLite3Close(myDbHandle)
    Debug.Print "SQLite3Close returned " & RetVal

End Sub


