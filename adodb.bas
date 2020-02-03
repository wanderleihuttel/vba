Option Explicit
'http://www.heritage-tech.net/908/inserting-data-into-mysql-from-excel-using-vba/
'Enable reference: "Microsoft ActiveX Data Objects 6.0 Library"
Global adoConnection As ADODB.connection
Global rs As ADODB.Recordset


'===============================================================================================================
Public Function ConnectDatabase(connection As String)
    On Error GoTo ErrorHandler
    Dim m As Integer
    
    'Dim adoConnection As String
    Set adoConnection = New ADODB.connection
    
    'Configuração Global Conexão
    'Firebird Example (Needs ODBC driver) https://firebirdsql.org/en/odbc-driver/
    'driver={Firebird/InterBase(r) driver};dbname=10.1.1.1:C:\database\database.fdb;user=sysdba;password=masterkey;c:\firebird\gds32.dll
    adoConnection.ConnectionString = connection
    adoConnection.Open

    Exit Function
ErrorHandler:
    m = MsgBox("Error number: " & Err.Number - vbObjectError & vbNewLine & Err.Description & vbNewLine & vbNewLine & "Contate o suporte técnico de TI!", vbCritical, "Mensagem de erro")
    Set adoConnection = Nothing
    
End Function


'===============================================================================================================
'Function CloseDatabase
Public Function CloseDatabase()
    adoConnection.Close
    Set adoConnection = Nothing
End Function


'===============================================================================================================
'Function Query
Public Function Query(ByVal sql_query As String, ByVal vSheetName As String, Optional ByVal vShowHeader As Boolean = True, Optional ByVal vStartCell As String = "A2")

    Dim Plan As Object
    Dim i As Integer
    Set Plan = fnGetSheetFromCodeName(vSheetName)
    Set rs = New ADODB.Recordset

    rs.CursorLocation = adUseClient
    rs.Open sql_query, adoConnection, adOpenDynamic, adLockOptimistic
    
    'Get header name from columns
    If (vShowHeader = True) Then
        For i = 0 To rs.Fields.Count - 1
            Plan.Cells(1, i + 1) = rs.Fields(i).Name
        Next i
    End If
    
    'Copy result from recorset to sheet
    Plan.Range(vStartCell).CopyFromRecordset rs
    
    'Close recorset
    rs.Close
    
End Function


'===============================================================================================================
'Function Execute
Public Function ExecuteSQL(sql_query As String)
    On Error GoTo ErrorHandler
    Dim m As Integer
    
    adoConnection.BeginTrans
    Set rs = adoConnection.Execute(sql_query)
    adoConnection.CommitTrans
    Set rs = Nothing
    
Exit Function
ErrorHandler:
    m = MsgBox("Error number: " & Err.Number - vbObjectError & vbNewLine & Err.Description & vbNewLine & vbNewLine & "Contate o suporte técnico de TI!", vbCritical, "Mensagem de erro")
    adoConnection.RollbackTrans
    Set rs = Nothing
    
End Function
