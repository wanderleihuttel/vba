Option Explicit
'http://www.heritage-tech.net/908/inserting-data-into-mysql-from-excel-using-vba/
'Enable reference: "Microsoft ActiveX Data Objects 6.0 Library"
Global adoConnection
Global rs


'Dim adoConnection As ADODB.Connection
'Dim rs As ADODB.Recordset

'Function ConnectDatabase
Public Function ConnectDatabase()
    Dim strADOConnection As String
    Set adoConnection = New ADODB.Connection
    
    'Configuração Global Conexão
    strADOConnection = Plan1.Range("A1")
    'Firebird Example (Needs ODBC driver) https://firebirdsql.org/en/odbc-driver/
    'driver={Firebird/InterBase(r) driver};dbname=10.1.1.1:C:\database\database.fdb;user=sysdba;password=masterkey
    adoConnection.ConnectionString = strADOConnection
    adoConnection.Open
End Function


'Function CloseDatabase
Public Function CloseDatabase()
    adoConnection.Close
    Set adoConnection = Nothing
End Function


'Function GetDataFromDB
Sub GetDataFromDB()
    
    Plan2.Cells.ClearContents
    Dim i As Integer
    Dim sql As String
    
    Set rs = New ADODB.Recordset
    Call ConnectDatabase
    
    sql = "SELECT * FROM sometable"
    rs.CursorLocation = adUseClient
    rs.Open sql, adoConnection, adOpenDynamic, adLockOptimistic
    
    'Get header name from columns
    For i = 0 To rs.Fields.count - 1
        Plan2.Cells(1, i + 1) = rs.Fields(i).Name
    Next i
    
    'Copy result from recorset to sheet
    Plan2.Range("A2").CopyFromRecordset rs
    
    rs.Close
    
End Sub
