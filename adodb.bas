'===============================================================================================================
' ADODB VBA
' Última atualização - 06/06/2022

'http://www.heritage-tech.net/908/inserting-data-into-mysql-from-excel-using-vba/
'Enable reference: "Microsoft ActiveX Data Objects 6.0 Library"
'Firebird ODBC driver - https://firebirdsql.org/en/odbc-driver/
'Firebird connection example: 
'    driver={Firebird/InterBase(r) driver};dbname=192.168.0.1/3050:C:\database\database.fdb;client=c:\firebird\fbclient.dll;user=sysdba;password=masterkey

Option Explicit
Global adoConnection As ADODB.connection
Global rs As ADODB.Recordset


'===============================================================================================================
Public Function ConnectDatabase(vConnectionString As String)
    On Error GoTo ErrorHandler
    
    'Dim adoConnection As String
    Set adoConnection = New ADODB.connection
    
    'Configuração Global Conexão
    adoConnection.ConnectionString = vConnectionString
    adoConnection.Open

    Exit Function
ErrorHandler:
    MsgBox "Error number: " & Err.Number - vbObjectError & vbNewLine & Err.Description & vbNewLine & vbNewLine & "Contate o suporte técnico de TI!", vbCritical, "Mensagem de erro"
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
Public Function Query(ByVal sql_query As String, ByVal vSheetName As String, Optional ByVal vShowHeader As Boolean = True, Optional ByVal vStartCell As String = "A1")

    Dim Plan As Worksheet
    Dim i As Integer, vRowIni As Integer, vColIni As Integer
    Set Plan = fnGetSheetFromCodeName(vSheetName)
    Set rs = New ADODB.Recordset

    rs.CursorLocation = adUseClient
    rs.Open sql_query, adoConnection, adOpenDynamic, adLockOptimistic
    
    With Plan
        vColIni = .Range(vStartCell).Column
        vRowIni = .Range(vStartCell).Row
        
        'Get header name from columns
        If (vShowHeader = True) Then
            For i = vColIni To vColIni + rs.Fields.Count - 1
                .Cells(vRowIni, i) = rs.Fields(i - vColIni).Name
            Next i
            vRowIni = vRowIni + 1
        End If
        
        'Copy result from recorset to sheet
        .Cells(vRowIni, vColIni).CopyFromRecordset rs
    End With
    
    'Close recorset
    rs.Close
    
End Function


'===============================================================================================================
'Function Execute
Public Function ExecuteSQL(sql_query As String)
    On Error GoTo ErrorHandler
    
    adoConnection.BeginTrans
    Set rs = adoConnection.Execute(sql_query)
    adoConnection.CommitTrans
    Set rs = Nothing
    
Exit Function
ErrorHandler:
    MsgBox "Error number: " & Err.Number - vbObjectError & vbNewLine & Err.Description & vbNewLine & vbNewLine & "Contate o suporte técnico de TI!", vbCritical, "Mensagem de erro"
    adoConnection.RollbackTrans
    Set rs = Nothing
    
End Function
