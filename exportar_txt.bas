Option Explicit

Sub ExportarArquivoTexto3()
    
    'Declaração de variáveis
    Dim vArquivo As Variant, vOutput As String, ws As Worksheet
    Dim i As Long, vLastUsedRow As Long
    Dim vData As String, vContaDebito As String, vContaCredito As String, _
        vCodigoHistorico As String, vComplemento As String, vValor As String
    Set ws = ActiveWorkbook.Sheets("Lançamentos")
    
    vArquivo = Application.GetSaveAsFilename(InitialFileName:="arquivo3", filefilter:="Arquivos Texto,*.txt")
    
    If (vArquivo = "" Or vArquivo = False) Then: Exit Sub
    
    With ws
        vOutput = ""
        vLastUsedRow = .Range("A" & Rows.Count).End(xlUp).Row
        For i = 2 To vLastUsedRow
            vData = .Cells(i, 1)                   'Campo Data
            vContaDebito = .Cells(i, 2)            'Campo Conta Débito
            vContaCredito = .Cells(i, 3)           'Campo Conta Crédito
            vCodigoHistorico = .Cells(i, 4)        'Campo Código Histórico
            vComplemento = .Cells(i, 5)            'Campo Complento Histórico
            vValor = FormatNumber(.Cells(i, 6), 2) 'Campo Valor do Lançamento
            
            vOutput = vOutput & vData & "|" & vContaDebito & "|" & vContaCredito & "|" & _
                      vCodigoHistorico & "|" & vComplemento & "|" & vValor & vbNewLine
            
        Next i
    End With
    
    'Abre o arquivo, escreve no arquivo e fecha o arquivo
    Open vArquivo For Output As #1
    Print #1, vOutput;
    Close #1
    
    MsgBox "Arquivo gerado com sucesso!", vbOKOnly + vbInformation, "Mensagem"

End Sub
