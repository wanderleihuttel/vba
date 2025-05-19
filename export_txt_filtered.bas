Option Explicit

Public Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Sub ExportarArquivoTexto3()
    
    'Declaração de variáveis
    Dim vArquivo As Variant, vOutput As String, ws As Worksheet
    Dim i As Long, vLastUsedRow As Long, vCont As Long
    Dim vRange As Range, vCell As Range
    Dim vData As String, vContaDebito As String, vContaCredito As String, _
        vCodigoHistorico As String, vComplemento As String, vValor As String
    Set ws = ActiveWorkbook.Sheets("Planilha1")
    
    SetCurrentDirectory Environ("USERPROFILE") & "\Desktop"
    vArquivo = Application.GetSaveAsFilename(InitialFileName:="arquivo.txt", FileFilter:="Arquivos Texto,*.txt")
    
    If (vArquivo = "" Or vArquivo = False) Then: Exit Sub
    
    With ws
        vOutput = ""
        vLastUsedRow = .Range("A" & Rows.Count).End(xlUp).Row
        Set vRange = .Range("A2:A" & vLastUsedRow)
        For Each vCell In vRange.SpecialCells(xlCellTypeVisible)
            i = vCell.Row
            vData = CDate(.Cells(i, 1).Value2)            'Campo Data
            vContaDebito = .Cells(i, 2).Value2            'Campo Conta Débito
            vContaCredito = .Cells(i, 3).Value2           'Campo Conta Crédito
            vCodigoHistorico = .Cells(i, 4).Value2        'Campo Código Histórico
            vComplemento = .Cells(i, 5).Value2            'Campo Complento Histórico
            vValor = FormatNumber(.Cells(i, 6).Value2, 2) 'Campo Valor do Lançamento
            vCont = vCont + 1
            
            vOutput = vOutput & vData & "|" & vContaDebito & "|" & vContaCredito & "|" & _
                      vCodigoHistorico & "|" & vComplemento & "|" & vValor & vbNewLine
            
        Next vCell
    End With
    
    'Abre o arquivo, escreve no arquivo e fecha o arquivo
    Open vArquivo For Output As #1
    Print #1, vOutput;
    Close #1
    
    MsgBox "Arquivo gerado com sucesso!" & vbNewLine & vbNewLine & _
           "Foram gerados " & vCont & " registros!", vbOKOnly + vbInformation, "Mensagem"
  
End Sub
