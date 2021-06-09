Sub AlterarLinhaArquivo()

    Dim vArquivo As Variant
    Dim vLinha As String, vOutput As String
    Dim i As Long
    
    vArquivo = Application.GetOpenFilename(filefilter:="Arquivo texto, *.txt, Todos os arquivos, *.*", Title:="Selecione um arquivo texto")

    If (vArquivo = "" Or vArquivo = Null) Then
        Exit Sub
    End If

    'Abre o arquivo e efetua a leitura linha a linha
    i = 1
    vOutput = ""
    Open vArquivo For Input As #1
        While Not EOF(1)
            Line Input #1, vLinha
            'Altera a linha 3
            If (i = 3) Then
                vLinha = "Linha 3 Alterada"
            End If
            vOutput = vOutput & vLinha & vbNewLine
            i = i + 1
        Wend
    Close #1
    
    
    'Abre o arquivo e escreve nele com as alterações
    Open vArquivo For Output As #1
    Print #1, vOutput;
    Close #1

End Sub
