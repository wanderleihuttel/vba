Sub ImportarCSV()
    On Error GoTo ErrorHandler
    
    'Exemplo CSV 4 Colunas
    'CÓDIGO;NOME;DATA DE NASCIMENTO;UF
    '1;WANDERLEI;28/07/1985;SC
    '2;JOÃO;18/12/1983;SP
    
    Dim Plan As Worksheet
    Set Plan = ActiveWorkbook.Sheets("Plan1")   'Coloque o nome da aba da sua planilha
    Dim vArquivo As Variant, vReg As Variant, vOutput As String, vDiretorio As String, vSeparador As String
    Dim vLinha As String, vColuna1 As String, vColuna2 As String, vColuna3 As String, vColuna4 As String
    Dim i As Long
    Dim vStartTime As Date, vTotalTime As String
    
    'Pega o tempo do início do script
    vStartTime = Now()
    
    'Habilitar velocidade do VBA
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableAnimations = False
        .EnableEvents = False
        .EnableCancelKey = xlInterrupt
    End With
    
    'Diretório padrão Desktop
    vDiretorio = Environ("UserProfile") & "\Desktop"
    ChDrive (vDiretorio)
    ChDir (vDiretorio)
    
    'Nome arquivo padrão
    vArquivo = vDiretorio & "\" & "arquivo.txt"
    
    'Abre a tela de salvar o arquivo
    vArquivo = Application.GetOpenFilename(filefilter:="Arquivos CSV, *.csv")

    'Se arquivo for vazio para a macro
    If (vArquivo = "" Or vArquivo = False) Then
        Exit Sub
    End If
    
    'Limpa os dados da planilha antes de importar
    Plan.Cells.ClearContents
    
    vOutput = ""
    vSeparador = ";"   'O separador pode ser ";" (ponto e vírgula) ou "," (vírgula), ou "|" (pipe), ou vbTab (tab)
    Open vArquivo For Input As #1
    i = 1
    Do While Not EOF(1)
        Line Input #1, vLinha                'Lê o arquivo linha a linha
        
        vReg = Split(vLinha, vSeparador)     'Quebra o arquivo pelo caracter separador
        
        vColuna1 = Trim(vReg(0))             'Atribui o primeiro campo do arquivo à coluna1
        vColuna2 = Trim(vReg(1))             'Atribui o primeiro campo do arquivo à coluna2
        vColuna3 = Trim(vReg(2))             'Atribui o primeiro campo do arquivo à coluna3
        vColuna4 = Trim(vReg(3))             'Atribui o primeiro campo do arquivo à coluna4
        
        Plan.Cells(i, 1) = vColuna1          'Imprime na planilha o valor da coluna1
        Plan.Cells(i, 2) = vColuna2          'Imprime na planilha o valor da coluna2
        Plan.Cells(i, 3) = vColuna3          'Imprime na planilha o valor da coluna3
        Plan.Cells(i, 4) = vColuna4          'Imprime na planilha o valor da coluna4
        
        i = i + 1
        
    Loop

    'Fecha o arquivo
    Close #1
    
    'Calcula o tempo de execução do script
    vTotalTime = Format$(Now() - vStartTime, "hh:mm:ss")
    MsgBox "Arquivo importado com sucesso!" & vbNewLine & vbNewLine & "Tempo de importação: " & vTotalTime & vbNewLine & vbNewLine & vArquivo, vbInformation
    
    'Voltar velocidade do VBA ao normal
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableAnimations = True
        .EnableEvents = True
        .EnableCancelKey = xlErrorHandler
    End With
    
    Exit Sub
    
ErrorHandler:
    'Voltar velocidade do VBA ao normal também em caso de erro
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableAnimations = True
        .EnableEvents = True
        .EnableCancelKey = xlErrorHandler
    End With
    Close #1
    MsgBox "Ocorreu um erro!" & vbNewLine & vbNewLine & "Error number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Mensagem"
    
End Sub
