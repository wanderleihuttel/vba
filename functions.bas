'===============================================================================================================
' Funções VBA
' Última atualização - 15/05/2023

' Declaração da Função Sleep do Kernel do Windows
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

'===============================================================================================================
' Funções Disponíveis
'
' fnStrMaskCPFCNPJ         Formatar CPF ou CNPJ
' fnValidarCNPJ            Validar CNPJ
' fnValidarCPF             Validar CPF
' fnIsNumber               Verificar se uma string é numérica
' fnOnlyNumbers            Retorna apenas os números de uma string
' fnTimeDiff               Retorna a diferença entre 2 horários
' fnLastDayOfMonth         Retornar o último dia do mês de uma data informada
' fnDeleteMultipleSpaces   Excluir múltiplos espaços de uma string
' fnDeleteMultipleTabs     Excluir múltiplas tabulações de uma string
' fnOpenDialogFile         Abrir Open Dialog File
' fnSaveDialogFile         Abrir Save As Dialog File
' fnFS                     Formatar String TXT
' fnRemoveSpecialChars     Remover acentos ou caracteres especiais
' fnTestRegExp             Testar Expressão regular em uma string
' fnDuplicataParcela       Retornar número de duplicata ou parcela
' fnGetSheetFromCodeName   Usar o nome da planilha interna como variável
' fnStringReverse          Inverte o sentido de uma string
' fnBytesHuman             Retorna o valor de bytes com sufixo (bytes, KB, MB, GB, etc)
' fnStrPos                 Retorna a posição na string onde determinado caracter se encontra
' fnExcelUpdateVBA         Habilitar ou desabilitar atualizações do excel (melhorar o desempenho de cálculo)
' fnStrUTF8ToASCII         Converter texto UTF8 para ASCII
' fnRegexDate              Buscar datas com expressão regular em uma string
' fnRegexReplace           Efetuar substituiçoes em strings com expressões regulares
' fnRegexMatch             Buscar partes de texto usando expressões regulares
' fnShowNamedRange         Exibir ou ocultar Named Ranges
' fnPadLeft                Acrescentar caracteres à esquerda de uma string
' fnPadRight               Acrescentar caracteres à direita de uma string
' fnLookUpX                Buscar dados de uma tabela em outra independente da coluna (fnLookUpX)
' fnJoin                   Retorna uma string a partir de uma range (separada por um caractere delimitado)

'===============================================================================================================
' Função para formatar inscrição Federal (CPF e CNPJ)
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnFCPFCNPJ
' @param   'string'      vString          String
' @return  'string'                       Retorna CPF ou CNPJ formatado
Public Function fnStrMaskCPFCNPJ(ByVal vString As String) As String

    If (Len(vString) = 11) Then
        fnStrMaskCPFCNPJ = Format$(vString, "000\.000\.000\-00")
    ElseIf (Len(sValue) <= 12) Then
        fnStrMaskCPFCNPJ = Format$(vString, "00\.000\.000\/0000\-00")
    Else
        fnStrMaskCPFCNPJ = vString
    End If
    
End Function


'===============================================================================================================
' Esta função serve para validar CNPJ
' Baseado na função PHP de Guilherme Sehn (https://gist.github.com/guisehn/3276302)
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnValidarCNPJ
' @param   'string'      vCNPJ         String contendo o CNPJ
' @return  'boolean'                   Retorna verdadeiro ou falso se validar o CNPJ
Function fnValidarCNPJ(ByVal vCNPJ As String) As Boolean
    Dim i As Integer, j As Integer, vSoma As Integer, vResto As Integer
    Dim vDV1 As String, vDV2 As String
    Dim vNumerosIguais As Boolean
    
    With CreateObject("VBScript.RegExp")
        
        'Remove caracteres que não são números
        .Pattern = "[^0-9]+"
        .Global = True
        vCNPJ = .Replace(vCNPJ, vbNullString)
        
        'Verifica se todos os caracteres são números iguais
        .Pattern = "([0-9])\1{13}"
        vNumerosIguais = .test(vCNPJ)
    End With
    
    'Coloca Zeros à Esquerda
    'vCNPJ = String(14 - Len(vCNPJ), "0") & vCNPJ
    
    'Verifica se é número e se os caracteres não são todos iguais
    If (Not IsNumeric(vCNPJ) Or vNumerosIguais = True Or Len(vCNPJ) <> 14) Then
        fnValidarCNPJ = False
        Exit Function
    End If

    'Validar primeiro dígito verificador
    i = 0
    j = 5
    vSoma = 0
    For i = 1 To 12 Step 1
        vSoma = vSoma + Mid(vCNPJ, i, 1) * j
        j = IIf((j = 2), 9, j - 1)
    Next i
    vResto = vSoma Mod 11
    vDV1 = IIf(vResto < 2, 0, 11 - vResto)
     
    'Validar segundo dígito verificador
    i = 0
    j = 6
    vSoma = 0
    For i = 1 To 13 Step 1
        vSoma = vSoma + Mid(vCNPJ, i, 1) * j
        j = IIf((j = 2), 9, j - 1)
    Next i
    vResto = vSoma Mod 11
    vDV2 = IIf(vResto < 2, 0, 11 - vResto)
    
    If (vDV1 & vDV2 = Mid(vCNPJ, 13, 2)) Then
        fnValidarCNPJ = True
    Else
        fnValidarCNPJ = False
    End If
    
End Function


'===============================================================================================================
' Esta função serve para validar CPF
' Baseado na função PHP de Guilherme Sehn (https://gist.github.com/guisehn/3276015)
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnValidarCPF
' @param   'string'      vCNPJ            String contendo o CPF
' @return  'boolean'                      Retorna verdadeiro ou falso se validar o CPF
Function fnValidarCPF(ByVal vCPF As String) As Boolean
    Dim i As Integer, j As Integer, vSoma As Integer, vResto As Integer
    Dim vDV1 As String, vDV2 As String, vCheck As String
    Dim vNumerosIguais As Boolean
    
    With CreateObject("VBScript.RegExp")
        
        'Remove caracteres que não são números
        .Pattern = "[^0-9]+"
        .Global = True
        vCPF = .Replace(vCPF, vbNullString)
        
        'Verifica se todos os caracteres são números iguais
        .Pattern = "([0-9])\1{10}"
        vNumerosIguais = .test(vCPF)
    End With
    
    'Coloca Zeros à Esquerda
    'vCPF = String(11 - Len(vCPF), "0") & vCPF
    
    'Verifica se é número e se os caracteres não são todos iguais
    If (Not IsNumeric(vCPF) Or vNumerosIguais = True Or Len(vCPF) <> 11) Then
        fnValidarCPF = False
        Exit Function
    End If

    'Validar primeiro dígito verificador
    i = 0
    j = 10
    vSoma = 0
    For i = 1 To 9 Step 1
        vSoma = vSoma + Mid(vCPF, i, 1) * j
        j = j - 1
    Next i
    vResto = vSoma Mod 11
    vDV1 = IIf(vResto < 2, 0, 11 - vResto)
     
    'Validar segundo dígito verificador
    i = 0
    j = 11
    vSoma = 0
    For i = 1 To 10 Step 1
        vSoma = vSoma + Mid(vCPF, i, 1) * j
        j = j - 1
    Next i
    vResto = vSoma Mod 11
    vDV2 = IIf(vResto < 2, 0, 11 - vResto)
    
    If (vDV1 & vDV2 = Mid(vCPF, 10, 2)) Then
        fnValidarCPF = True
    Else
        fnValidarCPF = False
    End If
    
End Function


'===============================================================================================================
' Função para verificar se determinada string é um número
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnIsNumber
' @param   'string'      sValue           String
' @return  'boolean'                      Retorna true se é número e falso se não for número
Public Function fnIsNumber(ByVal vString As String) As Boolean
  
    Dim DP As String
    Dim ts As String
    '   Get local setting for decimal point
    DP = Format$(0, ".")
    '   Get local setting for thousand's separator
    '   and eliminate them. Remove the next two lines
    '   if you don't want your users being able to
    '   type in the thousands separator at all.
    ts = Mid$(Format$(1000, "#,###"), 2, 1)
    vString = Replace$(vString, ts, "")
    '   Leave the next statement out if you don't
    '   want to provide for plus/minus signs
    If vString Like "[+-]*" Then vString = Mid$(vString, 2)
    fnIsNumber = Not vString Like "*[!0-9" & DP & "]*" And Not vString Like "*" & DP & "*" & DP & "*" And Len(vString) > 0 And vString <> DP
  
End Function


'===============================================================================================================
' Função para pegar retornar apenas números de uma string
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnOnlyNumbers
' @param   'string'      vString          String
' @return  'string'                       Retorna apenas números
'http://stackoverflow.com/questions/7239328/how-to-find-numbers-from-a-string
Public Function fnOnlyNumbers(ByVal vString As String) As String

    With CreateObject("VBScript.RegExp")
        .Pattern = "[^0-9]+"
        .Global = True
        fnOnlyNumbers = .Replace(vString, vbNullString)
    End With

End Function


'===============================================================================================================
' Função para contar o tempo entre 2 horários
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnTimeDiff
' @param   'string'      tTimeStart       Data/Hora no formato dd/mm/aaaa hh:mm:ss
' @param   'string'      tTimeFinish      Data/Hora no formato dd/mm/aaaa hh:mm:ss
' @return  'string'                       Duração
Public Function fnTimeDiff(ByVal tTimeStart As Date, ByVal tTimeFinish As Date) As String
   
    Dim DR As Long, DL As Long, JL As Long
    DL = (Hour(tTimeStart) * 3600) + (Minute(tTimeStart) * 60) + (Second(tTimeStart))
    DR = (Hour(tTimeFinish) * 3600) + (Minute(tTimeFinish) * 60) + (Second(tTimeFinish))
    
    If tTimeFinish < tTimeStart Then
        JL = 86400
    Else
        JL = 0
    End If
    JL = JL + (DR - DL)
    fnTimeDiff = Format(CStr(CInt((Int((JL / 3600)) Mod 24))), "00") _
        + ":" + Format(CStr(CInt((Int((JL / 60)) Mod 60))), "00") _
        + ":" + Format(CStr(CInt((JL Mod 60))), "00")
        
End Function


'===============================================================================================================
' Função para retornar o último dia do mês de uma data informada
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnLastDayOfMonth
' @param   'string'      vString          Data no formato dd/mm/aaaa
' @return  'string'                       Último dia do mês
Public Function fnLastDayOfMonth(ByVal vString As String) As String
    
    Dim d As Variant
    d = Split(vString, "/")
    'Acrescenta 1 mês e depois subtrai 1 dia
    fnLastDayOfMonth = DateAdd("d", -1, DateAdd("m", 1, DateSerial(d(2), d(1), 1)))
    
End Function


'===============================================================================================================
' Função para remover múltiplos espaços de uma string
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnDeleteMultipleSpaces
' @param   'string'      vString          String com espaços
' @return  'string'                       String sem múltiplos espaços
Public Function fnDeleteMultipleSpaces(ByVal vString As String) As String
    
    With CreateObject("VBScript.RegExp")
        .Pattern = "\s{2,}"
        .Global = True
        fnDeleteMultipleSpaces = .Replace(vString, Space(1))
    End With
    
End Function


'===============================================================================================================
' Função para remover múltiplas tabulações de uma string
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnDeleteMultipleTabs
' @param   'string'      vString          String com várias tabulações
' @return  'string'                       String sem múltiplas tabulações
Public Function fnDeleteMultipleTabs(ByVal vString As String) As String
    
    With CreateObject("VBScript.RegExp")
        .Pattern = "\t{2,}"
        .Global = True
        fnDeleteMultipleTabs = .Replace(vString, vbTab)
    End With
    
End Function


'===============================================================================================================
' Função para abrir o OpenDialogFile e digitar o nome de um arquivo
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnOpenDialogFile
' @param   'string'      sFileName        Nome do arquivo/diretório inicial
' @param   'integer'     iPath            Caminho padrão inicial
' @return  'string'                       String com o caminho do arquivo
Public Function fnOpenDialogFile(Optional ByVal sFileName As String = "", Optional ByVal iPath As Integer = 0) As String
    
    Dim fd As Office.FileDialog
    Dim vrtSelectedItem As Variant
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Define o caminho padrão para salvar o arquivo
    ' Desktop do usuário
    If (iPath = 1) Then
        sPath = Environ$("USERPROFILE") & "\Desktop\" & sFileName
    ' Caminho especifico
    ElseIf (iPath = 2) Then
        sPath = sFileName
    ' Raiz da planilha
    Else
       sPath = ActiveWorkbook.path & "\" & sFileName
    End If
    
    With fd
        .AllowMultiSelect = False
        .Filters.Add "Todos os Arquivos", "*.*", 1
        .Filters.Add "Arquivos Texto", "*.csv", 1
        .Filters.Add "Arquivos Texto", "*.txt", 1
        .InitialFileName = sPath
        
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                sPath = vrtSelectedItem 'Caminho e nome do arquivo
                fnOpenDialogFile = CStr(sPath)
            Next vrtSelectedItem
        Else
            fnOpenDialogFile = ""
        End If
    End With
    Set fd = Nothing
    
End Function


'===============================================================================================================
' Função para abrir o DialogSaveAs e salvar um arquivo
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnSaveDialogFile
' @param   'string'      sFileName        Nome do arquivo/diretório inicial
' @param   'integer'     iPath            Caminho padrão inicial
' @param   'string'      sExtension       Nome do arquivo/diretório inicial
' @return  'string'                       string com o caminho do arquivo
Public Function fnSaveDialogFile(sFileName As String, _
                                Optional ByVal iPath As Integer = 0, _
                                Optional ByVal sExtension As String = ".txt") As String

    Dim fd As Office.FileDialog
    Dim vrtSelectedItem As Variant
    Dim iFilterIndex As Long
    'Dim fso As New FileSystemObject
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
   
   
    ' Define os filtros de extensão | padrão = texto separado por tabulação '*.txt'
    If (sExtension = ".txt") Then
        IFilter = "Texto (separado por tabulações)"
    ElseIf (sExtension = ".csv") Then
        IFilter = "CSV (separado por vírgulas)"
    Else
        IFilter = "Texto (separado por tabulações)"
    End If
     
    ' Define o caminho padrão para salvar o arquivo'
    'Desktop
    If (iPath = 1) Then
        sPath = Environ$("USERPROFILE") & "\Desktop\" & sFileName & sExtension
    'Caminho específico
    ElseIf (iPath = 2) Then
        sPath = sFileName
    'Raiz da planilha
    Else
        sPath = ActiveWorkbook.path & "\" & sFileName & iExtension
    End If
    
    With fd
        ' Procura pelo índice correto da extensão desejada
        For iFilterIndex = 1 To .Filters.Count
            If (InStr(1, LCase(.Filters(iFilterIndex).Description), LCase(IFilter), vbTextCompare) _
                And (LCase(.Filters(iFilterIndex).Extensions) = "*" & LCase(sExtension))) Then
                .FilterIndex = iFilterIndex
                Exit For
            End If
        Next iFilterIndex
      
        .InitialFileName = sPath
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                sPath = vrtSelectedItem 'Caminho e nome do arquivo
                fnSaveDialogFile = CStr(sPath)
            Next vrtSelectedItem
        Else
            fnSaveDialogFile = ""
        End If
    End With
    Set fd = Nothing

End Function


'===============================================================================================================
' Função para escrever strings de tamanho variáveis, com caracteres de preenchimento e
' alinhamento (utilizado para gerar arquivos texto para layouts)
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnFS
' @param   'string'      vString          string para formatar
' @param   'integer'     vSize            tamanho da string formatada
' @param   'string'      vPosition        posição para alinhar a string (left ou right)
' @param   'string'      vChar            caracter para completar a string (space, zero, etc)
' @return  'string'                       string formatada
' Exemplo:  fnFS("variavel", 10, "R", " ")
Public Function fnFS(ByVal vString As String, ByVal vSize As Integer, ByVal vPosition As String, ByVal vChar As String)

    If (Len(vString) > vSize) Then
        fnFS = Left(vString, vSize)
    Else
        If (vPosition = "R") Then
            fnFS = String(vSize - Len(vString), vChar) & vString
        End If
        If (vPosition = "L") Then
            fnFS = vString & String(vSize - Len(vString), vChar)
        End If
    End If

End Function


'===============================================================================================================
' Função para substituir acentos ou caracters especiais usando regex
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnRemoveSpecialChars
' @param   'string'      sValue             String para substituir
' @param   'bool'        bSpecialCharacter  Boolean se é para remover caracteres especiais
' @return  'string'                         String sem acentos ou sem caracteres especiais
Public Function fnRemoveSpecialChars(ByVal vString As String, Optional ByVal vSpecialCharacter As Boolean = False) As String

    Dim vPattern As Variant
    Dim vArray As Variant
    
    If (vSpecialCharacter = False) Then
        vPattern = Array("[àáâãäå]|a", "[ÀÁÂÃÄÅ]|A", "[éèêë]|e", "[ÉÈÊË]|E", "[ìíîï]|i", "[ÌÍÎÏ]|I", "[óòôõö]|o", "[ÓÒÔÕÖ]|O", "[ùúûü]|u", "[ÙÚÛÜ]|U", "[ç]|c", "[Ç]|C", "[ñ]|n", "[Ñ]|N", "[ýÿ]|y", "[ÝŸ]|Y")
    Else
        vPattern = Array("-|", "/|", "\.|", "º|o", "ª|a", "\\|")
    End If
    
    With CreateObject("VBScript.RegExp")
        .Global = True
        For Each vItem In vPattern
            vArray = Split(vItem, "|")
            .Pattern = vArray(0)
            vReplaceWith = vArray(1)
            vString = .Replace(vString, vReplaceWith)
        Next vItem
        fnRemoveSpecialChars = vString
    End With
    

End Function



'===============================================================================================================
' Esta função serve testar uma expressão regular em uma string
'
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnTestRegExp
' @param   'string'      vPattern         Expressão regular
' @param   'string'      vString          String para verificar
' @return  'boolean'                      Retorna verdadeiro ou falso
' Exemplo:
' vString = "IS1 is2 IS3 is4"
' vPattern = "is."
' retorno = fnTestRegExp(vPattern, vString)
Public Function fnTestRegExp(ByVal vPattern As String, ByVal vString As String) As Boolean

    With CreateObject("VBScript.RegExp")
        .Pattern = vPattern
        .IgnoreCase = False
        .Global = True

        If (.Test(vString) = True) Then
            fnTestRegExp = True
        Else
            fnTestRegExp = False
        End If
    End With

End Function


'===============================================================================================================
' Função para pegar número da duplica e parcela quando quando parcelas são letras
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnDuplicataParcela
' @param   'string'      duplicata        Número da duplicata com parcela
' @param   'string'      tipo             "D" para duplicata ou "P" para parcela
' @return  'string'                       Último dia do mês
' Exemplo: 1234/A
Public Function fnDuplicataParcela(ByVal sDuplicata As String, ByVal sTipoRetorno As String) As String

    Dim str1, str2 As String
    Dim cont As Integer
    cont = 1
    'A 65 - Z 90
    '### Retornar Numero Duplicata ###
    str1 = sDuplicata
    str2 = Right(sDuplicata, 2)
    If (sTipoRetorno = "D") Then
        For i = 65 To 90
            str1 = Replace(str1, "/" & Chr(i), "")
            fnDuplicataParcela = str1
         Next i
    
    '### Retornar Numero Parcela ###
    ElseIf (sTipoRetorno = "P") Then
        For i = 65 To 90
            aux = "/" & Chr(i)
            If (str2 = aux) Then
                fnDuplicataParcela = cont
                Exit For
            Else
                fnDuplicataParcela = 1
            End If
           cont = cont + 1
        Next i
    End If
    
End Function


'===============================================================================================================
' Função para setar a planilha com o nome interno (CodeName)
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnGetSheetFromCodeName
' @param   'string'      sValue          String nome interno da planilha
' @return  'worksheet'                   Objeto Worksheet
' Based on https://www.spreadsheet1.com/vba-codenames.html
' Como usar:
' Dim Plan As Woksheet
' Set Plan = fnGetSheetFromCodeName("SheetCodeName")
Public Function fnGetSheetFromCodeName(ByVal vString As String) As Object
    
    Dim vSheet As Object
    For Each vSheet In ActiveWorkbook.Sheets
        If (vSheet.CodeName = vString) Then
            Set fnGetSheetFromCodeName = vSheet
            Exit For
        End If
    Next vSheet
    
End Function


'===============================================================================================================
' Função para inverter uma string
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnStringReverse
' @param   'string'      vString          String
' @return  'string'                       Retorna a string reversa
Public Function fnStringReverse(ByVal vString As String) As String
    
    fnStringReverse = StrReverse(vString)
    
End Function


'===============================================================================================================
' Função para retornar valor de bytes com sufixo
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnBytesHuman
' @param   'string'      vBytes         Double number
' @return  'string'                     Número com sufixo
Public Function fnBytesHuman(vBytes As Double) As String
    
    Dim vUnits As Variant
    Dim i As Integer
    vBytes = Trim(vBytes)
    i = 0
    vUnits = Array("bytes", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
    Do While (True)
        If (vBytes < 1024) Then
            If (i <= 0) Then
                fnBytesHuman = Round(vBytes, 2) & " " & vUnits(i)
            Else
                fnBytesHuman = FormatNumber(Round(vBytes, 2), 2) & " " & vUnits(i)
            End If
            Exit Do
        End If
        vBytes = vBytes / 1024
        i = i + 1
    Loop
    
End Function


'===============================================================================================================
' Função para retornar a posição de um determinado caracter em uma string string
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnStrPos
' @param   'string'      vString          string para procurar algum valor
' @param   'string'      vChar            valor procurado
' @param   'string'      vPosition        número da ocorrência procurada
'                                         0 ou não informado = 1ª ocorrência
'                                         -1                 =  última ocorrência
'                                         N                  = ocorrência número N
' @return  'integer'                      número da posição onde o caracter foi encontrado
' Exemplo:  fnStrPos("123ABC456BDEF", "B") - retorna 5
Public Function fnStrPos(ByVal vString As String, ByVal vChar As String, Optional ByVal vPosition As Integer = 0) As Integer
    
    Dim vIndex As Integer, vSize As Integer, i As Integer, vCont As Integer
    vSize = Len(vString)
    
    vCont = 0
    For i = 1 To vSize
        'Procura o caracter "vChar" na posição i
        vIndex = InStr(i, vString, vChar)
        
        'Se o vIndex for maior que 0 é porque encontrou
        If (vIndex <> 0) Then
            vCont = vCont + 1
            i = vIndex
            'Se a posição for igual a zero é a primeira ocorrência
            If (vPosition = 0) Then
                Exit For
            'Se a posição for igual ao vCont é a ocorrência N
            ElseIf (vPosition = vCont) Then
                Exit For
            End If
        'Se o vIndex for igual a zero, o vIndex recebe o valor de i
        ElseIf (vIndex = 0 And vPosition <= vCont) Then
            vIndex = i - 1
            Exit For
        End If
    Next i
    fnStrPos = vIndex

End Function


'===============================================================================================================
' Função habilitar e desabilitar atualizações do excel (melhorar o desempenho de cálculo)
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnExcelUpdateVBA
' @param   'boolean'     vOption              True ou False
' Exemplo:  fnExcelUpdateVBA(True)
Public Function fnExcelUpdateVBA(Optional ByVal vOption As Boolean = True)

    With Application
        .Calculation = IIf(vOption = True, xlCalculationAutomatic, xlCalculationManual)
        .ScreenUpdating = vOption
        .DisplayAlerts = vOption
        .EnableAnimations = vOption
        .EnableEvents = vOption
        .EnableCancelKey = IIf(vOption = True, xlInterrupt, xlErrorHandler)
    End With

End Function


'===============================================================================================================
' Função para converter terxto UTF8 para ASCII
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnStrUTF8ToASCII
' @param   'string'                 sInputString
' Exemplo:  sInputString = fnStrUTF8ToASCII(sInputString)
Public Function fnStrUTF8ToASCII(ByVal vString As String) As String
    
    Dim l As Long, sUTF8 As String
    Dim iChar As Integer
    Dim iChar2 As Integer
    On Error Resume Next
    
    For l = 1 To Len(vString)
        iChar = Asc(Mid(vString, l, 1))
        If iChar > 127 Then
            If Not iChar And 32 Then
            iChar2 = Asc(Mid(vString, l + 1, 1))
            sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
            l = l + 1
        Else
            Dim iChar3 As Integer
            iChar2 = Asc(Mid(vString, l + 1, 1))
            iChar3 = Asc(Mid(vString, l + 2, 1))
            sUTF8 = sUTF8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
            l = l + 2
        End If
            Else
            sUTF8 = sUTF8 & Chr$(iChar)
        End If
    Next l
    fnStrUTF8ToASCII = sUTF8
    
End Function


'===============================================================================================================
' Função para buscar datas com regex em uma string
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnRegexDate
' @param   'string'      vString    String contendo datas
' @return  'variant'                Array com todas as datas encontradas em uma string ou falso
Public Function fnRegexDate(ByVal vString As String) As Variant

    Dim vMatch As Object, vAllMatches As Object
    Dim vArrayDate() As String
    Dim i As Integer
    
    With CreateObject("VBScript.RegExp")
        .Pattern = "(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/([12][0-9]{3})"
        .Global = True
        i = 0
        Set vAllMatches = .Execute(vString)
        For Each vMatch In vAllMatches
            If IsDate(vMatch.Value) Then
                ReDim Preserve vArrayDate(i)
                vArrayDate(i) = CStr(CDate(vMatch.Value))
                i = i + 1
            End If
        Next
        
        If (vAllMatches.Count > 0) Then
            fnRegexDate = vArrayDate
        Else
            fnRegexDate = False
        End If
    End With

End Function


'===============================================================================================================
' Função para usar expressões regulares para substituir strings
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnRegexReplace
' @param   'string'      vString      String qualquer
' @param   'string'      vPattern     String com o padrão regex
' @param   'string'      vReplace     String para substituir os matches
' @param   'boolean'     vGlobal      Substituir todas as ocorrências
' @param   'boolean'     vIgnoreCase  Ignorar case
' @param   'boolean'     vMultiLine   Verificar múltiplas linhas
' @return  'string'                   String original ou modificada
Public Function fnRegexReplace(ByVal vString As String, _
                               ByVal vPattern As String, _
                               Optional ByVal vReplace As String = vbNullString, _
                               Optional ByVal vGlobal As Boolean = True, _
                               Optional ByVal vIgnoreCase As Boolean = False, _
                               Optional ByVal vMultiLine As Boolean = True _
                               ) As String
    
    With CreateObject("VBScript.RegExp")
        .Pattern = vPattern
        .IgnoreCase = vIgnoreCase
        .Global = vGlobal
        .MultiLine = vMultiLine
        fnRegexReplace = .Replace(vString, vReplace)
    End With
    
End Function


'===============================================================================================================
' Função para buscar partes de texto usando expressões regulares
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnRegexReplace
' @param   'string'      vString      String qualquer
' @param   'string'      vPattern     String com o padrão regex
' @param   'boolean'     vGlobal      Substituir todas as ocorrências
' @param   'boolean'     vIgnoreCase  Ignorar case
' @param   'boolean'     vMultiLine   Verificar múltiplas linhas
' @return  'string'                   String original ou modificada
Public Function fnRegexMatch(ByVal vString As String, _
                             ByVal vPattern As String, _
                             Optional ByVal vGlobal As Boolean = True, _
                             Optional ByVal vIgnoreCase As Boolean = False, _
                             Optional ByVal vMultiLine As Boolean = True _
                             ) As Variant
    
    Dim vMatch As Object, vAllMatches As Object
    Dim vArrayMatch() As Variant
    Dim i As Integer
    
    With CreateObject("VBScript.RegExp")
        .Pattern = vPattern
        .IgnoreCase = vIgnoreCase
        .Global = vGlobal
        .MultiLine = vMultiLine
        Set vAllMatches = .Execute(vString)
        If (vAllMatches.Count > 0) Then
            ReDim Preserve vArrayMatch(vAllMatches.Count - 1)
            For Each vMatch In vAllMatches
                vArrayMatch(i) = vMatch.Value
                i = i + 1
            Next
            fnRegexMatch = vArrayMatch
        Else
            fnRegexMatch = False
        End If
    End With
    
End Function


'===============================================================================================================
' Função ocultar nome de célula (range name)
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnShowNamedRange
' @param   'boolean'      vOption    True ou False
Public Function fnShowNamedRange(Optional ByVal vOption As Boolean = True)
    
    Dim n As Name
    For Each n In ThisWorkbook.Names
        n.Visible = vOption
    Next n
    
End Function


'===============================================================================================================
' Função para acrescentar caracteres à esquerda de uma string
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnPadLeft
' @param   'string'      vString            String de origem
' @param   'integer'     vSize              Tamanho da nova string
' @param   'string'      vChar              Caracter para preencher a string
' @return  'string'                         String com caracteres à esquerda
Public Function fnPadLeft(ByVal vString As String, vSize As Integer, vChar As String) As String
    
    If (Len(vString) > vSize Or vChar = vbNullString) Then
        fnPadLeft = vString
    Else
        fnPadLeft = Right(String(vSize, vChar) & vString, vSize)
    End If
    
End Function


'===============================================================================================================
' Função para acrescentar caracteres à direita de uma string
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnPadLeft
' @param   'string'      vString            String de origem
' @param   'integer'     vSize              Tamanho da nova string
' @param   'string'      vChar              Caracter para preencher a string
' @return  'string'                         String com caracteres à direita
Public Function fnPadRight(ByVal vString As String, vSize As Integer, vChar As String) As String
    
    If (Len(vString) > vSize Or vChar = vbNullString) Then
        fnPadRight = vString
    Else
        fnPadRight = Left(vString & String(vSize, vChar), vSize)
    End If
    
End Function


'===============================================================================================================
' Função para buscar dados de uma tabela em outra independente da coluna (fnLookUpX)
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnLoopUpX
' @param   'variant'      vSearchValue
' @param   'range'        vRangeSearch
' @param   'range'        vRangeReturn
' @param   'integer'      vColumnReturn
' @param   'Variant'      vValueNotFound
' @return  'variant'      Retorna o resultado da consulta ou erro
Public Function fnLookUpX(ByVal vSearchValue As Variant, _
                        ByVal vRangeSearch As Variant, _
                        ByVal vRangeReturn As Variant, _
                        Optional ByVal vColumnReturn As Integer = 0, _
                        Optional ByVal vValueNotFound As Variant = "#N/D") As Variant
    
    On Error Resume Next
    If (vColumnReturn = 0) Then
        fnLookUpX = Application.WorksheetFunction.Index(vRangeReturn, Application.WorksheetFunction.Match(vSearchValue, vRangeSearch, 0))
    Else
        fnLookUpX = Application.WorksheetFunction.Index(vRangeSearch, Application.WorksheetFunction.Match(vSearchValue, vRangeReturn, 0), vColumnReturn)
    End If
    
    If IsEmpty(fnLookUpX) Then fnLookUpX = vValueNotFound
    Err.Clear

End Function


'===============================================================================================================
' Função para converter uma range para uma string separada por um caracter delimitador
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnJoin
' @param   'range'       vRange             Range de origem
' @param   'char  '      vDelimiter         Caracter delimitador
' @return  'string'                         String separada por com caracter delimitador
Function fnJoin(ByVal vRange As Range, Optional vDelimiter As String = ",") As String
    Dim vResult As String
    Dim r As Range

    vResult = ""
    For Each r In vRange
        If (IsEmpty(r)) Then
            Exit For
        End If
        vResult = vResult & r & vDelimiter
    Next

    fnJoin = Left(vResult, Len(vResult) - 1)
End Function



'===============================================================================================================
' FUNÇÕES PARA MANIPULAÇÃO DA ÁREA DE TRANSFERÊNCIA
'===============================================================================================================
' SetClipBoardText               Copiar um texto para a área de transferência
' GetClipBoardText               Pegar um texto da área de transferência
' ClearClipBoardText             Limpar a área de transferência


' https://www.thespreadsheetguru.com/blog/2015/1/13/how-to-use-vba-code-to-copy-text-to-the-clipboard
' https://www.mrexcel.com/board/threads/vba-post-to-clipboard.1142841/


'===============================================================================================================
' Função para copiar um texto para a área de transferência
' @author  Unknown
' @name    SetClipBoardText
' @param   'variant'      Text
Function SetClipBoardText(ByVal vText As Variant)
    'MSForms 2.0 Object Library
    On Error Resume Next
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText vText
        .PutInClipboard
        .GetText
    End With
End Function

'===============================================================================================================
' Função para pegar um texto da área de transferência
' @name    SetClipBoardText
' @author  Unknown
Function GetClipBoardText()
    'MSForms 2.0 Object Library
    On Error Resume Next
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .GetFromClipboard
        GetClipBoardText = .GetText
    End With
End Function

'===============================================================================================================
' Função para limpar a área de transferência
' @name    ClearClipBoardText
' @author  Unknown
Function ClearClipBoardText()
    'MSForms 2.0 Object Library
    On Error Resume Next
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText ""
        .PutInClipboard
        .GetText
    End With
End Function





'===============================================================================================================
' FUNÇÕES PARA MANIPULAÇÃO DE ARQUIVOS
'===============================================================================================================
' FileExists               Verificar se arquivo existe
' FolderExists             Verificar se diretório existe
' GetFileName              Retorna o nome do arquivo com extensão
' GetBaseName              Retorna o nome do arquivo sem extensão
' GetExtensionName         Retorna a extensão de um arquivo
' GetDriveName             Retorna o drive do caminho especificado
' GetParentFolderName      Retorna o diretório pai do caminho especificado
' GetDesktopPath           Retorna o caminho do Desktop
' GetWorkbookPath          Retorna o caminho da planilha


'===============================================================================================================
' Função para verificar se arquivo existe
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    FileExists
' @param   'string'      FileSpec     Caminho do arquivo
' @return  'boolean'                  Verdadeiro se existir e falso se não existir
Public Function FileExists(ByVal FileSpec As String) As Boolean
    
    With CreateObject("Scripting.FileSystemObject")
        FileExists = .FileExists(FileSpec)
    End With
    
End Function


'===============================================================================================================
' Função para verificar se arquivo existe
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    FolderExists
' @param   'string'      FolderSpec   Caminho do diretório
' @return  'boolean'                  Verdadeiro se existir e falso se não existir
Public Function FolderExists(ByVal FolderSpec As String) As Boolean
    
    With CreateObject("Scripting.FileSystemObject")
        FolderExists = .FolderExists(FolderSpec)
    End With
    
End Function


'===============================================================================================================
' Função para retornar apenas o nome do arquivo com extensão de um diretório/arquivo informado
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    GetFileName
' @param   'string'      path         Caminho do diretório/arquivo
' @return  'string'                   Apenas nome do arquivo com extensão
Public Function GetFileName(ByVal path As String) As String
    
    With CreateObject("Scripting.FileSystemObject")
        GetFileName = .GetFileName(path)
    End With
    
End Function


'===============================================================================================================
' Função para retornar apenas o nome do arquivo sem extensão de um diretório/arquivo informado
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    GetBaseName
' @param   'string'      path         Caminho do diretório/arquivo
' @return  'string'                   Apenas nome do arquivo sem extensão
Public Function GetBaseName(ByVal path As String) As String
    
    With CreateObject("Scripting.FileSystemObject")
        GetBaseName = .GetBaseName(path)
    End With
    
End Function


'===============================================================================================================
' Função para retornar apenas a extensão de um diretório/arquivo informado
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    GetExtensionName
' @param   'string'      path         Caminho do diretório/arquivo
' @return  'string'                   Apenas nome do arquivo sem extensão
Public Function GetExtensionName(ByVal path As String) As String
    
    With CreateObject("Scripting.FileSystemObject")
        GetExtensionName = .GetExtensionName(path)
    End With
    
End Function


'===============================================================================================================
' Função para retornar o drive da unidade de um diretório/arquivo informado
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    GetDriveName
' @param   'string'      path         Caminho do diretório/arquivo
' @return  'string'                   Apenas letra da unidade
Public Function GetDriveName(ByVal path As String) As String
    
    With CreateObject("Scripting.FileSystemObject")
        GetDriveName = .GetDriveName(path)
    End With
    
End Function


'===============================================================================================================
' Função para retornar o diretório pai um diretório/arquivo informado
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    GetParentFolderName
' @param   'string'      path         Caminho do diretório/arquivo
' @return  'string'                   Caminho pai do diretório/arquivo
Public Function GetParentFolderName(ByVal path As String) As String
    
    With CreateObject("Scripting.FileSystemObject")
        GetParentFolderName = .GetParentFolderName(path)
    End With
    
End Function


'===============================================================================================================
' Função para retornar o caminho do Desktop
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    GetDesktopPath
' @return  'string'                   Caminho do Desktop do usuário atual
Public Function GetDesktopPath() As String
    
    With CreateObject("WScript.Shell")
        GetDesktopPath = .SpecialFolders("Desktop")
    End With
    
End Function


'===============================================================================================================
' Função para retornar o caminho da planilha atual
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    GetWorkbookPath
' @return  'string'                   Caminho do planilha atual
Public Function GetWorkbookPath() As String
    
    GetWorkbookPath = ActiveWorkbook.path
    
End Function
