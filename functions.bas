'===============================================================================================================
' Funções VBA
' Última atualização - 04/02/2020

' Declaração da Função Sleep do Kernel do Windows
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'===============================================================================================================
' Funções Disponíveis
'
' fnFCPFCNPJ               Formatar CPF ou CNPJ
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
' fnSetSheetName           Usar o nome da planilha interna como variável (deprecated)
' fnGetSheetFromCodeName   Usar o nome da planilha interna como variável
' fnStringReverse          Inverte o sentido de uma string
' fnBytesHuman             Retorna o valor de bytes com sufixo (bytes, KB, MB, GB, etc)
' fnStrPos                 Retorna a posição na string onde determinado caracter se encontra
' fnExcelUpdateVBA         Habilitar ou desabilitar atualizações do excel (melhorar o desempenho de cálculo)
' fnStrUTF8ToASCII         Converter texto UTF8 para ASCII
' fnRegexDate              Buscar datas com expressão regular em uma string
' fnRegexReplace           Efetuar substituiçoes em strings com expressões regulares

'===============================================================================================================
' Função para formatar inscrição Federal (CPF e CNPJ)
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnFCPFCNPJ
' @param   'string'      sValue           String
' @return  'string'                       Retorna CPF ou CNPJ formatado
Public Function fnFCPFCNPJ(ByVal sValue As String) As String

    If (Len(sValue) = 11) Then
        fnFCPFCNPJ = Mid(sValue, 1, 3) & "." & Mid(sValue, 4, 3) & "." & Mid(sValue, 7, 3) & "-" & Mid(sValue, 10, 2)
    ElseIf (Len(sValue) = 14) Then
        fnFCPFCNPJ = Mid(sValue, 1, 2) & "." & Mid(sValue, 3, 3) & "." & Mid(sValue, 6, 3) & "/" & Mid(sValue, 9, 4) & "-" & Mid(sValue, 13, 2)
    End If

End Function


'===============================================================================================================
' Função para verificar se determinada string é um número
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnIsNumber
' @param   'string'      sValue           String
' @return  'boolean'                      Retorna true se é número e falso se não for número
Public Function fnIsNumber(ByVal sValue As String) As Boolean
  
    Dim DP As String
    Dim TS As String
    '   Get local setting for decimal point
    DP = Format$(0, ".")
    '   Get local setting for thousand's separator
    '   and eliminate them. Remove the next two lines
    '   if you don't want your users being able to
    '   type in the thousands separator at all.
    TS = Mid$(Format$(1000, "#,###"), 2, 1)
    sValue = Replace$(sValue, TS, "")
    '   Leave the next statement out if you don't
    '   want to provide for plus/minus signs
    If sValue Like "[+-]*" Then sValue = Mid$(sValue, 2)
    fnIsNumber = Not sValue Like "*[!0-9" & DP & "]*" And Not sValue Like "*" & DP & "*" & DP & "*" And Len(sValue) > 0 And sValue <> DP
  
End Function


'===============================================================================================================
' Função para pegar retornar apenas números de uma string
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnOnlyNumbers
' @param   'string'      sValue           String
' @return  'string'                       Retorna apenas os número ou zero
'http://stackoverflow.com/questions/7239328/how-to-find-numbers-from-a-string
Public Function fnOnlyNumbers(ByVal sValue As String) As String
    
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    objRegex.Global = True
    objRegex.Pattern = "[^\d]+"
    fnOnlyNumbers = objRegex.Replace(sValue, vbNullString)
    
End Function


'===============================================================================================================
' Função para contar o tempo entre 2 horários
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnTimeDiff
' @param   'string'      tTimeStart       Data/Hora no formato dd/mm/aaaa hh:mm:ss
' @param   'string'      tTimeFinish      Data/Hora no formato dd/mm/aaaa hh:mm:ss
' @return  'string'                       Duração
Public Function fnTimeDiff(ByVal tTimeStart As Date, ByVal tTimeFinish As Date) As String
    
    Dim DR, DL, JL As Long
    DL = (Hour(tTimeStart) * 3600) + (Minute(tTimeStart) * 60) + (Second(tTimeStart))
    DR = (Hour(tTimeFinish) * 3600) + (Minute(tTimeFinish) * 60) + (Second(tTimeFinish))
    
    If tTimeFinish < tTimeStart Then
        JL = 86400
    Else
        JL = 0
    End If
    JL = JL + (DR - DL)
    fnTimeDiff = Format(str(Int((Int((JL / 3600)) Mod 24))), "00") _
        + ":" + Format(str(Int((Int((JL / 60)) Mod 60))), "00") _
        + ":" + Format(str(Int((JL Mod 60))), "00")
        
End Function


'===============================================================================================================
' Função para retornar o último dia do mês de uma data informada
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnLastDayOfMonth
' @param   'string'      sDate            Data no formato dd/mm/aaaa
' @return  'string'                       Último dia do mês
Public Function fnLastDayOfMonth(ByVal sDate As String) As String
    
    Dim dNewDate As Date
    strAux = Split(sDate, "/")
    dNewDate = DateAdd("m", 1, DateSerial(strAux(2), strAux(1), 1))
    dNewDate = DateAdd("d", -1, dNewDate)
    fnLastDayOfMonth = CStr(dNewDate)
    
End Function


'===============================================================================================================
' Função para remover múltiplos espaços de uma string
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnDeleteMultipleSpaces
' @param   'string'      sValue          String com espaços
' @return  'string'                       String sem múltiplos espaços
Public Function fnDeleteMultipleSpaces(ByVal sValue As String) As String

    Do While InStr(sValue, "  ") > 0
        sValue = Replace(sValue, "  ", " ")
    Loop
    fnDeleteMultipleSpaces = sValue

End Function


'===============================================================================================================
' Função para remover múltiplas tabulações de uma string
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnDeleteMultipleTabs
' @param   'string'      sValue           String com espaços
' @return  'string'                       String sem múltiplas tabulações
Public Function fnDeleteMultipleTabs(ByVal sValue As String) As String

    Do While InStr(sValue, vbTab & vbTab) > 0
        sValue = Replace(sValue, vbTab & vbTab, vbTab)
    Loop
    fnDeleteMultipleTabs = sValue

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
       sPath = ActiveWorkbook.Path & "\" & sFileName
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
        sPath = ActiveWorkbook.Path & "\" & sFileName & iExtension
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
' @param   'string'      sValue           string para formatar
' @param   'integer'     iSize            tamanho da string formatada
' @param   'string'      sPosition        posição para alinhar a string (left ou right)
' @param   'string'      sChar            caracter para completar a string (space, zero, etc)
' @return  'string'                       string formatada
' Exemplo:  fnFS("variavel", 10, "R", " ")
Public Function fnFS(ByVal sValue As String, ByVal iSize As Integer, ByVal sPosition As String, ByVal sChar As String)

    If (Len(sValue) > iSize) Then
        fnFS = Left(sValue, iSize)
    Else
        If (sPosition = "R") Then
            fnFS = String(iSize - Len(sValue), sChar) & sValue
        End If
        If (sPosition = "L") Then
            fnFS = sValue & String(iSize - Len(sValue), sChar)
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
Public Function fnRemoveSpecialChars(ByVal sValue As String, Optional bSpecialCharacter As Boolean = False) As String

    Dim re As Object
    Dim vPattern As Variant
    If (bSpecialCharacter = False) Then
        vPattern = Array("[áàâãä]|a", "[ÁÀÂÃÄ]|A", "[éèê]|e", "[ÉÈÊ]|E", "[íì]|i", "[ÍÌ]|I", "[óòôõö]|o", "[ÓÒÔÕÖ]|O", "[úùü]|u", "[ÚÙÜ]|U", "ç|c", "Ç|C")
    Else
        vPattern = Array("-|", "/|", "\.|", "º|o", "ª|a", "\\|")
    End If
    
    Set re = CreateObject("vbscript.regexp")
    re.Global = True
    For i = 0 To UBound(vPattern)
        aux = Split(vPattern(i), "|")
        re.Pattern = aux(0)
        sReplaceWith = aux(1)
        sValue = re.Replace(sValue, sReplaceWith)
    Next i
    fnRemoveSpecialChars = sValue

End Function



'===============================================================================================================
' Esta função serve testar uma expressão regular em uma string
'
' @author  unknown name
' @name    fnTestRegExp
' @param   'string'      vPattern         expressão regular
' @return  'string'      vString          string para verificar
' Exemplo:
' sMyString = "IS1 is2 IS3 is4"
' sMyPattern = "is."
' retorno = fnTestRegExp(vPattern, vString)
Public Function fnTestRegExp(vPattern As String, vString As String) As Boolean

    Dim re As Object, match As Object, allmatches As Object
   
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = vPattern
    re.IgnoreCase = False
    re.Global = True

    If (re.Test(vString) = True) Then
        'Get the matches.
        Set allmatches = re.Execute(vString)
        fnTestRegExp = True
    Else
       fnTestRegExp = False
   End If

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
' Dim Plan As Object
' Set Plan = fnGetSheetFromCodeName("SheetCodeName")
Public Function fnGetSheetFromCodeName(sCodename As String) As Object
    Dim oSht As Object
    For Each oSht In ActiveWorkbook.Sheets
        If oSht.CodeName = sCodename Then
            Set fnGetSheetFromCodeName = oSht
            Exit For
        End If
    Next oSht
End Function


'===============================================================================================================
' Função para inverter uma string
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnStringReverse
' @param   'string'      strIn            String
' @return  'string'                       Retorna a string reversa
Public Function fnStringReverse(strIn As String) As String
    Dim output As String
    For i = 0 To Len(strIn) - 1
        output = output & Mid(CStr(strIn), Len(CStr(strIn)) - i, 1)
    Next
    fnStringReverse = output
End Function


'===============================================================================================================
' Função para retornar valor de bytes com sufixo
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnBytesHuman
' @param   'string'      bytes          Double number
' @return  'string'                     Número com sufixo
Public Function fnBytesHuman(bytes As Double) As String
    Dim units As Variant
    bytes = Trim(bytes)
    i = 0
    units = Array("bytes", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
    Do While (True)
        If (bytes < 1024) Then
            If (i <= 0) Then
                fnBytesHuman = Round(bytes, 2) & " " & units(i)
            Else
                fnBytesHuman = FormatNumber(Round(bytes, 2), 2) & " " & units(i)
            End If
            Exit Do
        End If
        bytes = bytes / 1024
        i = i + 1
    Loop
End Function



'===============================================================================================================
' Função para retornar a posição de um determinado caracter em uma string string
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnStrPos
' @param   'string'      sValue           string para procurar algum valor
' @param   'string'      sChar            valor procurado
' @param   'string'      iPosition        número da ocorrência procurada
'                                         0 ou não informado = 1ª ocorrência
'                                         -1                 =  última ocorrência
'                                         N                  = ocorrência número N
' @return  'integer'                      número da posição onde o caracter foi encontrado
' Exemplo:  fnStrPos("123ABC456BDEF", "B") - retorna 5
Public Function fnStrPos(ByVal sValue As String, ByVal sChar As String, Optional ByVal iPosition As Integer = 0) As Integer
    
    Dim index, size As Integer
    size = Len(sValue)
    
    cont = 0
    For i = 1 To size
        'Procura o caracter "sChar" na posição i
        index = InStr(i, sValue, sChar)
        
        'Se o index for maior que 0 é porque encontrou
        If (index <> 0) Then
           cont = cont + 1
           i = index
           'Se a posição for igual a zero é a primeira ocorrência
           If (iPosition = 0) Then
                Exit For
           'Se a posição for igual ao cont é a ocorrência N
           ElseIf (iPosition = cont) Then
                Exit For
           End If
        'Se o index for igual a zero, o index recebe o valor de i
        ElseIf (index = 0 And iPosition <= cont) Then
           index = i - 1
           Exit For
        End If
    Next i
    fnStrPos = index

End Function


'===============================================================================================================
' Função habilitar e desabilitar atualizações do excel (melhorar o desempenho de cálculo)
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnExcelUpdateVBA
' @param   'boolean'     opt              True ou False
' Exemplo:  fnExcelUpdateVBA(True)
Public Function fnExcelUpdateVBA(ByVal opt As Boolean)

    If (opt = True) Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    Application.ScreenUpdating = opt
    Application.DisplayAlerts = opt

End Function


'===============================================================================================================
' Função para converter terxto UTF8 para ASCII
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnStrUTF8ToASCII
' @param   'string'                 sInputString
' Exemplo:  sInputString = fnStrUTF8ToASCII(sInputString)
Public Function fnStrUTF8ToASCII(ByVal sInputString As String) As String
    Dim l As Long, sUTF8 As String
    Dim iChar As Integer
    Dim iChar2 As Integer
    On Error Resume Next
    
    For l = 1 To Len(sInputString)
        iChar = Asc(Mid(sInputString, l, 1))
        If iChar > 127 Then
            If Not iChar And 32 Then
            iChar2 = Asc(Mid(sInputString, l + 1, 1))
            sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
            l = l + 1
        Else
            Dim iChar3 As Integer
            iChar2 = Asc(Mid(sInputString, l + 1, 1))
            iChar3 = Asc(Mid(sInputString, l + 2, 1))
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
' @param   'string'      str   String contendo data
' @return  'variant'           Array com todas as datas encontradas em uma string ou falso
Public Function fnRegexDate(ByVal str As String) As Variant
    Dim re As Object, match As Object, AllMatches As Object
    Dim arr_date() As String
    Dim i As Integer
    Set re = CreateObject("vbscript.regexp")
    're.Pattern = "[\d]{2}[\/-][\d]{2}[\/-][\d]{4}"
    re.Pattern = "([0-2][0-9]|(3)[0-1])(\/)(((0)[0-9])|((1)[0-2]))(\/)\d{4}"
    re.Global = True

    i = 0
    Set AllMatches = re.Execute(str)
    For Each match In AllMatches
        If IsDate(match.Value) Then
            ReDim Preserve arr_date(i)
            arr_date(i) = CStr(CDate(match.Value))
            i = i + 1
        End If
    Next
    
    If AllMatches.Count > 0 Then
        fnRegexDate = arr_date
    Else
        fnRegexDate = False
    End If
    Set re = Nothing

End Function


'===============================================================================================================
' Função para usar expressões regulares para substituir strings
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnRegexReplace
' @param   'string'      vString    String qualquer
' @param   'string'      vPattern   String com o padrão regex
' @param   'string'      vReplace   String para substituir os matches
' @return  'string'                 String original ou modificada
Public Function fnRegexReplace(ByVal vString As String, ByVal vPattern As String, ByVal vReplace As String) As Variant
    Dim re As Object, match As Object, AllMatches As Object
    Dim arr_date() As String
    Dim i As Long
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = vPattern
    re.Global = True
    fnRegexReplace = re.Replace(vString, vReplace)
    Set re = Nothing
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
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Dim fso As FileSystemObject
    'Set fso = New FileSystemObject
    FileExists = fso.FileExists(FileSpec)
End Function


'===============================================================================================================
' Função para verificar se arquivo existe
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    FolderExists
' @param   'string'      FolderSpec   Caminho do diretório
' @return  'boolean'                  Verdadeiro se existir e falso se não existir
Public Function FolderExists(ByVal FolderSpec As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(FolderSpec)
End Function


'===============================================================================================================
' Função para retornar apenas o nome do arquivo com extensão de um diretório/arquivo informado
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    GetFileName
' @param   'string'      path         Caminho do diretório/arquivo
' @return  'string'                   Apenas nome do arquivo com extensão
Public Function GetFileName(ByVal path As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileName = fso.GetFileName(path)
End Function


'===============================================================================================================
' Função para retornar apenas o nome do arquivo sem extensão de um diretório/arquivo informado
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    GetBaseName
' @param   'string'      path         Caminho do diretório/arquivo
' @return  'string'                   Apenas nome do arquivo sem extensão
Public Function GetBaseName(ByVal path As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetBaseName = fso.GetBaseName(path)
End Function


'===============================================================================================================
' Função para retornar apenas a extensão de um diretório/arquivo informado
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    GetExtensionName
' @param   'string'      path         Caminho do diretório/arquivo
' @return  'string'                   Apenas nome do arquivo sem extensão
Public Function GetExtensionName(ByVal path As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetExtensionName = fso.GetExtensionName(path)
End Function


'===============================================================================================================
' Função para retornar o drive da unidade de um diretório/arquivo informado
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    GetDriveName
' @param   'string'      path         Caminho do diretório/arquivo
' @return  'string'                   Apenas letra da unidade
Public Function GetDriveName(ByVal path As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetDriveName = fso.GetDriveName(path)
End Function


'===============================================================================================================
' Função para retornar o diretório pai um diretório/arquivo informado
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    GetParentFolderName
' @param   'string'      path         Caminho do diretório/arquivo
' @return  'string'                   Caminho pai do diretório/arquivo
Public Function GetParentFolderName(ByVal path As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetParentFolderName = fso.GetParentFolderName(path)
End Function


'===============================================================================================================
' Função para retornar o caminho do Desktop
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    GetDesktopPath
' @return  'string'                   Caminho do Desktop do usuário atual
Public Function GetDesktopPath() As String
    Dim wso As Object
    Set wso = CreateObject("WScript.Shell")
    GetDesktopPath = wso.SpecialFolders("Desktop")
    Set wso = Nothing
End Function


'===============================================================================================================
' Função para retornar o caminho da planilha atual
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    GetWorkbookPath
' @return  'string'                   Caminho do planilha atual
Public Function GetWorkbookPath() As String
    GetWorkbookPath = ActiveWorkbook.path
End Function
