'################################################################################################################
' Funções VBA
' Última atualização - 06/11/2018

' Declaração da Função Sleep do Kernel do Windows
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'################################################################################################################
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
' fnGetUserName            Pegar apenas o nome do usuário
' fnGetFileName            Pegar o nome do usuário
' fnRemoveSpecialChars     Remover acentos ou caracteres especiais
' fnTestRegExp             Testar Expressão regular em uma string
' fnDuplicataParcela       Retornar número de duplicata ou parcela



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



' Função para verificar se determinada string é um número
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnIsNumber
' @param   'string'      sValue           String
' @return  'boolean'                      Retorna true se é número e falso se não for número
Function fnIsNumber(ByVal sValue As String) As Boolean
  
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



' Função para pegar retornar apenas números de uma string
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnOnlyNumbers
' @param   'string'      sValue           String
' @return  'string'                       Retorna apenas os número ou zero
'http://stackoverflow.com/questions/7239328/how-to-find-numbers-from-a-string
Function fnOnlyNumbers(ByVal sValue As String) As String
    
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    objRegex.Global = True
    objRegex.Pattern = "[^\d]+"
    fnOnlyNumbers = objRegex.Replace(sValue, vbNullString)
    If (fnOnlyNumbers = "") Then
        fnOnlyNumber = "0"
    End If
    
End Function



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
    fnTimeDiff = Format(Str(Int((Int((JL / 3600)) Mod 24))), "00") _
        + ":" + Format(Str(Int((Int((JL / 60)) Mod 60))), "00") _
        + ":" + Format(Str(Int((JL Mod 60))), "00")
        
End Function



'################################################################################################################
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



'################################################################################################################
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



'################################################################################################################
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



'################################################################################################################
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
            fnAbrirArquivo = ""
        End If
    End With
    Set fd = Nothing
    
End Function



'################################################################################################################
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
                fnSalvarArquivo = CStr(sPath)
            Next vrtSelectedItem
        Else
            fnSaveDialogFile = ""
        End If
    End With
    Set fd = Nothing

End Function



'################################################################################################################
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



'################################################################################################################
' Função serve retornar o nome do usuário atual
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnGetUserName
' @return  'string'                        nome do usuário
' 28/04/2015
Public Function fnGetUserName() As String

    Const separator = "\"
    Dim i As Integer
    Dim arquivo As String
    strUser = Environ$("USERPROFILE")
    strAux = strUser
    For i = Len(strUser) To 1 Step -1
        If Mid$(strUser, i, 1) = separator Then
            strAux = LCase(Mid$(strUser, i + 1, Len(strUser) - i + 1))
            fnGetUserName = strAux
            Exit Function
        End If
    Next

End Function



'################################################################################################################
' Função para pegar o nome de um arquivo
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnBasename
' @param   'string'      sFileName         arquivo com o caminho completo
' @return  'string'                        apenas nome do arquivo
Public Function fnGetFilename(ByVal sFileName As String, Optional ByVal setCase = True) As String
    
    Const separator = "\"
    Dim i As Integer
    strAux = sFileName
    For i = Len(sFileName) To 1 Step -1
        If Mid$(sFileName, i, 1) = separator Then
            strAux = Mid$(sFileName, i + 1, Len(sFileName) - i + 1)
            If (setCase = True) Then
                strAux = LCase(strAux)
            End If
            fnGetFilename = strAux
            Exit Function
        End If
    Next
    
End Function



'################################################################################################################
' Função para substituir acentos ou caracters especiais usando regex
' @author  Wanderlei Hüttel <wanderlei dot huttel at gmail dot com>
' @name    fnRemoveSpecialChars
' @param   'string'      sValue             String para substituir
' @param   'bool'        bSpecialCharacter  Boolean se é para remover caracteres especiais
' @return  'string'                         String sem acentos ou sem caracteres especiais
Function fnRemoveSpecialChars(ByVal sValue As String, Optional bSpecialCharacter As Boolean = False) As String

    Dim RE As Object
    Dim vPattern As Variant
    If (bSpecialCharacter = False) Then
        vPattern = Array("[áàâãä]|a", "[ÁÀÂÃÄ]|A", "[éèê]|e", "[ÉÈÊ]|E", "[íì]|i", "[ÍÌ]|I", "[óòôõö]|o", "[ÓÒÔÕÖ]|O", "[úùü]|u", "[ÚÙÜ]|U", "ç|c", "Ç|C")
    Else
        vPattern = Array("-|", "/|", "\.|", "º|o", "ª|a", "\\|")
    End If
    
    Set RE = CreateObject("vbscript.regexp")
    RE.Global = True
    For i = 0 To UBound(vPattern)
        aux = Split(vPattern(i), "|")
        RE.Pattern = aux(0)
        sReplaceWith = aux(1)
        sValue = RE.Replace(sValue, sReplaceWith)
    Next i
    fnRemoveSpecialChars = sValue

End Function




'################################################################################################################
' Esta função serve testar uma expressão regular em uma string
'
' @author  unknown name
' @name    fnTestRegExp
' @param   'string'      sMyPattern         expressão regular
' @return  'string'      sMyString          string para verificar
' Exemplo: 
' sMyString = "IS1 is2 IS3 is4"
' sMyPattern = "is."
' retorno = fnTestRegExp(sMyPattern, sMyString)
Public Function fnTestRegExp(sMyPattern As String, sMyString As String) As Boolean

    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
   
    Set objRegExp = New RegExp       ' Create a regular expression object.
    objRegExp.Pattern = sMyPattern   ' Set the pattern by using the Pattern property.
    objRegExp.IgnoreCase = False     ' Set Case Insensitivity.
    objRegExp.Global = True          ' Set global applicability.

    If (objRegExp.Test(sMyString) = True) Then    ' Test whether the String can be compared.
        'Get the matches.
        Set colMatches = objRegExp.Execute(sMyString)   ' Execute search.
        fnTestRegExp = True
    Else
       fnTestRegExp = False
   End If

End Function



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
