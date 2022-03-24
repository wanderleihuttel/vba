'Para essa função funcionar corretamente é preciso efetuar o download do utilitário "xpdf-tools" no link abaixo
'e salvar em algum diretório, por exemplo: C:\xpdf-tools\
'https://www.xpdfreader.com/download.html
'O utilitário xpdf-tools, só converte PDF que são gerados a partir de texto (onde é possível selecionar o texto) com imagem vai dar erro.

' Função para converter PDF para TXT
' @author  Wanderlei Hüttel <wanderlei.huttel at gmail.com>
' @name    ConvertPdf2Text
' @param   'string'      vArquivoPDF      String
' @param   'string'      vOptions         String
' @return  'variant'                      Retorna o caminho do arquivo texto ou falso em caso de erro
Function ConvertPdf2Text(ByVal vArquivoPDF As String, Optional ByVal vOptions As String = "-table -eol dos -nopgbrk") As Variant
    
    Dim wsh As Object, fs As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = vbHide 'or whatever suits you best
    Dim errorCode As Integer
    Dim vPdf2Text As String, vCommand As String, vArquivoTXT As String
    
    'Caminho do binário "pdftotext.exe"
    vPdf2text = "C:\xpdf-tools\bin64\pdftotext.exe"
    
    'Verifica se o executável do pdftotext existe
    If Not (fs.FileExists(vPdf2Text)) Then
        MsgBox "O utilitário pdftotext não foi encontrado!" & vbNewLine & vbNewLine & vPdf2Text, vbCritical
        ConvertPDF2TXT = False
        Exit Function
    End If
    
    'Verifica se o arquivo pdf existe
    If Not (fs.FileExists(vArquivoPDF)) Then
        MsgBox "O arquivo PDF não foi encontrado!" & vbNewLine & vbNewLine & vArquivoPDF, vbCritical
        ConvertPDF2TXT = False
        Exit Function
    End If
    
    'Gera um arquivo texto como o mesmo nome do arquivo PDF
    vArquivoTXT = fs.GetParentFolderName(vArquivoPDF) & "\" & fs.GetBaseName(vArquivoPDF) & ".txt"
    
    'Cria o comando para executar a conversão do PDF para TXT
    vCommand = Chr(34) & vPdf2Text & Chr(34) & " " & vOptions & " " & Chr(34) & vArquivoPDF & Chr(34) & " " & Chr(34) & vArquivoTXT & Chr(34)
    
    'Executa o comando e recebe o retorno
    errorCode = wsh.Run(vCommand, 0, waitOnReturn)
    
    If errorCode <> 0 Then
        ConvertPdf2Text = False
        MsgBox "Ocorreu um erro ao abrir ou converter o arquivo PDF!" & vbNewLine & "Contate o suporte técnico!", vbCritical
    Else
        ConvertPdf2Text = vArquivoTXT
    End If

End Function

