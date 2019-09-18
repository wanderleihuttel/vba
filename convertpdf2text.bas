'Função para converter PDF para TXT
'https://www.xpdfreader.com/download.html
'Efetuar download xpdf-tools e salvar em algum diretório
Function ConvertPdf2Text(ByVal pdf As Variant, Optional ByVal opt As String = "-table -eol dos -nopgbrk") As Variant
    
    Dim wsh As Object, fs As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = vbHide 'or whatever suits you best
    Dim errorCode As Integer
    Dim pdf2text As String, command As String, txt As String
    
    'Caminho do binário pdftotext.exe
    pdf2text = "C:\xpdf-tools\bin64\pdftotext.exe"
    
    'Verifica se o executável do pdftotext existe
    If Not (fs.FileExists(pdf2text)) Then
        msg = MsgBox("O utilitário pdftotext não foi encontrado!" & vbCrLf & vbCrLf & pdf2text, vbCritical)
        ConvertPDF2TXT = False
        Exit Function
    End If
    
    'Verifica se o arquivo pdf existe
    If Not (fs.FileExists(pdf)) Then
        msg = MsgBox("O arquivo PDF não foi encontrado!" & vbCrLf & vbCrLf & pdf, vbCritical)
        ConvertPDF2TXT = False
        Exit Function
    End If
    
    txt = fs.GetParentFolderName(pdf) & "\" & fs.GetBaseName(pdf) & ".txt"
    
    'Cria o comando para executar
    command = """" & pdf2text & """" & " " & opt & " " & """" & pdf & """" & " " & """" & txt & """"
    
    errorCode = wsh.Run(command, 0, waitOnReturn)
    
    If errorCode <> 0 Then
        ConvertPdf2Text = False
        msg = MsgBox("Ocorreu um erro ao abrir ou converter o arquivo PDF!" & vbCrLf & "Contate o suporte técnico!", vbCritical)
    Else
        ConvertPdf2Text = txt
    End If

End Function
