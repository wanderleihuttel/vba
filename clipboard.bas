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
' @return  'boolean'      Retorna verdadeiro ou falso
Function SetClipBoardText(ByVal Text As Variant) As Boolean
    SetClipBoardText = CreateObject("htmlfile").ParentWindow.ClipboardData.SetData("Text", Text)
End Function

'===============================================================================================================
' Função para pegar um texto da área de transferência
' @name    SetClipBoardText
' @author  Unknown
Function GetClipBoardText() As String
    On Error Resume Next
    GetClipBoardText = CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("Text")
End Function

'===============================================================================================================
' Função para limpar a área de transferência
' @name    ClearClipBoardText
' @author  Unknown
Function ClearClipBoardText() As Boolean
    ClearClipBoardText = CreateObject("htmlfile").ParentWindow.ClipboardData.clearData("Text")
End Function

