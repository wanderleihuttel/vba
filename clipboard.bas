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
