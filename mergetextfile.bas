Sub MergeTextFiles()
    
    Dim fs As Object
    Dim vSourceFile As Object, vNewFile As Object
    Dim vDialog As FileDialog
    Dim vPath As String, vFileContent As String
    Dim vFile As Variant
    Dim m As Long
    
    Set fs = CreateObject("Scripting.FileSystemObject")

    'Abre o dialog para seleção de arquivos
    Set vDialog = Application.FileDialog(msoFileDialogFilePicker)
    With vDialog
        .AllowMultiSelect = True
        .Title = "Selecione todos os arquivos texto e clique em OK"
        .InitialView = msoFileDialogViewList
        .Filters.Clear
        .Filters.Add "Arquivos texto", "*.txt"
        .Show
    End With
    
    'Verifica se foram selecionados os arquivos TXT
    If (vDialog.SelectedItems.Count > 0) Then
       
        'Recebe o caminho do diretório
        vPath = fs.GetParentFolderName(vDialog.SelectedItems(1))
        
        'Cria o novo arquivo
        Set vNewFile = fs.CreateTextFile(vPath & "\arquivos_juntados.txt")
        
        'Percorre o array de arquivos selecionados
        For Each vFile In vDialog.SelectedItems
            
            'Lê todo o contéudo do arquivo de origem
            'Salva no arquivo novo e fecha o arquivo de origem
            Set vSourceFile = fs.OpenTextFile(vFile, 1)
            vFileContent = vSourceFile.ReadAll
            vSourceFile.Close
            vNewFile.WriteLine vFileContent
            
        Next vFile
        
        'Fecha o arquivo novo
        vNewFile.Close
        
    Else
        Exit Sub
    End If
    
    m = MsgBox("Arquivos mesclados com sucesso!", vbInformation, "Mensagem")
    
End Sub
