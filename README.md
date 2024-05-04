# FileManipulationExcel-VBA

Sub ExtractDataFromAllWorksheets()
    
    Dim targetWS As Worksheet
    Dim targetColumn As Integer
    Dim lastRowTarget As Long
    Dim fileName As Variant
    Dim code As String
    
    ' Define a planilha alvo
    Set targetWS = ThisWorkbook.Sheets("Banco de Dados")
    
    ' Define a coluna alvo para os dados
    targetColumn = 2
    
    ' Inicializa a variável para a linha onde os dados serão colados (começando na linha 2)
    lastRowTarget = 2
    
    ' Solicita ao usuário para selecionar arquivos
    fileName = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Select Files", , True)
    
    ' Verifica se os arquivos foram selecionados
    If Not IsArray(fileName) Then Exit Sub
    
    ' Loop pelos arquivos selecionados
    For Each file In fileName
        Dim sourceWB As Workbook
        Dim sourceWS As Worksheet
        Dim lastRowSource As Long
        Dim sourceRange As Range
        
        ' Abre a pasta de trabalho
        Set sourceWB = Workbooks.Open(file)
        
        ' Extrai o código do nome do arquivo
        code = Mid(sourceWB.Name, InStrRev(sourceWB.Name, "A"), 6)
        
        ' Loop pelas planilhas na pasta de trabalho
        For Each sourceWS In sourceWB.Worksheets
            ' Encontra a última linha com dados na planilha de origem
            lastRowSource = sourceWS.Cells(sourceWS.Rows.Count, "A").End(xlUp).Row
            
            ' Define o intervalo de origem
            Set sourceRange = sourceWS.Range("A2:E" & lastRowSource)
            
            ' Copia os dados para a planilha alvo
            sourceRange.Copy targetWS.Cells(lastRowTarget, targetColumn)
            
            ' Adiciona o código a todas as linhas dos dados colados
            targetWS.Range(targetWS.Cells(lastRowTarget, targetColumn - 1), _
                           targetWS.Cells(lastRowTarget + lastRowSource - 2, targetColumn - 1)).Value = code
            
            ' Atualiza a última linha alvo
            lastRowTarget = lastRowTarget + lastRowSource - 1
        Next sourceWS
        
        ' Fecha a pasta de trabalho sem salvar alterações
        sourceWB.Close SaveChanges:=False
    Next file
    
    MsgBox "Extração de dados concluída.", vbInformation
End Sub
