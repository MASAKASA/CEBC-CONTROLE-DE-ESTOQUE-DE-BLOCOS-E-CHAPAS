Attribute VB_Name = "ExportarArquivos"
Option Explicit

Private msgSucesso As clsErrorStyle
'Botão exportar em pdf as chapas despachadas por motoristas
Private Sub BtnExportarCarregoPDF_Click()

    'Variaveis do metodo
    Dim dataFormatada As String
    Dim nomeArquivo As String
    Dim ListBox As Object
    Dim linhaLista As Double
    Dim linha As Integer
    Dim i As Long
    
    'Atribuições
    linha = 9
    linhaLista = ListBoxCarregos.ListIndex
    
    dataFormatada = ListBoxCarregos.list(linhaLista, 0)
    dataFormatada = Util.ConverterFormatoData(dataFormatada)
    
    nomeArquivo = "Carrego " & ListBoxCarregos.list(linhaLista, 1) & " " & dataFormatada
    
    Set ListBox = ListBoxMateriaisCarrego
    
    ' Verifique se a ListBox não está vazia
    If ListBox.ListCount <= 1 Then
        MsgBox "A Lista está vazia.", vbExclamation
        Exit Sub
    End If
    
    'Seleciona a planilha
    PlanilhaPDFListaDespache.Select
    
    With PlanilhaPDFListaDespache
        
        .Range("A9:X1048564").ClearContents ' Apaga se tiver conteúdo na planilha
        
        'Nome motorista
        .Cells(4, 2).Value = ListBoxCarregos.list(linhaLista, 1)
        'Destino
        .Cells(2, 2).Value = ListBoxCarregos.list(linhaLista, 2)
        
        'Percorre a lista o cola os valores na planilha
        For i = 1 To ListBox.ListCount - 1
                
            'Cola os dados
            .Cells(linha, 1).Value = ListBox.list(i, 0)
            .Cells(linha, 2).Value = ListBox.list(i, 1)
            
            linha = linha + 1
        Next i
        
    End With
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
    "D:\Desktop\" & nomeArquivo & ".pdf", Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
    True
    
    ' Exiba uma mensagem de confirmação
    MsgBox "Dados exportados para PDF com sucesso!", vbInformation

    'Seleciona a planilha principal
    PlanilhaAuxiliar.Select
End Sub
'Botão exportar em pdf as chapas despachadas
Private Sub BtnExportarEstoqueChapasPDF_Click()
    
    'Variaveis do metodo
    Dim NomeDaPlanilha As String
    Dim custo As Double
    Dim qtdM2 As Double
    Dim comp As Double
    Dim alt As Double
    Dim custoPolimento As Double
    Dim expessura As String
    Dim idPedreira As String
    Dim nomeArquivo As String
    Dim ListBox As Object
    Dim linha As Integer
    Dim i As Long
    
    'Atribuições
    linha = 8
    nomeArquivo = txtNomeArquivoEstoqueChapas.Value
    Set ListBox = ListBoxEstoqueChapas
    
    ' Verifique se a ListBox não está vazia
    If ListBox.ListCount <= 1 Then
        MsgBox "A Lista está vazia.", vbExclamation
        Exit Sub
    End If
    
    'Verifica o nome do arquivo
    If nomeArquivo = "" Or nomeArquivo = "NOME DO ARQUIVO" Then
        
        'Deixa o cursor na ser adicionado o id
        txtNomeArquivoEstoqueChapas.SetFocus

        'Altera cor para melhor visualização
        txtNomeArquivoEstoqueChapas.BackColor = RGB(255, 182, 193)
        
        txtNomeArquivoEstoqueChapas.Value = ""

        'Mensagem de erro
        MsgBox "Adicione um nome para o arquivo!", vbCritical, "Nome não informado"

        'Para o fluxo do sistema para a correção
        Exit Sub
    End If
    
    'Volta a cor patrão
    txtNomeArquivoEstoqueChapas.BackColor = RGB(255, 255, 255)
    
    'Seleciona a planilha
    PlanilhaPDFEstoqueChapas.Select
    
    With PlanilhaPDFEstoqueChapas
        
        .Range("A9:O1048564").ClearContents ' Apaga se tiver conteúdo na planilha
        
        Range("A8").Select
        Selection.ClearContents
        Range("B8").Select
        Selection.ClearContents
        Range("C8").Select
        Selection.ClearContents
        Range("D8").Select
        Selection.ClearContents
        ActiveCell.FormulaR1C1 = "0"
        Range("E8").Select
        ActiveCell.FormulaR1C1 = "0"
        Range("F8").Select
        ActiveCell.FormulaR1C1 = "0"
        Range("G8").Select
        ActiveCell.FormulaR1C1 = "0"
        Range("H8").Select
        ActiveCell.FormulaR1C1 = "0"
        Range("I8").Select
        Selection.ClearContents
        Range("J8").Select
        Selection.ClearContents
        Range("K8").Select
        Selection.ClearContents
        Range("L8").Select
        Selection.ClearContents
        'Percorre a lista o cola os valores na planilha
        For i = 1 To ListBox.ListCount - 1
            
            custo = ListBox.list(i, 3)
            qtdM2 = ListBox.list(i, 4)
            comp = ListBox.list(i, 6)
            alt = ListBox.list(i, 7)
            'custoPolimento = ListBox.List(i, 12)
            expessura = ListBox.list(i, 8)
            idPedreira = ListBox.list(i, 11)
            
            'Cola os dados
            .Cells(linha, 1).Value = ListBox.list(i, 0)
            .Cells(linha, 2).Value = ListBox.list(i, 1)
            .Cells(linha, 3).Value = ListBox.list(i, 2)
            .Cells(linha, 4).Value = custo
            .Cells(linha, 5).Value = qtdM2
            .Cells(linha, 6).Value = ListBox.list(i, 5)
            .Cells(linha, 7).Value = comp
            .Cells(linha, 8).Value = alt
            .Cells(linha, 9).Value = expessura
            .Cells(linha, 10).Value = ListBox.list(i, 9)
            .Cells(linha, 11).Value = ListBox.list(i, 10)
            .Cells(linha, 12).Value = idPedreira
            '.Cells(linha, 13).Value = custoPolimento
            
            If linha <> 8 Then
               
               .Cells(linha, 13).Value = "=[@[CUSTO M²]]*[@[QTD CHAPAS]]"
            End If
        
            linha = linha + 1
        Next i
        
    End With

    'Exporta para PDF
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
    "D:\Desktop\" & nomeArquivo & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:= _
    True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    
    ' Exiba uma mensagem de confirmação
    MsgBox "Dados exportados para PDF com sucesso!", vbInformation
    
    'Seleciona a planilha principal
    PlanilhaAuxiliar.Select

End Sub
' Salva na planilha os dados e exporta em pdf estoque bloco
Public Sub exportarEstoqueBloco(listaBlocos As Collection, nomeArquivo As String)

    'Variaveis do metodo
    Dim bloco As objBloco
    Dim caminhoSalvar As String
    Dim linha As Integer
    Dim i As Long
    
    ' Atribuições
    linha = 8 ' Linha da tabela onde vai começar ser setado os dados
    caminhoSalvar = M_METODOS_GLOBAL.caminhoSalvarEstoqueBlocos & nomeArquivo & ".pdf" ' Caminho onde irá ser salvo pdf
    
    'Seleciona a planilha
    PlanPDFBlocos.Select
    
    With PlanPDFBlocos
        ' Cria bloco
        Set bloco = ObjectFactory.factoryBloco(bloco)
        ' Apaga se tiver conteúdo na planilha
        .Range("A8:X1048564").ClearContents
        'Percorre a lista o cola os valores na planilha
        For i = 1 To listaBlocos.Count
        
            Set bloco = listaBlocos.Item(i)

            'Cola os dados
            .Cells(linha, 1).Value = bloco.idSistema
            .Cells(linha, 2).Value = bloco.nomeMaterial
            .Cells(linha, 3).Value = bloco.tipoMaterial.nome
            .Cells(linha, 4).Value = bloco.custoMaterial
            .Cells(linha, 5).Value = bloco.qtdM3
            .Cells(linha, 6).Value = bloco.qtdChapas
            .Cells(linha, 7).Value = bloco.compLiquidoBloco
            .Cells(linha, 8).Value = bloco.altLiquidoBloco
            .Cells(linha, 9).Value = bloco.largLiquidoBloco
            .Cells(linha, 10).Value = bloco.valorTotalPolimento
            .Cells(linha, 11).Value = bloco.valorMetroSerrada
            .Cells(linha, 12).Value = bloco.valoresAdicionais
            .Cells(linha, 13).Value = bloco.valorTotalBloco
            .Cells(linha, 14).Value = bloco.freteBloco
            .Cells(linha, 15).Value = bloco.valorBloco
            .Cells(linha, 16).Value = bloco.dataCadastro
            .Cells(linha, 17).Value = bloco.estoque.nome
            .Cells(linha, 18).Value = bloco.numeroBlocoPedreira
            .Cells(linha, 19).Value = bloco.status.nome
            .Cells(linha, 20).Value = bloco.pedreira.nome
            .Cells(linha, 21).Value = bloco.serraria.nome
            .Cells(linha, 22).Value = bloco.valorMetroSerrada
            .Cells(linha, 23).Value = bloco.valorMetroPolimento
            .Cells(linha, 24).Value = bloco.observacao
            
            linha = linha + 1
            ' Libera espaço
            Set bloco = Nothing
        Next i
    End With
    
    ' Tira filtros
    Selection.AutoFilter
    Selection.AutoFilter
    ' Filtra só as linhas com conteudo
    Range("A8").Select
    ActiveSheet.ListObjects("ESTOQUE_BLOCOS").Range.AutoFilter Field:=1, _
    Criteria1:="<>"
    'Exporta para PDF
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
    caminhoSalvar, Quality:=xlQualityStandard, IncludeDocProperties:= _
    True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    ' Utilizando metodo para mensagem de sucesso
    Set msgSucesso = New clsErrorStyle
    msgSucesso.SemDadosError EXPORTADO_SUCESSO_MENSAGEM, EXPORTADO_SUCESSO_TITULO
    ' Tira filtros
    Selection.AutoFilter
    Selection.AutoFilter
    'Seleciona a planilha principal
    PlanAuxiliar.Select
End Sub
'Exportar em pdf as chapas despachadas
Private Sub exportarMateriaisDespachado()

    'Variaveis do metodo
    Dim custo As Double
    Dim qtdM2 As Double
    Dim comp As Double
    Dim alt As Double
    Dim frete As Double
    Dim dataDespache As Date
    Dim nomeArquivo As String
    Dim ListBox As Object
    Dim linha As Integer
    Dim i As Long

    'Atribuições
    linha = 8
    nomeArquivo = txtNomeArquivo.Value
    Set ListBox = ListBoxDespachado

    ' Verifique se a ListBox não está vazia
    If ListBox.ListCount <= 1 Then
        MsgBox "A Lista está vazia.", vbExclamation
        Exit Sub
    End If

    'Verifica o nome do arquivo
    If nomeArquivo = "" Or nomeArquivo = "NOME PARA ARQ PDF" Then

        'Deixa o cursor na ser adicionado o id
        txtNomeArquivo.SetFocus

        'Altera cor para melhor visualização
        txtNomeArquivo.BackColor = RGB(255, 182, 193)

        txtNomeArquivo.Value = ""

        'Mensagem de erro
        MsgBox "Adicione um nome para o arquivo!", vbCritical, "Nome não informado"

        'Para o fluxo do sistema para a correção
        Exit Sub
    End If

    'Volta a cor patrão
    txtNomeArquivo.BackColor = RGB(255, 255, 255)

    'Seleciona a planilha
    PlanilhaPDFChapasDespachadas.Select

    With PlanilhaPDFChapasDespachadas

        .Range("A8:P1048564").ClearContents ' Apaga se tiver conteúdo na planilha
        .Cells(4, 10).Value = cbDestino.Value ' Destino

        'Percorre a lista o cola os valores na planilha
        For i = 1 To ListBox.ListCount - 1

            custo = ListBox.list(i, 3)
            qtdM2 = ListBox.list(i, 4)
            comp = ListBox.list(i, 6)
            alt = ListBox.list(i, 7)
            frete = ListBox.list(i, 10)
            dataDespache = ListBox.list(i, 15)

            'Cola os dados
            .Cells(linha, 1).Value = ListBox.list(i, 0)
            .Cells(linha, 2).Value = ListBox.list(i, 1)
            .Cells(linha, 3).Value = ListBox.list(i, 2)
            .Cells(linha, 4).Value = custo
            .Cells(linha, 5).Value = qtdM2
            .Cells(linha, 6).Value = ListBox.list(i, 5)
            .Cells(linha, 7).Value = comp
            .Cells(linha, 8).Value = alt
            .Cells(linha, 9).Value = ListBox.list(i, 8)
            .Cells(linha, 10).Value = ListBox.list(i, 9)
            .Cells(linha, 11).Value = frete
            .Cells(linha, 12).Value = ListBox.list(i, 11)
            .Cells(linha, 13).Value = ListBox.list(i, 12)
            .Cells(linha, 14).Value = ListBox.list(i, 13)
            .Cells(linha, 15).Value = ListBox.list(i, 14)
            .Cells(linha, 16).Value = dataDespache

            linha = linha + 1
        Next i

    End With

    'Exporta para PDF
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
    "D:\Desktop\" & nomeArquivo & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:= _
    True, IgnorePrintAreas:=False, OpenAfterPublish:=True

    ' Exiba uma mensagem de confirmação
    MsgBox "Dados exportados para PDF com sucesso!", vbInformation

    'Seleciona a planilha principal
    PlanilhaAuxiliar.Select
End Sub

