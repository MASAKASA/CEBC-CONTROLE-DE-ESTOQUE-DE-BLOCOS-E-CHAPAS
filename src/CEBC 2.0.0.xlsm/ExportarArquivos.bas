Attribute VB_Name = "ExportarArquivos"
Option Explicit

Private msgSucesso As clsErrorStyle

' Salva na planilha os dados e exporta em pdf as chapas despachadas por motoristas
Public Sub exportarCarregoPDF(despache As objDespache)

    'Variaveis do metodo
    Dim chapaDespache As objChapa
    Dim tamanhoChapaDespache As objTamanho
    Dim caminhoSalvar As String
    Dim dataFormatada As String
    Dim nomeArquivo As String
    Dim comp As Double
    Dim alt As Double
    Dim linha As Integer
    Dim i As Long
    
    'Atribuições
    linha = 9
    
    dataFormatada = despache.dataDespache
    dataFormatada = M_METODOS_GLOBAL.ConverterFormatoData(dataFormatada)
    
    nomeArquivo = "Carregamento " & despache.getMotorista.nome & " " & dataFormatada
    
    caminhoSalvar = M_METODOS_GLOBAL.caminhoSalvarCarregoChapas & nomeArquivo & ".pdf" ' Caminho onde irá ser salvo pdf
    
    'Seleciona a planilha
    PlanPDFListaDespache.Select
    
    With PlanPDFListaDespache
        
        .Range("A9:X1048564").ClearContents ' Apaga se tiver conteúdo na planilha
        
        ' Destino
        .Cells(2, 3).Value = despache.getDestino.nome
        
        ' Nome motorista
        .Cells(4, 3).Value = despache.getMotorista.nome
        
        ' Total chapas
        .Cells(6, 3).Value = despache.qtdChapas
        
        ' Percorre a lista o cola os valores na planilha
        For i = 1 To despache.getListaChapas.Count
                
            Set chapaDespache = despache.getListaChapas.Item(i)
            Set tamanhoChapaDespache = chapaDespache.getTamanhos.Item(1)
            
            comp = CDbl(tamanhoChapaDespache.compremento)
            alt = CDbl(tamanhoChapaDespache.altura)
            
            ' Cola os dados
            .Cells(linha, 1).Value = chapaDespache.nomeMaterial
            .Cells(linha, 2).Value = tamanhoChapaDespache.qtdEstoque
            .Cells(linha, 3).Value = comp
            .Cells(linha, 4).Value = alt
            .Cells(linha, 5).Value = chapaDespache.numeroBlocoPedreira
            
            linha = linha + 1
            
            Set chapaDespache = Nothing
            Set tamanhoChapaDespache = Nothing
        Next i
        
    End With
    
    ' Tira filtros
    Selection.AutoFilter
    Selection.AutoFilter
    
    ' Filtra só as linhas com conteudo
    Range("A8").Select
    ActiveSheet.ListObjects("LISTA_CARGA").Range.AutoFilter Field:=1, _
    Criteria1:="<>"
    
    ' Exporta para PDF
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
    caminhoSalvar, Quality:=xlQualityStandard, IncludeDocProperties:= _
    True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    
    ' Utilizando metodo para mensagem de sucesso
    Set msgSucesso = New clsErrorStyle
    msgSucesso.Informativo EXPORTADO_SUCESSO_MENSAGEM, EXPORTADO_SUCESSO_TITULO
    
    ' Tira filtros
    Selection.AutoFilter
    Selection.AutoFilter
    
    ' Seleciona a planilha principal
    PlanInicio.Select
End Sub

' Salva na planilha os dados e exporta em pdf estoque chapa
Public Sub exportarEstoqueChapa(listaChapas As Collection, nomeArquivo As String)
    
    'Variaveis do metodo
    Dim chapa As objChapa
    Dim tamanho As objTamanho
    Dim caminhoSalvar As String
    Dim totalM2 As Double
    Dim totalEstoque As Integer
    Dim linha As Integer
    Dim i As Integer
    Dim j As Integer
    
    'Atribuições
    linha = 8  ' Linha da tabela onde vai começar ser setado os dados
    caminhoSalvar = M_METODOS_GLOBAL.caminhoSalvarEstoqueChapas & nomeArquivo & ".pdf" ' Caminho onde irá ser salvo pdf
    
    'Seleciona a planilha
    PlanPDFEstoqueChapas.Select
    
    With PlanPDFEstoqueChapas
        ' Cria chapa
        'Set chapa = ObjectFactory.factoryBloco(chapa)
        ' Apaga se tiver conteúdo na planilha
        .Range("A8:M1048564").ClearContents
        'Percorre a lista o cola os valores na planilha
        For i = 1 To listaChapas.Count
            ' Seta chapa
            Set chapa = listaChapas.Item(i)
            
            ' Soma m² e estoque total
            For j = 1 To chapa.tamanhos.Count
                ' Seta tamanho
                Set tamanho = chapa.tamanhos.Item(j)
                
                ' Soma
                totalM2 = totalM2 + CDbl(tamanho.qtdM2)
                totalEstoque = totalEstoque + CInt(tamanho.qtdEstoque)
                ' Libera espaço memoria
                Set tamanho = Nothing
            Next j
            
            'Cola os dados
            .Cells(linha, 1).Value = chapa.idSistema
            .Cells(linha, 2).Value = chapa.nomeMaterial
            .Cells(linha, 3).Value = chapa.tipoPolimento.nome
            .Cells(linha, 4).Value = totalM2
            .Cells(linha, 5).Value = totalEstoque
            .Cells(linha, 6).Value = chapa.numeroBlocoPedreira
            .Cells(linha, 7).Value = chapa.valorTotal
            
            linha = linha + 1
        Next i
        ' Libera espaço memoria
        Set chapa = Nothing
    End With

    ' Filtra só as linhas com conteudo
    Range("A8").Select
    ' Tira filtros
    Selection.AutoFilter
    Selection.AutoFilter
    ActiveSheet.ListObjects("ESTOQUE_CHAPAS").Range.AutoFilter Field:=1, _
    Criteria1:="<>"
    
    ' Exporta para PDF
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
    caminhoSalvar, Quality:=xlQualityStandard, IncludeDocProperties:= _
    True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    
    ' Utilizando metodo para mensagem de sucesso
    Set msgSucesso = New clsErrorStyle
    msgSucesso.Informativo EXPORTADO_SUCESSO_MENSAGEM, EXPORTADO_SUCESSO_TITULO
    
    ' Tira filtros
    Selection.AutoFilter
    Selection.AutoFilter
    
    ' Seleciona a planilha principal
    PlanInicio.Select
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
    
    ' Exporta para PDF
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
    caminhoSalvar, Quality:=xlQualityStandard, IncludeDocProperties:= _
    True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    
    ' Utilizando metodo para mensagem de sucesso
    Set msgSucesso = New clsErrorStyle
    msgSucesso.Informativo EXPORTADO_SUCESSO_MENSAGEM, EXPORTADO_SUCESSO_TITULO
    
    ' Tira filtros
    Selection.AutoFilter
    Selection.AutoFilter
    
    ' Seleciona a planilha principal
    PlanInicio.Select
End Sub
