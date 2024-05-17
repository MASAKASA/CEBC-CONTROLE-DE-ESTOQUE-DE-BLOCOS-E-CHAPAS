VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objChapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public idSistema As String
Public nomeMaterial As String
Public qtdEstoque As String
Public idPedreira As String
Public custoPolimento As String
Public custoTotal As String
Public compBruto As String
Public altBruto As String
Public qtdM2Bruto As String
Public bloco As objBloco
Public polideira As objPolideira
Public tipoPolimento As objTipoPolimento
Public estoque As objEstoqueChapa
Public tamanhos As Collection

' Construtor para cadastro
Public Sub carregarChapaCadastro(idSistema As String, nomeMaterial As String, qtdEstoque As String, idPedreira As String, _
                    custoPolimento As String, custoTotal As String, compBruto As String, altBruto As String, qtdM2Bruto As String, _
                    bloco As objBloco, polideira As objPolideira, tipoPolimento As objTipoPolimento, estoque As objEstoqueChapa, _
                    tamanhos As Collection)

    ' Atribuições
    Me.idSistema = idSistema
    Me.nomeMaterial = nomeMaterial
    Me.qtdEstoque = qtdEstoque
    Me.idPedreira = idPedreira
    Me.custoPolimento = custoPolimento
    Me.custoTotal = custoTotal
    Me.compBruto = compBruto
    Me.altBruto = altBruto
    Me.qtdM2Bruto = qtdM2Bruto
    Set Me.bloco = bloco
    Set Me.polideira = polideira
    Set Me.tipoPolimento = tipoPolimento
    Set Me.estoque = estoque
    Set Me.tamanhos = tamanhos
End Sub

' Construtor para manipulção
Public Sub carregarChapaManipulacao(idSistema As String, nomeMaterial As String, qtdEstoque As String, idPedreira As String, _
                    custoPolimento As String, custoTotal As String, compBruto As String, altBruto As String, qtdM2Bruto As String, _
                    bloco As objBloco, polideira As objPolideira, tipoPolimento As objTipoPolimento, estoque As objEstoqueChapa, _
                    tamanhos As Collection)

    ' Atribuições
    Me.idSistema = idSistema
    Me.nomeMaterial = nomeMaterial
    Me.qtdEstoque = qtdEstoque
    Me.idPedreira = idPedreira
    Me.custoPolimento = custoPolimento
    Me.custoTotal = custoTotal
    Me.compBruto = compBruto
    Me.altBruto = altBruto
    Me.qtdM2Bruto = qtdM2Bruto
    Set Me.bloco = bloco
    Set Me.polideira = polideira
    Set Me.tipoPolimento = tipoPolimento
    Set Me.estoque = estoque
    Set Me.tamanhos = tamanhos
End Sub

' Gets e sets
Public Function getBloco() As objBloco
    getBloco = bloco
End Function

Public Function setBloco(bloco As objBloco)
    Set Me.bloco = bloco
End Function

Public Function getPolideira() As objPolideira
    getPolideira = polideira
End Function

Public Function setPolideira(polideira As objPolideira)
    Set Me.polideira = polideira
End Function

Public Function getTipoPolimento() As objTipoPolimento
    getTipoPolimento = tipoPolimento
End Function

Public Function setTipoPolimento(tipoPolimento As objTipoPolimento)
    Set Me.tipoPolimento = tipoPolimento
End Function

Public Function getEstoque() As objEstoqueChapa
    getEstoque = estoque
End Function

Public Function setEstoque(estoque As objEstoqueChapa)
    Set Me.estoque = estoque
End Function

Public Function getTamanhos() As Collection
    getTamanhos = tamanhos
End Function

Public Function setTamanhos(tamanhos As Collection)
    Set Me.tamanhos = tamanhos
End Function
