VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objTamanho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As String
Public compremento As String
Public altura As String
Public qtdM2 As String
Public qtdEstoque As String
Public valorPolimento As String
Public espessura As String
Public chapa As objChapa
Public tipoMaterial As objTipoMaterial
Public polideira As objPolideira
Public estoque As objEstoqueChapa

' Construtor para cadastro
Public Sub carregarTamanhoCadastro(compremento As String, altura As String, qtdM2 As String, _
                            qtdEstoque As String, valorPolimento As String, espessura As String, _
                            tipoMaterial As objTipoMaterial, chapa As objChapa, polideira As objPolideira, _
                            estoque As objEstoqueChapa)
    
    ' Atribuições
    Me.compremento = compremento
    Me.altura = altura
    Me.qtdM2 = qtdM2
    Me.qtdEstoque = qtdEstoque
    Me.valorPolimento = valorPolimento
    Me.espessura = espessura
    Set Me.chapa = chapa
    Set Me.tipoMaterial = tipoMaterial
    Set Me.polideira = polideira
    Set Me.estoque = estoque
End Sub

' Construtor para manupulação
Public Sub carregarTamanhoManipulacao(id As String, compremento As String, altura As String, qtdM2 As String, _
                            qtdEstoque As String, valorPolimento As String, espessura As String, _
                            tipoMaterial As objTipoMaterial, chapa As objChapa, polideira As objPolideira, _
                            estoque As objEstoqueChapa)
    
    ' Atribuições
    Me.id = id
    Me.compremento = compremento
    Me.altura = altura
    Me.qtdM2 = qtdM2
    Me.qtdEstoque = qtdEstoque
    Me.valorPolimento = valorPolimento
    Me.espessura = espessura
    Set Me.chapa = chapa
    Set Me.tipoMaterial = tipoMaterial
    Set Me.polideira = polideira
    Set Me.estoque = estoque
End Sub

' Gets e sets
Public Function getTipoMaterial() As objTipoMaterial
    getTipoMaterial = Me.tipoMaterial
End Function

Public Function setTipoMaterial(tipoMaterial As objTipoMaterial)
    Set Me.tipoMaterial = tipoMaterial
End Function

Public Function getChapa() As objChapa
    getChapa = Me.chapa
End Function

Public Function setChapa(chapa As objChapa)
    Set Me.chapa = chapa
End Function

Public Function getPolideira() As objPolideira
    getPolideira = Me.polideira
End Function

Public Function setPolideira(polideira As objPolideira)
    Set Me.polideira = polideira
End Function

Public Function getEstoque() As objEstoqueChapa
    getEstoque = Me.estoque
End Function

Public Function setEstoque(estoque As objEstoqueChapa)
    Set Me.estoque = estoque
End Function
