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
Public valorTotal As String
Public numeroBlocoPedreira As String
Public estoqueZero As String
Public tipoPolimento As objTipoPolimento
Public bloco As objBloco
Public tamanhos As Collection

' Construtor
Public Sub carregarChapa(idSistema As String, nomeMaterial As String, valorTotal As String, numeroBlocoPedreira As String, _
                    tipoPolimento As objTipoPolimento, bloco As objBloco, tamanhos As Collection)

    ' Atribuições
    Me.idSistema = idSistema
    Me.nomeMaterial = nomeMaterial
    Me.valorTotal = valorTotal
    Me.numeroBlocoPedreira = numeroBlocoPedreira
    Set Me.tipoPolimento = tipoPolimento
    Set Me.bloco = bloco
    Set Me.tamanhos = tamanhos
End Sub

' Gets e sets
Public Function getTipoPolimento() As objTipoPolimento
    getTipoPolimento = tipoPolimento
End Function

Public Function setTipoPolimento(tipoPolimento As objTipoPolimento)
    Set Me.tipoPolimento = tipoPolimento
End Function

Public Function getBloco() As objBloco
    getBloco = bloco
End Function

Public Function setBloco(bloco As objBloco)
    Set Me.bloco = bloco
End Function

Public Function getTamanhos() As Collection
    Set getTamanhos = tamanhos
End Function

Public Function setTamanhos(tamanhos As Collection)
    Set Me.tamanhos = tamanhos
End Function
