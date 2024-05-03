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
Public idChapa As String
Public compremento As String
Public altura As String
Public qtdM2 As String
Public qtdEstoque As String
Public espessura As String
Public tipoMaterial As objTipoMaterial

' Construtor para cadastro
Public Sub carregarTamanhoCadastro(compremento As String, altura As String, qtdM2 As String, _
                            qtdEstoque As String, espessura As String, tipoMaterial As objTipoMaterial)
    
    ' Atribuições
    Me.compremento = compremento
    Me.altura = altura
    Me.qtdM2 = qtdM2
    Me.qtdEstoque = qtdEstoque
    Me.espessura = espessura
    Set Me.tipoMaterial = tipoMaterial
End Sub

' Construtor para manupulação
Public Sub carregarTamanhoManipulacao(id As String, compremento As String, altura As String, qtdM2 As String, _
                            qtdEstoque As String, espessura As String, tipoMaterial As objTipoMaterial)
    
    ' Atribuições
    Me.id = id
    Me.compremento = compremento
    Me.altura = altura
    Me.qtdM2 = qtdM2
    Me.qtdEstoque = qtdEstoque
    Me.espessura = espessura
    Set Me.tipoMaterial = tipoMaterial
End Sub

' Gets e sets
Public Function getTipoMaterial() As objTipoMaterial
    getTipoMaterial = tipoMaterial
End Function

Public Function setTipoMaterial(tipoMaterial As objTipoMaterial)
    Set Me.tipoMaterial = tipoMaterial
End Function
