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
Public compBruto As String
Public altBruto As String
Public custo As String
Public total As String
'Public tamanhos() As objTamanho
Public polideira As objPolideira
Public tipoPolimento As objTipoPolimento

' Construtor para cadastro
Public Sub carregarChapaCadastro(nomeMaterial As String, compBruto As String, _
                        altBruto As String, custo As String, total As String, _
                        tamanhos() As objTamanho, polideira As objPolideira, _
                        tipoPolimento As objTipoPolimento)

    ' Atribuições
    Me.nomeMaterial = nomeMaterial
    Me.compBruto = compBruto
    Me.altBruto = altBruto
    Me.custo = custo
    Me.total = total
    Set Me.tamanhos() = tamanhos
    Set Me.polideira = polideira
    Set Me.tipoPolimento = tipoPolimento
End Sub

' Construtor para manipulção
Public Sub carregarChapaManipulacao(idSistema As String, nomeMaterial As String, _
                        compBruto As String, altBruto As String, custo As String, _
                        total As String, tamanhos() As objTamanho, polideira As objPolideira, _
                        tipoPolimento As objTipoPolimento)

    ' Atribuições
    Me.idSistema = idSistema
    Me.nomeMaterial = nomeMaterial
    Me.compBruto = compBruto
    Me.altBruto = altBruto
    Me.custo = custo
    Me.total = total
    Set Me.tamanhos() = tamanhos
    Set Me.polideira = polideira
    Set Me.tipoPolimento = tipoPolimento
End Sub

