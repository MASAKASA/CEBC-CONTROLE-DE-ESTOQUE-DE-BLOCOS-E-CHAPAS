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
Public tamanhos As New Collection
Public polideira As objPolideira
Public tipoPolimento As objTipoPolimento
Private tamanho As objTamanho
Private tipoMaterial As objTipoMaterial

' Construtor para cadastro
Public Sub carregarChapaCadastro(nomeMaterial As String, compBruto As String, _
                        altBruto As String, custo As String, total As String, _
                        polideira As objPolideira, tipoPolimento As objTipoPolimento)

    ' Atribuições
    Me.nomeMaterial = nomeMaterial
    Me.compBruto = compBruto
    Me.altBruto = altBruto
    Me.custo = custo
    Me.total = total
    Set Me.polideira = polideira
    Set Me.tipoPolimento = tipoPolimento
End Sub

' Construtor para manipulção
Public Sub carregarChapaManipulacao(idSistema As String, nomeMaterial As String, _
                        compBruto As String, altBruto As String, custo As String, _
                        total As String, polideira As objPolideira, tipoPolimento As objTipoPolimento)

    ' Atribuições
    Me.idSistema = idSistema
    Me.nomeMaterial = nomeMaterial
    Me.compBruto = compBruto
    Me.altBruto = altBruto
    Me.custo = custo
    Me.total = total
    Set Me.polideira = polideira
    Set Me.tipoPolimento = tipoPolimento
End Sub

' Adiciona os tamanhos das chapas
Public Sub addTamanhos(comp As String, alt As String, qtdM2 As String, qtdEstoque As String, esp As String, _
            tipoMaterial As objTipoMaterial)
    ' Atribuições
    Me.tamanho = New objTamanho
    Set Me.tipoMaterial = objTipoMaterial
    
    ' Criação do objeto
    tamanho.carregarTamanhoCadastro comp, alt, qtdM2, qtdEstoque, esp, tipoMaterial
    
    ' Adição do tamanho
    tamanhos.Add tamanho
End Sub
