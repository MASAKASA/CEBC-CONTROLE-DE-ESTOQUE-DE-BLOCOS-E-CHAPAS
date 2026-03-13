VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objTipoPolimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Public nome As String
Public ativo As String

' Construtor para cadastro
Public Sub carregarTipoPolimentoCadastro(nome As String, ativo As String)
    Me.nome = nome
    Me.ativo = ativo
End Sub

' Construtor para manipulação
Public Sub carregarTipoPolimentoManipulacao(id As Long, nome As String, ativo As String)
    Me.id = id
    Me.nome = nome
    Me.ativo = ativo
End Sub
