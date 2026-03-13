VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objPolideira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Public nome As String
Public ativo As String

' Construtor cadastro
Public Sub carregarPolideiraCadastro(nomePolideira As String, ativo As String)
    Me.nome = nomePedreira
    Me.ativo = ativo
End Sub

' Construtor para manipulação
Public Sub carregarPolideiraManipulacao(id As Long, nomePolideira As String, ativo As String)
    Me.id = id
    Me.nome = nomePolideira
    Me.ativo = ativo
End Sub
