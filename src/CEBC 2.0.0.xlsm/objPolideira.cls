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

Public id As String
Public nome As String

' Construtor cadastro
Public Sub carregarPolideiraCadastro(nomePolideira As String)
    Me.nome = nomePedreira
End Sub

' Construtor para manipulação
Public Sub carregarPolideiraManipulacao(id As String, nomePolideira As String)
    Me.id = id
    Me.nome = nomePolideira
End Sub
