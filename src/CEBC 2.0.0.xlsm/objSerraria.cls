VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objSerraria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Public nome As String
Public ativo As String

' Construtor para cadastro
Public Sub carregarSerrariaCadastro(nomeSerraria As String, ativo As String)
    Me.nome = nomeSerraria
    Me.ativo = ativo
End Sub

' Construtor para manipulação
Public Sub carregarSerrariaManipulacao(id As Long, nomeSerraria As String, ativo As String)
    Me.id = id
    Me.nome = nomeSerraria
    Me.ativo = ativo
End Sub
