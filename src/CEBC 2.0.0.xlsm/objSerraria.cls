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

Public id As String
Public nome As String

' Construtor para cadastro
Public Sub carregarSerrariaCadastro(nomeserraria As String)
    Me.nome = nomeserraria
End Sub

' Construtor para manipulação
Public Sub carregarSerrariaManipulacao(id As String, nomeserraria As String)
    Me.id = id
    Me.nome = nomeserraria
End Sub
