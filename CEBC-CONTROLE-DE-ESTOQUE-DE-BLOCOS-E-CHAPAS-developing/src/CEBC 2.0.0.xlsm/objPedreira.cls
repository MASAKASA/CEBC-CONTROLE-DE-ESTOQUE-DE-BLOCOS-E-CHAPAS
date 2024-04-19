VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objPedreira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As String
Public nome As String

' Construtor para cadastro
Public Sub carregarPedreiraCadastro(nomePedreira As String)
    Me.nome = nomePedreira
End Sub

' Construtor para manipulação
Public Sub carregarPedreiraManipulacao(id As String, nomePedreira As String)
    Me.id = id
    Me.nome = nomePedreira
End Sub
