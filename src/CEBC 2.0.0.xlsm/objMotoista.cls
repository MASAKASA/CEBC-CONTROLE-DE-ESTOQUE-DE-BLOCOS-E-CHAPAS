VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objMotoista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As String
Public nome As String

' Construtor para cadastro
Public Sub carregarMotoristaCadastro(nome As String)
    Me.nome = nome
End Sub

' Construtor para manipulação
Public Sub carregarMotoristaManipulacao(id As String, nome As String)
    Me.id = id
    Me.nome = nome
End Sub
