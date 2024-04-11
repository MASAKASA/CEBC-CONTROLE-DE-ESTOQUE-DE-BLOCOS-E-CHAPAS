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
' Construtor da serraria no cadastro de bloco
Public Sub carregarSerraria(nomeserraria As String)
    nome = nomeserraria
End Sub

