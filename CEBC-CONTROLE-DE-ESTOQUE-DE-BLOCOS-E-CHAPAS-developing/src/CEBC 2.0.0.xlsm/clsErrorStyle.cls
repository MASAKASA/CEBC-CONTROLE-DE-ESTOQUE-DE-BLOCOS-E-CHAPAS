VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrorStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents efeitoTextBox As MSForms.TextBox
Attribute efeitoTextBox.VB_VarHelpID = -1
Public WithEvents efeitocomboBox As MSForms.comboBox
Attribute efeitocomboBox.VB_VarHelpID = -1
Public WithEvents efeitoOptionButton As MSForms.OptionButton
Attribute efeitoOptionButton.VB_VarHelpID = -1

' Pesquisa sem dados
Public Sub SemDadosError(mensagemError As String, tituloError As String)
    'Mensagem de erro
    MsgBox mensagemError, vbInformation, tituloError
End Sub

' Formata com estilo de error
Public Sub EntrarErrorStyleTextBox(controle As MSForms.TextBox, mensagemError As String, tituloError As String)
    ' Deixa o focus no controle com erro
    controle.SetFocus
    'Altera cor para melhor visualiza��o
    controle.BackColor = RGB(250, 180, 240)
    
    'Mensagem de erro
    MsgBox mensagemError, vbCritical, tituloError
End Sub

' Volta a formata��o principal
Public Sub sairErrorStyleTextBox(controle As MSForms.TextBox)
    'Deixa na cor patr�o
    controle.BackColor = RGB(200, 230, 255)
End Sub

' Formata com estilo de error
Public Sub EntrarErrorStyleComboBox(controle As MSForms.comboBox, mensagemError As String, tituloError As String)
    ' Deixa o focus no controle com erro
    controle.SetFocus
    'Altera cor para melhor visualiza��o
    controle.BackColor = RGB(250, 180, 240)
    
    'Mensagem de erro
    MsgBox mensagemError, vbCritical, tituloError
End Sub

' Volta a formata��o principal
Public Sub sairErrorStyleComboBox(controle As MSForms.comboBox)
    'Deixa na cor patr�o
    controle.BackColor = RGB(200, 230, 255)
End Sub

' Formata com estilo de error
Public Sub EntrarErrorStyleOptionButton(controle As MSForms.OptionButton, mensagemError As String, tituloError As String)
    ' Deixa o focus no controle com erro
    controle.SetFocus
    'Altera cor para melhor visualiza��o
    controle.ForeColor = RGB(250, 180, 240)
    
    'Mensagem de erro
    MsgBox mensagemError, vbCritical, tituloError
End Sub

' Volta a formata��o principal
Public Sub sairErrorStyleOptionButton(controle As MSForms.OptionButton)
    'Deixa na cor patr�o
    controle.ForeColor = RGB(0, 32, 96)
End Sub
