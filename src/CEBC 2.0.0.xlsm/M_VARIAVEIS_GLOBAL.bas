Attribute VB_Name = "M_VARIAVEIS_GLOBAL"
Option Explicit

' Variavel para controlar inicio e fim do relogio
Global INICIAR_RELOGIO As Boolean

' Variaveis para manipula��o de banco de dados
Global CONEXAO_BD As New ADODB.Connection

' Datas em texto patr�o para manipula��o
Global DATA_INICIO As String
Global DATA_FINAL As String

