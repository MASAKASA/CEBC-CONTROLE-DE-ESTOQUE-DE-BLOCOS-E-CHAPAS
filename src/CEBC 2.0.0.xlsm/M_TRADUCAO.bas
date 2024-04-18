Attribute VB_Name = "M_TRADUCAO"
Option Explicit

' Variaveis globais para os erros
Global SEM_DADOS_MENSAGEM As String
Global SEM_DADOS_TITULO As String
Global ADICIONE_DATA_MENSAGEM As String
Global ADICIONE_DATA_TITULO As String
Global ADICIONE_STATUS_MENSAGEM As String
Global ADICIONE_STATUS_TITULO As String
' Carrega as variaveis dos erros
Public Sub carregarTraducaoErros()
    ' Mensagem para usuário
    SEM_DADOS_MENSAGEM = "Nada encontrado!"
    SEM_DADOS_TITULO = "Sem dados"
    ADICIONE_DATA_MENSAGEM = "Adicione uma data válida!"
    ADICIONE_DATA_TITULO = "Data inválida"
    ADICIONE_STATUS_MENSAGEM = "Selecione um Status para pesquisa!"
    ADICIONE_STATUS_TITULO = "Informe um Status"
End Sub
