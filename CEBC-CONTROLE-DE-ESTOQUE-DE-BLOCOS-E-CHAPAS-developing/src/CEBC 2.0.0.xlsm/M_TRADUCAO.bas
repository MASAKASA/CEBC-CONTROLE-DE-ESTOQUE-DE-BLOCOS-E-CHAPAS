Attribute VB_Name = "M_TRADUCAO"
Option Explicit

' Variaveis globais para os erros
Global EXPORTADO_SUCESSO_MENSAGEM As String
Global EXPORTADO_SUCESSO_TITULO As String
Global LIST_SEM_DADOS_MENSAGEM As String
Global LIST_SEM_DADOS_TITULO As String
Global ARQUIVO_SEM_NOME_MENSAGEM As String
Global ARQUIVO_SEM_NOME_TITULO As String
Global SEM_DADOS_MENSAGEM As String
Global SEM_DADOS_TITULO As String
Global ADICIONE_DATA_MENSAGEM As String
Global ADICIONE_DATA_TITULO As String
Global ADICIONE_STATUS_MENSAGEM As String
Global ADICIONE_STATUS_TITULO As String
' Carrega as variaveis dos erros
Public Sub carregarTraducaoErros()
    ' Mensagem para usuário
    EXPORTADO_SUCESSO_MENSAGEM = "Dados exportados para PDF com sucesso!"
    EXPORTADO_SUCESSO_TITULO = "Sucesso na esportação"
    LIST_SEM_DADOS_MENSAGEM = "Faça primeiro uma pesquisa para poder exportar!"
    LIST_SEM_DADOS_TITULO = "Lista sem dados"
    ARQUIVO_SEM_NOME_MENSAGEM = "Digite um nome para o arquivo!"
    ARQUIVO_SEM_NOME_TITULO = "Arquivo sem nome"
    SEM_DADOS_MENSAGEM = "Nada encontrado!"
    SEM_DADOS_TITULO = "Sem dados"
    ADICIONE_DATA_MENSAGEM = "Adicione uma data válida!"
    ADICIONE_DATA_TITULO = "Data inválida"
    ADICIONE_STATUS_MENSAGEM = "Selecione um Status para pesquisa!"
    ADICIONE_STATUS_TITULO = "Informe um Status"
End Sub
