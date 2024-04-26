Attribute VB_Name = "M_TRADUCAO"
Option Explicit

' Variaveis globais para os erros
' TELA CADASTRO DE BLOCO
Global SUCESSO_EDICAO_MENSAGEM As String
Global SUCESSO_EDICAO_TITULO As String
Global ERRO_DESCONHECIDO_MENSAGEM As String
Global ERRO_DESCONHECIDO_TITULO As String
Global CADASTRO_CONFIRMADO_MENSAGEM As String
Global CADASTRO_CONFIRMADO_TITULO As String
Global ACAO_CANCELADA_MENSAGEM As String
Global ACAO_CANCELADA_TITULO As String
Global CONFIRMACAO_CADASTRO_MENSAGEM As String
Global CONFIRMACAO_CADASTRO_TITULO As String
Global NOME_BLOCO_MENSAGEM As String
Global NOME_BLOCO_PEDREIRA_TITULO As String
Global NUMERO_BLOCO_PEDREIRA_MENSAGEM As String
Global NUMERO_BLOCO_PEDREIRA_TITULO As String
Global NOME_PEDREIRA_MENSAGEM As String
Global NOME_PEDREIRA_TITULO As String
Global STATUS_SERRARIA_MENSAGEM As String
Global STATUS_SERRARIA_TITULO As String

' TELA ESTOQUE BLOCOS M³
Global SELECIONE_TEM_MENSAGEM As String
Global SELECIONE_TEM_TITULO As String
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
    ' TELA CADASTRO DE BLOCO
    ERRO_DESCONHECIDO_MENSAGEM = "Erro desconhecido! Procurar o suporte!"
    ERRO_DESCONHECIDO_TITULO = "Erro desconhecido"
    SUCESSO_EDICAO_MENSAGEM = "Edição feito com sucesso!"
    SUCESSO_EDICAO_TITULO = "Edição"
    CADASTRO_CONFIRMADO_MENSAGEM = "Cadastrado com sucesso!"
    CADASTRO_CONFIRMADO_TITULO = "Cadastro confirmado"
    ACAO_CANCELADA_MENSAGEM = "A ação foi cancelada!"
    ACAO_CANCELADA_TITULO = "Cancelamento"
    CONFIRMACAO_CADASTRO_MENSAGEM = "Confira se o número e descrição/material do bloco estão corretos, pois a junção deles irá criar o ID do bloco no sistema. ID do bloco não poderá ser alterado posteriormente. Tudo conferido e podemos seguir com o cadastro?"
    CONFIRMACAO_CADASTRO_TITULO = "Atenção - Confirmação"
    NOME_BLOCO_MENSAGEM = "Adicone a descrição do bloco!"
    NOME_BLOCO_PEDREIRA_TITULO = "Descrição do bloco não informada"
    NUMERO_BLOCO_PEDREIRA_MENSAGEM = "Adicone o número do bloco!"
    NUMERO_BLOCO_PEDREIRA_TITULO = "ID do bloco não informado"
    STATUS_SERRARIA_MENSAGEM = "Se o bloco já estiver na serraria, selecione o nome da 'Serraria'!, Caso não esteja, selecione a opção: 'BLOCO ESTÁ NA? PEDREIRA"
    STATUS_SERRARIA_TITULO = "Status do bloco"
    NOME_PEDREIRA_MENSAGEM = "Selecione o nome da pedreira!"
    NOME_PEDREIRA_TITULO = "Nome da pedreira não informada"
    
    ' TELA ESTOQUE BLOCOS M³
    SELECIONE_TEM_MENSAGEM = "Selecione um item da lista!"
    SELECIONE_TEM_TITULO = "Nada selecioando"
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
