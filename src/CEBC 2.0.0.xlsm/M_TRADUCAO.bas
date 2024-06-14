Attribute VB_Name = "M_TRADUCAO"
Option Explicit

' Variaveis globais para os erros
' TROCA ESTOQUE
Global ADICIONE_CHAPA_MENSAGEM As String
Global ADICIONE_CHAPA_TITULO As String

' TELA TROCA ESTOQUE
Global VALOR_SUPERIOR_MENSAGEM As String
Global VALOR_SUPERIOR_TITULO As String

' TELA CADASTRO AVULSO
Global AVULSO_JA_CADASTRADO_MENSAGEM As String
Global AVULSO_JA_CADASTRADO_TITULO As String
Global NOME_AVULSO_MENSAGEM As String
Global NOME_AVULSO_TITULO As String

' TELA EDI��O DE BLOCO
Global HABILITE_EDICAO_MENSAGEM As String
Global HABILITE_EDICAO_TITULO As String

' TELA CADASTRO DE BLOCO
Global BLOCO_JA_CADASTRADO_MENSAGEM As String
Global BLOCO_JA_CADASTRADO_TITULO As String
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

' TELA ESTOQUE BLOCOS M�
Global ESCOLHA_CHAPA_MENSAGEM As String
Global ESCOLHA_CHAPA_TITULO As String
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
    ' Mensagem para usu�rio
    ' TROCA ESTOQUE
    ADICIONE_CHAPA_MENSAGEM = "Adicione chapa para troca!"
    ADICIONE_CHAPA_TITULO = "Sem chapas"

    ' TELA TROCA ESTOQUE
    VALOR_SUPERIOR_MENSAGEM = "A quantidade para lan�amento � maior do que a disp�nivel em estoque!"
    VALOR_SUPERIOR_TITULO = "Sem estoque"""
    
    ' TELA CADASTRO AVULSO
    AVULSO_JA_CADASTRADO_MENSAGEM = "Avulso j� cadastrado no sistema!"
    AVULSO_JA_CADASTRADO_TITULO = "Cadastro duplicado"
    NOME_AVULSO_MENSAGEM = "Adicione uma descri��o!"
    NOME_AVULSO_TITULO = "Descri��o n�o informada"

    ' TELA EDI��O DE BLOCO
    HABILITE_EDICAO_MENSAGEM = "Habilite edi��o!"
    HABILITE_EDICAO_TITULO = "Edi��o desabilitada"
    
    ' TELA CADASTRO DE BLOCO
    ERRO_DESCONHECIDO_MENSAGEM = "Erro desconhecido! Procurar o suporte!"
    ERRO_DESCONHECIDO_TITULO = "Erro desconhecido"
    BLOCO_JA_CADASTRADO_MENSAGEM = "Bloco j� cadastrado no sistema!"
    BLOCO_JA_CADASTRADO_TITULO = "Cadastrado duplicado"
    CADASTRO_CONFIRMADO_MENSAGEM = "Cadastrado com sucesso!"
    CADASTRO_CONFIRMADO_TITULO = "Cadastro confirmado"
    ACAO_CANCELADA_MENSAGEM = "A a��o foi cancelada!"
    ACAO_CANCELADA_TITULO = "Cancelamento"
    CONFIRMACAO_CADASTRO_MENSAGEM = "Confira se o n�mero e descri��o/material est�o corretos, pois a jun��o deles ir� criar o ID do bloco no sistema. ID n�o poder� ser alterado posteriormente. Tudo conferido e podemos seguir com o cadastro?"
    CONFIRMACAO_CADASTRO_TITULO = "Aten��o - Confirma��o"
    NOME_BLOCO_MENSAGEM = "Adicone a descri��o do bloco!"
    NOME_BLOCO_PEDREIRA_TITULO = "Descri��o do bloco n�o informada"
    NUMERO_BLOCO_PEDREIRA_MENSAGEM = "Adicone o n�mero do bloco!"
    NUMERO_BLOCO_PEDREIRA_TITULO = "ID do bloco n�o informado"
    STATUS_SERRARIA_MENSAGEM = "Se o bloco j� estiver na serraria, selecione o nome da 'Serraria'!, Caso n�o esteja, selecione a op��o: 'BLOCO EST� NA? PEDREIRA"
    STATUS_SERRARIA_TITULO = "Status do bloco"
    NOME_PEDREIRA_MENSAGEM = "Selecione o nome da pedreira!"
    NOME_PEDREIRA_TITULO = "Nome da pedreira n�o informada"
    
    ' TELA ESTOQUE BLOCOS M�
    ESCOLHA_CHAPA_MENSAGEM = "Selecione uma chapa na lista!"
    ESCOLHA_CHAPA_TITULO = "Escolha chapa"
    SELECIONE_TEM_MENSAGEM = "Selecione um item da lista!"
    SELECIONE_TEM_TITULO = "Nada selecioando"
    EXPORTADO_SUCESSO_MENSAGEM = "Dados exportados para PDF com sucesso!"
    EXPORTADO_SUCESSO_TITULO = "Sucesso na esporta��o"
    LIST_SEM_DADOS_MENSAGEM = "Fa�a primeiro uma pesquisa para poder exportar!"
    LIST_SEM_DADOS_TITULO = "Lista sem dados"
    ARQUIVO_SEM_NOME_MENSAGEM = "Digite um nome para o arquivo!"
    ARQUIVO_SEM_NOME_TITULO = "Arquivo sem nome"
    SEM_DADOS_MENSAGEM = "Nada encontrado!"
    SEM_DADOS_TITULO = "Sem dados"
    ADICIONE_DATA_MENSAGEM = "Adicione uma data v�lida!"
    ADICIONE_DATA_TITULO = "Data inv�lida"
    ADICIONE_STATUS_MENSAGEM = "Selecione um Status para pesquisa!"
    ADICIONE_STATUS_TITULO = "Informe um Status"
End Sub
