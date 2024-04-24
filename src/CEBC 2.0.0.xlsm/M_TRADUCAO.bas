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
Global VALOR_M3_MENSAGEM As String
Global VALOR_M3_TITULO As String
Global LARG_BLOCO_MENSAGEM As String
Global LARG_BLOCO_TITULO As String
Global ALT_BLOCO_MENSAGEM As String
Global ALT_BLOCO_TITULO As String
Global COMP_BLOCO_MENSAGEM As String
Global COMP_BLOCO_TITULO As String
Global TEM_NOTA_MENSAGEM As String
Global TEM_NOTA_TITULO As String
Global TIPO_MATERIAL_MENSAGEM As String
Global TIPO_MATERIAL_TITULO As String
Global NOME_BLOCO_MENSAGEM As String
Global NOME_BLOCO_PEDREIRA_TITULO As String
Global NUMERO_BLOCO_PEDREIRA_MENSAGEM As String
Global NUMERO_BLOCO_PEDREIRA_TITULO As String
Global NOME_PEDREIRA_MENSAGEM As String
Global NOME_PEDREIRA_TITULO As String
Global STATUS_SERRARIA_MENSAGEM As String
Global STATUS_SERRARIA_TITULO As String

' TELA ESTOQUE BLOCOS M�
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
    ' TELA CADASTRO DE BLOCO
    ERRO_DESCONHECIDO_MENSAGEM = "Erro desconhecido! Procurar o suporte!"
    ERRO_DESCONHECIDO_TITULO = "Erro desconhecido"
    SUCESSO_EDICAO_MENSAGEM = "Edi��o feito com sucesso!"
    SUCESSO_EDICAO_TITULO = "Edi��o"
    CADASTRO_CONFIRMADO_MENSAGEM = "Cadastrado com sucesso!"
    CADASTRO_CONFIRMADO_TITULO = "Cadastro confirmado"
    ACAO_CANCELADA_MENSAGEM = "A a��o foi cancelada!"
    ACAO_CANCELADA_TITULO = "Cancelamento"
    CONFIRMACAO_CADASTRO_MENSAGEM = "Confira se o n�mero e descri��o/material do bloco est�o corretos, pois a jun��o deles ir� criar o ID do bloco no sistema. ID do bloco n�o poder� ser alterado posteriormente. Tudo conferido e podemos seguir com o cadastro?"
    CONFIRMACAO_CADASTRO_TITULO = "Aten��o - Confirma��o"
    VALOR_M3_MENSAGEM = "Adicione o Valor M�!"
    VALOR_M3_TITULO = "Valor M� n�o informado"
    LARG_BLOCO_MENSAGEM = "Adicione a largura!"
    LARG_BLOCO_TITULO = "Largura n�o informada"
    ALT_BLOCO_MENSAGEM = "Adicione a altura!"
    ALT_BLOCO_TITULO = "Altura n�o informada"
    COMP_BLOCO_MENSAGEM = "Adicione o comprimento!"
    COMP_BLOCO_TITULO = "Comprimento n�o informado"
    TEM_NOTA_MENSAGEM = "Informe se o bloco tem a nota fiscal!"
    TEM_NOTA_TITULO = "Nota fiscal!"
    TIPO_MATERIAL_MENSAGEM = "Selecione o tipo do material!"
    TIPO_MATERIAL_TITULO = "Tipo de material n�o informado"
    NOME_BLOCO_MENSAGEM = "Adicone a descri��o do bloco!"
    NOME_BLOCO_PEDREIRA_TITULO = "Descri��o do bloco n�o informada"
    NUMERO_BLOCO_PEDREIRA_MENSAGEM = "Adicone o n�mero do bloco!"
    NUMERO_BLOCO_PEDREIRA_TITULO = "ID do bloco n�o informado"
    STATUS_SERRARIA_MENSAGEM = "Se o bloco j� estiver na serraria, selecione o nome da 'Serraria'!, Caso n�o esteja, selecione a op��o: 'BLOCO EST� NA? PEDREIRA"
    STATUS_SERRARIA_TITULO = "Status do bloco"
    NOME_PEDREIRA_MENSAGEM = "Selecione o nome da pedreira!"
    NOME_PEDREIRA_TITULO = "Nome da pedreira n�o informada"
    
    ' TELA ESTOQUE BLOCOS M�
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
