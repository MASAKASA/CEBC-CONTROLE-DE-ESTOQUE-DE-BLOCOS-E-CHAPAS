Attribute VB_Name = "daoEstoqueChapa"
Option Explicit

Private listaEstoques As Collection
Private estoque As objEstoqueChapa

' Cadastra e edita objeto
Function cadastrarEEditar(estoque As objEstoqueChapa)
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim rsAuxiliar As ADODB.Recordset ' Recordset para consulta
    Dim strSql As String ' String para consultas
    Dim cadastro As Boolean
    Dim i As Long

    ' Seta true em cadastro
    cadastro = True

    'Faz a consulta para saber se o código do bloco já exite
    strSql = "SELECT * FROM Estoque_chapas" _
        & " WHERE Id_Estoque = " & estoque.id & ";"

    ' Abrindo conexão com banco
    Call conctarBanco

    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    Set rsAuxiliar = ObjectFactory.factoryRsAuxiliar(rsAuxiliar)
    ' Abrindo Recordset para consulta
    rsAuxiliar.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    ' Retorno da consulta
    While Not rsAuxiliar.EOF
        ' Seta false porquê vai ser uma edição
        cadastro = False

        rsAuxiliar.MoveNext
    Wend
    ' Libera recurso Recordset
    rsAuxiliar.Close
    Set rsAuxiliar = Nothing

    ' Direciona para os comandos certos de cadastro ou edição
    If cadastro = True Then ' Se cadastro
        'Concatenando comando SQL e cadastrando bloco no banco de dados
        strSql = "INSERT INTO Estoque_chapas ( Nome_Empresa )VALUES ('" & estoque.nome & "');"

        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    Else
        ' Se edição
        strSql = "UPDATE Estoque_chapas SET Nome_Empresa = '" & estoque.nome & "', " _
            & "' WHERE Id_Estoque = '" & estoque.id & "';"

        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    End If

    Set rs = Nothing
    'Fechando conexão com banco
    Call fecharConexaoBanco
End Function

' Exclui objeto
Function excluir(id As String)
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim strSql As String ' String para consultas

    'Faz a consulta para saber se o código do bloco já exite
    strSql = "DELETE * FROM Estoque_chapas WHERE Id_Estoque = " & id & ";"

    ' Abrindo conexão com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)

    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic

    Set rs = Nothing
    'Fechando conexão com banco
    Call fecharConexaoBanco
End Function

' Pesquisa objeto por id
Function pesquisarPorId(id As String) As objEstoqueChapa
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String

    ' Criação e atribuição do objeto
    Set estoque = ObjectFactory.factoryEstoqueChapas(estoque)

    ' String para consulta
    strSql = "SELECT * FROM Estoque_chapas" _
        & " WHERE Id_Estoque = '" & id & "';"

    'Abrindo conexão com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly

    While Not rs.EOF
        estoque.id = rs.Fields("Id_Estoque").Value
        estoque.nome = rs.Fields("Nome_Empresa").Value

        rs.MoveNext
    Wend

    ' Libera espaço na memoria
    Set rs = Nothing
    'Fechar conexão com banco
    Call fecharConexaoBanco

    ' Retorno
    Set pesquisarPorId = estoque
    ' Libera espaço na memoria
    Set estoque = Nothing
End Function

' Pesquisa objeto por nome
Function pesquisarPorNome(nomeEmpresa As String) As objEstoqueChapa
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String

    ' Criação e atribuição do objeto
    Set estoque = ObjectFactory.factoryEstoqueChapas(estoque)

    ' String para consulta
    strSql = "SELECT * FROM Estoque_chapas" _
        & " WHERE Nome_Empresa = '" & nomeEmpresa & "';"

    'Abrindo conexão com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly

    While Not rs.EOF
        estoque.id = rs.Fields("Id_Estoque").Value
        estoque.nome = rs.Fields("Nome_Empresa").Value

        rs.MoveNext
    Wend

    ' Libera espaço na memoria
    Set rs = Nothing
    'Fechar conexão com banco
    Call fecharConexaoBanco

    ' Retorno
    Set pesquisarPorNome = estoque
    ' Libera espaço na memoria
    Set estoque = Nothing
End Function

' Pesquisa objeto
Function listarEstoqueChapas() As Collection
    ' String para consultas
    Dim strSql As String ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal

    ' String para consulta
    strSql = "SELECT * FROM Estoque_chapas ORDER BY Nome_Empresa;"

    'Abrindo conexão com banco
    Call conctarBanco
    ' Criação e atribuição dos objeto
    Set listaEstoques = ObjectFactory.factoryLista(listaEstoques)
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly

    While Not rs.EOF
        ' Criação e atribuição do objeto
        Set estoque = ObjectFactory.factoryEstoqueChapas(estoque)

        estoque.id = rs.Fields("Id_Estoque").Value
        estoque.nome = rs.Fields("Nome_Empresa").Value

        ' Adiciona na lista
        listaEstoques.Add estoque

        ' Libera espaço para nova pesquisa se ouver
        Set estoque = Nothing

        rs.MoveNext
    Wend
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    ' Fechar conexão com banco
    Call fecharConexaoBanco

    ' Retorna pesquisa
    Set listarEstoqueChapas = listaEstoques

    ' Libera espaço
    Set listaEstoques = Nothing
End Function

