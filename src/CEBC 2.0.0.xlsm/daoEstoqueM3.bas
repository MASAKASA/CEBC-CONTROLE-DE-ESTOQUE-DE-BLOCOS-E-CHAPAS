Attribute VB_Name = "daoEstoqueM3"
Option Explicit

Private listaEstoquesM3 As Collection
Private estoque As objEstoque

' Cadastra e edita objeto
Function cadastrarEEditar(estoqueM3 As objEstoque)
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim rsAuxiliar As ADODB.Recordset ' Recordset para consulta
    Dim strSql As String ' String para consultas
    Dim cadastro As Boolean
    Dim i As Long

    ' Seta true em cadastro
    cadastro = True
    
    'Faz a consulta para saber se o c�digo do bloco j� exite
    strSql = "SELECT * FROM Estoque_blocos" _
        & " WHERE Id_Estoque = " & estoqueM3.id & ";"
    
    ' Abrindo conex�o com banco
    Call conctarBanco
    
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    Set rsAuxiliar = ObjectFactory.factoryRsAuxiliar(rsAuxiliar)
    ' Abrindo Recordset para consulta
    rsAuxiliar.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    ' Retorno da consulta
    While Not rsAuxiliar.EOF
        ' Seta false porqu� vai ser uma edi��o
        cadastro = False
        
        rsAuxiliar.MoveNext
    Wend
    ' Libera recurso Recordset
    rsAuxiliar.Close
    Set rsAuxiliar = Nothing
    
    ' Direciona para os comandos certos de cadastro ou edi��o
    If cadastro = True Then ' Se cadastro
        'Concatenando comando SQL e cadastrando bloco no banco de dados
        strSql = "INSERT INTO Estoque_blocos ( Empresa )VALUES ('" & estoqueM3.nome & "');"
        
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    Else
        ' Se edi��o
        strSql = "UPDATE Estoque_blocos SET Empresa = '" & estoqueM3.nome & "', " _
            & "' WHERE Id_Estoque = '" & estoqueM3.id & "';"
            
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    End If
    
    Set rs = Nothing
    'Fechando conex�o com banco
    Call fecharConexaoBanco
End Function

' Exclui objeto
Function excluir(id As String)
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim strSql As String ' String para consultas
    
    'Faz a consulta para saber se o c�digo do bloco j� exite
    strSql = "DELETE * FROM Estoque_blocos WHERE Id_Estoque = " & id & ";"
    
    ' Abrindo conex�o com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    
    Set rs = Nothing
    'Fechando conex�o com banco
    Call fecharConexaoBanco
End Function

' Pesquisa objeto por id
Function pesquisarPorId(id As String) As objEstoque
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Cria��o e atribui��o do objeto
    Set estoque = ObjectFactory.factoryDestino(estoque)
    
    ' String para consulta
    strSql = "SELECT * FROM Estoque_blocos" _
        & " WHERE Id_Estoque = '" & id & "';"
        
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        estoque.id = rs.Fields("Id_Estoque").Value
        estoque.nome = rs.Fields("Empresa").Value
        
        rs.MoveNext
    Wend
    
    ' Libera espa�o na memoria
    Set rs = Nothing
    'Fechar conex�o com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorId = estoque
    ' Libera espa�o na memoria
    Set estoque = Nothing
End Function

' Pesquisa objeto por nome
Function pesquisarPorNome(nomeEmpresa As String) As objEstoque
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Cria��o e atribui��o do objeto
    Set estoque = ObjectFactory.factoryEstoque(estoque)
    
    ' String para consulta
    strSql = "SELECT * FROM Estoque_blocos" _
        & " WHERE Empresa = '" & nomeEmpresa & "';"
        
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        estoque.id = rs.Fields("Id_Estoque").Value
        estoque.nome = rs.Fields("Empresa").Value
        
        rs.MoveNext
    Wend
    
    ' Libera espa�o na memoria
    Set rs = Nothing
    'Fechar conex�o com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorNome = estoque
    ' Libera espa�o na memoria
    Set estoque = Nothing
End Function

' Pesquisa objeto
Function listarEstoqueM3() As Collection
    ' String para consultas
    Dim strSql As String ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    
    ' String para consulta
    strSql = "SELECT * FROM Estoque_blocos ORDER BY Empresa;"
    
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Cria��o e atribui��o dos objeto
    Set listaEstoquesM3 = ObjectFactory.factoryLista(listaEstoquesM3)
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        ' Cria��o e atribui��o do objeto
        Set estoque = ObjectFactory.factoryEstoque(estoque)
        
        estoque.id = rs.Fields("Id_Estoque").Value
        estoque.nome = rs.Fields("Empresa").Value
        
        ' Adiciona na lista
        listaEstoquesM3.Add estoque
        
        ' Libera espa�o para nova pesquisa se ouver
        Set estoque = Nothing
        
        rs.MoveNext
    Wend
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    ' Fechar conex�o com banco
    Call fecharConexaoBanco
    
    ' Retorna pesquisa
    Set listarEstoqueM3 = listaEstoquesM3
    
    ' Libera espa�o
    Set listaEstoquesM3 = Nothing
End Function
