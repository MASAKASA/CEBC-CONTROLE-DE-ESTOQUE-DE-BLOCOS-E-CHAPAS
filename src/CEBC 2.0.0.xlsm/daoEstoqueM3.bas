Attribute VB_Name = "daoEstoqueM3"
Option Explicit

Private estoque As objEstoque

' Cadastra e edita objeto
Function cadastrarEEditar()

End Function

' Exclui objeto
Function excluir()

End Function

' Pesquisa objeto por id
Function pesquisarPorId()

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
Function listarBlocosFilter()

End Function
