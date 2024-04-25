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
    
    ' Criação e atribuição do objeto
    Set estoque = ObjectFactory.factoryEstoque(estoque)
    
    ' String para consulta
    strSql = "SELECT * FROM Estoque_blocos" _
        & " WHERE Empresa = '" & nomeEmpresa & "';"
        
    'Abrindo conexão com banco
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
Function listarBlocosFilter()

End Function
