Attribute VB_Name = "daoPedreira"
Option Explicit

Private pedreira As objPedreira

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
Function pesquisarPorNome(nomePedreira As String) As objPedreira
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Cria��o e atribui��o do objeto
    Set pedreira = ObjectFactory.factoryPedreira(pedreira)
    
    ' String para consulta
    strSql = "SELECT * FROM Pedreiras" _
        & " WHERE Nome_Pedreira = '" & nomePedreira & "';"
        
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        pedreira.id = rs.Fields("Id_Pedreira").Value
        pedreira.nome = rs.Fields("Nome_Pedreira").Value
        
        rs.MoveNext
    Wend
    
    ' Libera espa�o na memoria
    Set rs = Nothing
    'Fechar conex�o com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorNome = pedreira
    ' Libera espa�o na memoria
    Set pedreira = Nothing
End Function

' Pesquisa objeto
Function listarBlocosFilter()

End Function
