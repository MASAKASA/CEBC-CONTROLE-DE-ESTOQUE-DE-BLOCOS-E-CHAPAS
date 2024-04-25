Attribute VB_Name = "daoStatus"
Option Explicit

Private status As objStatus

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
Function pesquisarPorNome(nomeStatus As String) As objStatus
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Cria��o e atribui��o do objeto
    Set status = ObjectFactory.factoryStatus(status)
    
    ' String para consulta
    strSql = "SELECT * FROM Status" _
        & " WHERE Nome_Status = '" & nomeStatus & "';"
        
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        status.id = rs.Fields("Id_Status").Value
        status.nome = rs.Fields("Nome_Status").Value
        
        rs.MoveNext
    Wend
    
    ' Libera espa�o na memoria
    Set rs = Nothing
    'Fechar conex�o com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorNome = status
    ' Libera espa�o na memoria
    Set status = Nothing
End Function

' Pesquisa objeto
Function listarBlocosFilter()

End Function
