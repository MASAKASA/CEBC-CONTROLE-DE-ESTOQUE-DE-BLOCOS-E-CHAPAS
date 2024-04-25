Attribute VB_Name = "daoTipoPolimento"
Option Explicit

Private tipoPolimento As objTipoPolimento

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
Function pesquisarPorNome(nomeTipoPolimento As String) As objTipoPolimento
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Criação e atribuição do objeto
    Set tipoPolimento = ObjectFactory.factoryTipoPolimento(tipoPolimento)
    
    ' String para consulta
    strSql = "SELECT * FROM Tipo_Polimento" _
        & " WHERE Id_Polimento = '" & nomeTipoPolimento & "';"
        
    'Abrindo conexão com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        tipoPolimento.id = rs.Fields("Id_Polimento").Value
        tipoPolimento.nome = rs.Fields("Id_Polimento").Value
        
        rs.MoveNext
    Wend
    
    ' Libera espaço na memoria
    Set rs = Nothing
    'Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorNome = tipoPolimento
    ' Libera espaço na memoria
    Set tipoPolimento = Nothing
End Function

' Pesquisa objeto
Function listarBlocosFilter()

End Function
