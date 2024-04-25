Attribute VB_Name = "daoTipoMaterial"
Option Explicit

Private tipoMaterial As objTipoMaterial

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
Function pesquisarPorNome(nomeTipoMaterial As String) As objTipoMaterial
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Criação e atribuição do objeto
    Set tipoMaterial = ObjectFactory.factoryTipoMaterial(tipoMaterial)
    
    ' String para consulta
    strSql = "SELECT * FROM Tipo_Material" _
        & " WHERE Nome_Tipo_Material = '" & nomeTipoMaterial & "';"
        
    'Abrindo conexão com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        tipoMaterial.id = rs.Fields("Id_Tipo_Material").Value
        tipoMaterial.nome = rs.Fields("Nome_Tipo_Material").Value
        
        rs.MoveNext
    Wend
    
    ' Libera espaço na memoria
    Set rs = Nothing
    'Fechar conexão com banco
    Call fecharConexaoBanco
    ' Retorno
    Set pesquisarPorNome = tipoMaterial
    ' Libera espaço na memoria
    Set tipoMaterial = Nothing
End Function

' Pesquisa objeto
Function listarBlocosFilter()

End Function
