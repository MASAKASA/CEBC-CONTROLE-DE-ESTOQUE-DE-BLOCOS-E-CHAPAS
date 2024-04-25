Attribute VB_Name = "daoSerrada"
Option Explicit

Private serraria As objSerraria

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
Function pesquisarPorNome(nomeSerraria As String) As objSerraria
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Cria��o e atribui��o do objeto
    Set serraria = ObjectFactory.factorySerraria(serraria)
    
    ' String para consulta
    strSql = "SELECT * FROM Serrarias" _
        & " WHERE Nome_Serraria = '" & nomeSerraria & "';"
        
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        serraria.id = rs.Fields("Id_Serraria").Value
        serraria.nome = rs.Fields("Nome_Serraria").Value
        
        rs.MoveNext
    Wend
    
    ' Libera espa�o na memoria
    Set rs = Nothing
    'Fechar conex�o com banco
    Call fecharConexaoBanco
    ' Retorno
    Set pesquisarPorNome = serraria
    ' Libera espa�o na memoria
    Set serraria = Nothing
End Function

' Pesquisa objeto
Function listarBlocosFilter()

End Function
