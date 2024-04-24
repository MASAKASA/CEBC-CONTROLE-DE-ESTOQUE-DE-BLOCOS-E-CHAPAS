Attribute VB_Name = "daoPolideira"
Option Explicit

Private polideira As objPolideira

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
Function pesquisarPorNome(nomePolideira As String) As objPolideira
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Cria��o e atribui��o do objeto
    Set polideira = ObjectFactory.factoryPolideira(polideira)
    
    ' String para consulta
    strSql = "SELECT * FROM Polidorias" _
        & " WHERE Nome_Polidoria = '" & nomePolideira & "';"
        
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Consulta banco
    rs.Open strSql, BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        polideira.id = rs.Fields("Id_Polidoria").Value
        polideira.nome = rs.Fields("Nome_Polidoria").Value
    Wend
    
    'Fechar conex�o com banco
    Call fecharConexaoBanco
    ' Retorno
    Set pesquisarPorNome = polideira
    ' Libera espa�o na memoria
    Set polideira = Nothing
End Function

' Pesquisa objeto
Function listarBlocosFilter()

End Function
