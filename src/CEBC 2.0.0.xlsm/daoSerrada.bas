Attribute VB_Name = "daoSerrada"
Option Explicit

Private listaSerrarias As Collection
Private serraria As objSerraria

' Cadastra e edita objeto
Function cadastrarEEditar(serraria As objSerraria)
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim rsAuxiliar As ADODB.Recordset ' Recordset para consulta
    Dim strSql As String ' String para consultas
    Dim cadastro As Boolean
    Dim i As Long

    ' Seta true em cadastro
    cadastro = True
    
    'Faz a consulta para saber se o c�digo do bloco j� exite
    strSql = "SELECT * FROM Serrarias" _
        & " WHERE Id_Serraria = " & serraria.id & ";"
    
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
        strSql = "INSERT INTO Serrarias ( Nome_Serraria )VALUES ('" & serraria.nome & "');"
        
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    Else
        ' Se edi��o
        strSql = "UPDATE Serrarias SET Nome_Pedreira = '" & serraria.nome & "', " _
            & "' WHERE Id_Serraria = '" & serraria.id & "';"
            
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
    strSql = "DELETE * FROM Serrarias WHERE Id_Serraria = " & id & ";"
    
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
Function pesquisarPorId(id As String) As objSerraria
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Cria��o e atribui��o do objeto
    Set serraria = ObjectFactory.factorySerraria(serraria)
    
    ' String para consulta
    strSql = "SELECT * FROM Serrarias" _
        & " WHERE Id_Pedreira = '" & id & "';"
        
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        serraria.id = rs.Fields("Id_Pedreira").Value
        serraria.nome = rs.Fields("Nome_Serraria").Value
        
        rs.MoveNext
    Wend
    
    ' Libera espa�o na memoria
    Set rs = Nothing
    'Fechar conex�o com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorId = serraria
    ' Libera espa�o na memoria
    Set serraria = Nothing
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
Function listarSerrarias() As Collection
    ' String para consultas
    Dim strSql As String ' String para consultas
    Dim rsBloco As ADODB.Recordset ' Recordset para consulta principal
    
    ' String para consulta
    strSql = "SELECT * FROM Serrarias ORDER BY Nome_Serraria;"
    
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Cria��o e atribui��o dos objeto
    Set listaSerrarias = ObjectFactory.factoryLista(listaSerrarias)
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    
    While Not rs.EOF
        ' Cria��o e atribui��o do objeto
        Set serraria = ObjectFactory.factorySerraria(serraria)
        
        serraria.id = rs.Fields("Id_Pedreira").Value
        serraria.nome = rs.Fields("Nome_Pedreira").Value
        
        ' Adiciona na lista
        listaSerrarias.Add serraria
        
        ' Libera espa�o para nova pesquisa se ouver
        Set serraria = Nothing
        
        rs.MoveNext
    Wend
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    ' Fechar conex�o com banco
    Call fecharConexaoBanco
    
    ' Retorna pesquisa
    Set listarSerrarias = listaSerrarias
    
    ' Libera espa�o
    Set listaSerrarias = Nothing
End Function
