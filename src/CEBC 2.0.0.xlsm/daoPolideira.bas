Attribute VB_Name = "daoPolideira"
Option Explicit

Private listaPolideiras As Collection
Private polideira As objPolideira

' Cadastra e edita objeto
Function cadastrarEEditar(polideira As objPolideira)
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim rsAuxiliar As ADODB.Recordset ' Recordset para consulta
    Dim strSql As String ' String para consultas
    Dim cadastro As Boolean
    Dim i As Long

    ' Seta true em cadastro
    cadastro = True
    
    'Faz a consulta para saber se o c�digo do bloco j� exite
    strSql = "SELECT * FROM Polideiras" _
        & " WHERE Id_Polidoria = " & polideira.id & ";"
    
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
        strSql = "INSERT INTO Polideiras ( Nome_Polidoria )VALUES ('" & polideira.nome & "');"
        
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    Else
        ' Se edi��o
        strSql = "UPDATE Polideiras SET Nome_Pedreira = '" & polideira.nome & "', " _
            & "' WHERE Id_Polidoria = '" & polideira.id & "';"
            
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
    strSql = "DELETE * FROM Polideiras WHERE Id_Polidoria = " & id & ";"
    
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
Function pesquisarPorId(id As String) As objPolideira
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Cria��o e atribui��o do objeto
    Set polideira = ObjectFactory.factoryPolideira(polideira)
    
    ' String para consulta
    strSql = "SELECT * FROM Polideiras" _
        & " WHERE Id_Polidoria = '" & id & "';"
        
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        polideira.id = rs.Fields("Id_Polidoria").Value
        polideira.nome = rs.Fields("Nome_Polidoria").Value
        
        rs.MoveNext
    Wend
    
    ' Libera espa�o na memoria
    Set rs = Nothing
    'Fechar conex�o com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorId = polideira
    ' Libera espa�o na memoria
    Set polideira = Nothing
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
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        polideira.id = rs.Fields("Id_Polidoria").Value
        polideira.nome = rs.Fields("Nome_Polidoria").Value
        
        rs.MoveNext
    Wend
    
    ' Libera espa�o na memoria
    Set rs = Nothing
    'Fechar conex�o com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorNome = polideira
    ' Libera espa�o na memoria
    Set polideira = Nothing
End Function

' Pesquisa objeto
Function listarPolideiras() As Collection
    ' String para consultas
    Dim strSql As String ' String para consultas
    Dim rsBloco As ADODB.Recordset ' Recordset para consulta principal
    
    ' String para consulta
    strSql = "SELECT * FROM Polideiras ORDER BY Nome_Polidoria;"
    
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Cria��o e atribui��o dos objeto
    Set listaPolideiras = ObjectFactory.factoryLista(listaPolideiras)
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    
    While Not rs.EOF
        ' Cria��o e atribui��o do objeto
        Set polideira = ObjectFactory.factoryPolideira(polideira)
        
        polideira.id = rs.Fields("Id_Polidoria").Value
        polideira.nome = rs.Fields("Nome_Polidoria").Value
        
        ' Adiciona na lista
        listaPolideiras.Add polideira
        
        ' Libera espa�o para nova pesquisa se ouver
        Set polideira = Nothing
        
        rs.MoveNext
    Wend
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    ' Fechar conex�o com banco
    Call fecharConexaoBanco
    
    ' Retorna pesquisa
    Set listarPolideiras = listaPolideiras
    
    ' Libera espa�o
    Set listaPolideiras = Nothing
End Function
