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
    
    'Faz a consulta para saber se o código do bloco já exite
    strSql = "SELECT * FROM Polideiras" _
        & " WHERE Id_Polidoria = " & polideira.id & ";"
    
    ' Abrindo conexão com banco
    Call conctarBanco
    
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    Set rsAuxiliar = ObjectFactory.factoryRsAuxiliar(rsAuxiliar)
    ' Abrindo Recordset para consulta
    rsAuxiliar.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    ' Retorno da consulta
    While Not rsAuxiliar.EOF
        ' Seta false porquê vai ser uma edição
        cadastro = False
        
        rsAuxiliar.MoveNext
    Wend
    ' Libera recurso Recordset
    rsAuxiliar.Close
    Set rsAuxiliar = Nothing
    
    ' Direciona para os comandos certos de cadastro ou edição
    If cadastro = True Then ' Se cadastro
        'Concatenando comando SQL e cadastrando bloco no banco de dados
        strSql = "INSERT INTO Polideiras ( Nome_Polidoria )VALUES ('" & polideira.nome & "');"
        
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    Else
        ' Se edição
        strSql = "UPDATE Polideiras SET Nome_Pedreira = '" & polideira.nome & "', " _
            & "' WHERE Id_Polidoria = '" & polideira.id & "';"
            
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    End If
    
    Set rs = Nothing
    'Fechando conexão com banco
    Call fecharConexaoBanco
End Function

' Exclui objeto
Function excluir(id As String)
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim strSql As String ' String para consultas
    
    'Faz a consulta para saber se o código do bloco já exite
    strSql = "DELETE * FROM Polideiras WHERE Id_Polidoria = " & id & ";"
    
    ' Abrindo conexão com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    
    Set rs = Nothing
    'Fechando conexão com banco
    Call fecharConexaoBanco
End Function

' Pesquisa objeto por id
Function pesquisarPorId(id As String) As objPolideira
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Criação e atribuição do objeto
    Set polideira = ObjectFactory.factoryPolideira(polideira)
    
    ' String para consulta
    strSql = "SELECT * FROM Polideiras" _
        & " WHERE Id_Polidoria = '" & id & "';"
        
    'Abrindo conexão com banco
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
    
    ' Libera espaço na memoria
    Set rs = Nothing
    'Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorId = polideira
    ' Libera espaço na memoria
    Set polideira = Nothing
End Function

' Pesquisa objeto por nome
Function pesquisarPorNome(nomePolideira As String) As objPolideira
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Criação e atribuição do objeto
    Set polideira = ObjectFactory.factoryPolideira(polideira)
    
    ' String para consulta
    strSql = "SELECT * FROM Polidorias" _
        & " WHERE Nome_Polidoria = '" & nomePolideira & "';"
        
    'Abrindo conexão com banco
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
    
    ' Libera espaço na memoria
    Set rs = Nothing
    'Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorNome = polideira
    ' Libera espaço na memoria
    Set polideira = Nothing
End Function

' Pesquisa objeto
Function listarPolideiras() As Collection
    ' String para consultas
    Dim strSql As String ' String para consultas
    Dim rsBloco As ADODB.Recordset ' Recordset para consulta principal
    
    ' String para consulta
    strSql = "SELECT * FROM Polideiras ORDER BY Nome_Polidoria;"
    
    'Abrindo conexão com banco
    Call conctarBanco
    ' Criação e atribuição dos objeto
    Set listaPolideiras = ObjectFactory.factoryLista(listaPolideiras)
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    
    While Not rs.EOF
        ' Criação e atribuição do objeto
        Set polideira = ObjectFactory.factoryPolideira(polideira)
        
        polideira.id = rs.Fields("Id_Polidoria").Value
        polideira.nome = rs.Fields("Nome_Polidoria").Value
        
        ' Adiciona na lista
        listaPolideiras.Add polideira
        
        ' Libera espaço para nova pesquisa se ouver
        Set polideira = Nothing
        
        rs.MoveNext
    Wend
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    ' Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Retorna pesquisa
    Set listarPolideiras = listaPolideiras
    
    ' Libera espaço
    Set listaPolideiras = Nothing
End Function
