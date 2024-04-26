Attribute VB_Name = "daoStatus"
Option Explicit

Private listaStatus As Collection
Private status As objStatus

' Cadastra e edita objeto
Function cadastrarEEditar(status As objStatus)
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim rsAuxiliar As ADODB.Recordset ' Recordset para consulta
    Dim strSql As String ' String para consultas
    Dim cadastro As Boolean
    Dim i As Long

    ' Seta true em cadastro
    cadastro = True
    
    'Faz a consulta para saber se o código do bloco já exite
    strSql = "SELECT * FROM Status" _
        & " WHERE Id_Status = " & status.id & ";"
    
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
        strSql = "INSERT INTO Status ( Nome_Status )VALUES ('" & status.nome & "');"
        
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    Else
        ' Se edição
        strSql = "UPDATE Status SET Nome_Status = '" & status.nome & "', " _
            & "' WHERE Id_Status = '" & status.id & "';"
            
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
    strSql = "DELETE * FROM Status WHERE Id_Status = " & id & ";"
    
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
Function pesquisarPorId(id As String) As objStatus
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Criação e atribuição do objeto
    Set status = ObjectFactory.factoryStatus(status)
    
    ' String para consulta
    strSql = "SELECT * FROM Status" _
        & " WHERE Id_Status = '" & id & "';"
        
    'Abrindo conexão com banco
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
    
    ' Libera espaço na memoria
    Set rs = Nothing
    'Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorId = pedreira
    ' Libera espaço na memoria
    Set status = Nothing
End Function

' Pesquisa objeto por nome
Function pesquisarPorNome(nomeStatus As String) As objStatus
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Criação e atribuição do objeto
    Set status = ObjectFactory.factoryStatus(status)
    
    ' String para consulta
    strSql = "SELECT * FROM Status" _
        & " WHERE Nome_Status = '" & nomeStatus & "';"
        
    'Abrindo conexão com banco
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
    
    ' Libera espaço na memoria
    Set rs = Nothing
    'Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorNome = status
    ' Libera espaço na memoria
    Set status = Nothing
End Function

' Pesquisa objeto
Function listarStatus() As Collection
    ' String para consultas
    Dim strSql As String ' String para consultas
    Dim rsBloco As ADODB.Recordset ' Recordset para consulta principal
    
    ' String para consulta
    strSql = "SELECT * FROM Status ORDER BY Nome_Status;"
    
    'Abrindo conexão com banco
    Call conctarBanco
    ' Criação e atribuição dos objeto
    Set listaStatus = ObjectFactory.factoryLista(listaStatus)
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    
    While Not rs.EOF
        ' Criação e atribuição do objeto
        Set status = ObjectFactory.factoryStatus(status)
        
        status.id = rs.Fields("Id_Status").Value
        status.nome = rs.Fields("Nome_Status").Value
        
        ' Adiciona na lista
        listaStatus.Add status
        
        ' Libera espaço para nova pesquisa se ouver
        Set status = Nothing
        
        rs.MoveNext
    Wend
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    ' Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Retorna pesquisa
    Set listarStatus = listaStatus
    
    ' Libera espaço
    Set listaStatus = Nothing
End Function
