Attribute VB_Name = "daoPedreira"
Option Explicit

Private listaPedreiras As Collection
Private pedreira As objPedreira

' Cadastra e edita objeto
Function cadastrarEEditar(pedreira As objPedreira)
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim rsAuxiliar As ADODB.Recordset ' Recordset para consulta
    Dim strSql As String ' String para consultas
    Dim cadastro As Boolean
    Dim i As Long

    ' Seta true em cadastro
    cadastro = True
    
    'Faz a consulta para saber se o c�digo do bloco j� exite
    strSql = "SELECT * FROM Pedreiras" _
        & " WHERE Id_Pedreira = " & pedreira.id & ";"
    
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
        strSql = "INSERT INTO Pedreiras ( Nome_Pedreira )VALUES ('" & pedreira.nome & "');"
        
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    Else
        ' Se edi��o
        strSql = "UPDATE Pedreiras SET Nome_Pedreira = '" & pedreira.nome & "', " _
            & "' WHERE Id_Pedreira = '" & pedreira.id & "';"
            
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
    strSql = "DELETE * FROM Pedreiras WHERE Id_Pedreira = " & id & ";"
    
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
Function pesquisarPorId(id As String) As objPedreira
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Cria��o e atribui��o do objeto
    Set pedreira = ObjectFactory.factoryPedreira(pedreira)
    
    ' String para consulta
    strSql = "SELECT * FROM Pedreiras" _
        & " WHERE Id_Pedreira = '" & id & "';"
        
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
    Set pesquisarPorId = pedreira
    ' Libera espa�o na memoria
    Set pedreira = Nothing
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
Function listarPedreiras() As Collection
    ' String para consultas
    Dim strSql As String ' String para consultas
    Dim rsBloco As ADODB.Recordset ' Recordset para consulta principal
    
    ' String para consulta
    strSql = "SELECT * FROM Pedreiras ORDER BY Nome_Pedreira;"
    
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Cria��o e atribui��o dos objeto
    Set listaPedreiras = ObjectFactory.factoryLista(listaPedreiras)
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    
    While Not rs.EOF
        ' Cria��o e atribui��o do objeto
        Set pedreira = ObjectFactory.factoryPedreira(pedreira)
        
        pedreira.id = rs.Fields("Id_Pedreira").Value
        pedreira.nome = rs.Fields("Nome_Pedreira").Value
        
        ' Adiciona na lista
        listaPedreiras.Add pedreira
        
        ' Libera espa�o para nova pesquisa se ouver
        Set pedreira = Nothing
        
        rs.MoveNext
    Wend
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    ' Fechar conex�o com banco
    Call fecharConexaoBanco
    
    ' Retorna pesquisa
    Set listarPedreiras = listaPedreiras
    
    ' Libera espa�o
    Set listaPedreiras = Nothing
End Function
