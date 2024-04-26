Attribute VB_Name = "daoMotorista"
Option Explicit

Private listaMotoristas As Collection
Private motorista As objMotorista

' Cadastra e edita objeto
Function cadastrarEEditar(motorista As objMotorista)
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim rsAuxiliar As ADODB.Recordset ' Recordset para consulta
    Dim strSql As String ' String para consultas
    Dim cadastro As Boolean
    Dim i As Long

    ' Seta true em cadastro
    cadastro = True
    
    'Faz a consulta para saber se o c�digo do bloco j� exite
    strSql = "SELECT * FROM Motoristas" _
        & " WHERE Id_Motorista = " & motorista.id & ";"
    
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
        strSql = "INSERT INTO Motoristas ( Nome_Motorista )VALUES ('" & motorista.nome & "');"
        
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    Else
        ' Se edi��o
        strSql = "UPDATE Motoristas SET Nome_Pedreira = '" & motorista.nome & "', " _
            & "' WHERE Id_Motorista = '" & motorista.id & "';"
            
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
    strSql = "DELETE * FROM Motoristas WHERE Id_Motorista = " & id & ";"
    
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
Function pesquisarPorId(id As String) As objMotoista
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Cria��o e atribui��o do objeto
    Set motorista = ObjectFactory.factoryMotorista(motorista)
    
    ' String para consulta
    strSql = "SELECT * FROM Motoristas" _
        & " WHERE Id_Motorista = '" & id & "';"
        
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        motorista.id = rs.Fields("Id_Motorista").Value
        motorista.nome = rs.Fields("Nome_Pedreira").Value
        
        rs.MoveNext
    Wend
    
    ' Libera espa�o na memoria
    Set rs = Nothing
    'Fechar conex�o com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorId = motorista
    ' Libera espa�o na memoria
    Set motorista = Nothing
End Function

' Pesquisa objeto por nome
Function pesquisarPorNome(nomeMotorista) As objMotoista
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Cria��o e atribui��o do objeto
    Set motorista = ObjectFactory.factoryMotorista(motorista)
    
    ' String para consulta
    strSql = "SELECT * FROM Motoristas" _
        & " WHERE Nome_Pedreira = '" & nomeMotorista & "';"
        
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        motorista.id = rs.Fields("Id_Motorista").Value
        motorista.nome = rs.Fields("Nome_Pedreira").Value
        
        rs.MoveNext
    Wend
    
    ' Libera espa�o na memoria
    Set rs = Nothing
    'Fechar conex�o com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorNome = motorista
    ' Libera espa�o na memoria
    Set motorista = Nothing
End Function

' Pesquisa objeto
Function listarMotoristas() As Collection
    ' String para consultas
    Dim strSql As String ' String para consultas
    Dim rsBloco As ADODB.Recordset ' Recordset para consulta principal
    
    ' String para consulta
    strSql = "SELECT * FROM Motoristas ORDER BY Nome_Pedreira;"
    
    'Abrindo conex�o com banco
    Call conctarBanco
    ' Cria��o e atribui��o dos objeto
    Set listaMotoristas = ObjectFactory.factoryLista(listaMotoristas)
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsBloco(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        ' Cria��o e atribui��o do objeto
        Set motorista = ObjectFactory.factoryMotorista(motorista)
        
        motorista.id = rs.Fields("Id_Motorista").Value
        motorista.nome = rs.Fields("Nome_Pedreira").Value
        
        ' Adiciona na lista
        listaMotoristas.Add motorista
        
        ' Libera espa�o para nova pesquisa se ouver
        Set motorista = Nothing
        
        rs.MoveNext
    Wend
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    ' Fechar conex�o com banco
    Call fecharConexaoBanco
    
    ' Retorna pesquisa
    Set listarMotoristas = listaMotoristas
    
    ' Libera espa�o
    Set listaMotoristas = Nothing
End Function
