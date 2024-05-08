Attribute VB_Name = "daoTipoPolimento"
Option Explicit

Private listaTiposPolimentos As Collection
Private tipoPolimento As objTipoPolimento

' Cadastra e edita objeto
Function cadastrarEEditar(tipoPolimento As objTipoPolimento)
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim rsAuxiliar As ADODB.Recordset ' Recordset para consulta
    Dim strSql As String ' String para consultas
    Dim cadastro As Boolean
    Dim i As Long

    ' Seta true em cadastro
    cadastro = True
    
    'Faz a consulta para saber se o código do bloco já exite
    strSql = "SELECT * FROM Tipo_Polimento" _
        & " WHERE Id_Polimento = " & tipoPolimento.id & ";"
    
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
        strSql = "INSERT INTO Tipo_Polimento ( Nome_Polimento )VALUES ('" & tipoPolimento.nome & "');"
        
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    Else
        ' Se edição
        strSql = "UPDATE Tipo_Polimento SET Nome_Polimento = '" & tipoPolimento.nome & "', " _
            & "' WHERE Id_Polimento = '" & tipoPolimento.id & "';"
            
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
    strSql = "DELETE * FROM Tipo_Polimento WHERE Id_Polimento = " & id & ";"
    
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
Function pesquisarPorId(id As String) As objTipoPolimento
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Criação e atribuição do objeto
    Set tipoPolimento = ObjectFactory.factoryTipoPolimento(tipoPolimento)
    
    ' String para consulta
    strSql = "SELECT * FROM Tipo_Polimento" _
        & " WHERE Id_Polimento = '" & id & "';"
        
    'Abrindo conexão com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        tipoPolimento.id = rs.Fields("Id_Polimento").Value
        tipoPolimento.nome = rs.Fields("Nome_Polimento").Value
        
        rs.MoveNext
    Wend
    
    ' Libera espaço na memoria
    Set rs = Nothing
    'Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorId = tipoPolimento
    ' Libera espaço na memoria
    Set tipoPolimento = Nothing
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
        & " WHERE Nome_Polimento = '" & nomeTipoPolimento & "';"
        
    'Abrindo conexão com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    ' Consulta banco
    
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        tipoPolimento.id = rs.Fields("Id_Polimento").Value
        tipoPolimento.nome = rs.Fields("Nome_Polimento").Value
        
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
Function listarTipoPolideiras() As Collection
    ' String para consultas
    Dim strSql As String ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    
    ' String para consulta
    strSql = "SELECT * FROM Tipo_Polimento ORDER BY Nome_Polimento;"
    
    'Abrindo conexão com banco
    Call conctarBanco
    ' Criação e atribuição dos objeto
    Set listaTiposPolimentos = ObjectFactory.factoryLista(listaTiposPolimentos)
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        ' Criação e atribuição do objeto
        Set tipoPolimento = ObjectFactory.factoryTipoPolimento(tipoPolimento)
        
        tipoPolimento.id = rs.Fields("Id_Polimento").Value
        tipoPolimento.nome = rs.Fields("Nome_Polimento").Value
        
        ' Adiciona na lista
        listaTiposPolimentos.Add tipoPolimento
        
        ' Libera espaço para nova pesquisa se ouver
        Set tipoPolimento = Nothing
        
        rs.MoveNext
    Wend
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    ' Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Retorna pesquisa
    Set listarTipoPolideiras = listaTiposPolimentos
    
    ' Libera espaço
    Set listaTiposPolimentos = Nothing
End Function
