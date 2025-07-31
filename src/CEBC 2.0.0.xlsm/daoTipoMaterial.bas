Attribute VB_Name = "daoTipoMaterial"
Option Explicit

Private listaTipoMaterial As Collection
Private tipoMaterial As objTipoMaterial

' Cadastra e edita objeto
Function cadastrarEEditar(tipoMaterial As objTipoMaterial)
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim rsAuxiliar As ADODB.Recordset ' Recordset para consulta
    Dim strSql As String ' String para consultas
    Dim cadastro As Boolean
    Dim i As Long

    ' Seta true em cadastro
    cadastro = True
    
    'Faz a consulta para saber se o código do bloco já exite
    strSql = "SELECT * FROM Tipo_Material" _
        & " WHERE Id_Tipo_Material = " & tipoMaterial.id & ";"
    
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
        strSql = "INSERT INTO Tipo_Material ( Nome_Tipo_Material )VALUES ('" & tipoMaterial.nome & "');"
        
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    Else
        ' Se edição
        strSql = "UPDATE Tipo_Material SET Nome_Pedreira = '" & tipoMaterial.nome & "', " _
            & "' WHERE Id_Tipo_Material = '" & tipoMaterial.id & "';"
            
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
    strSql = "DELETE * FROM Tipo_Material WHERE Id_Tipo_Material = " & id & ";"
    
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
Function pesquisarPorId(id As String)
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    ' Criação e atribuição do objeto
    Set tipoMaterial = ObjectFactory.factoryTipoMaterial(tipoMaterial)
    
    ' String para consulta
    strSql = "SELECT * FROM Tipo_Material" _
        & " WHERE Id_Tipo_Material = '" & id & "';"
        
    'Abrindo conexão com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        tipoMaterial.id = rs.Fields("Id_Tipo_Material").Value
        tipoMaterial.nome = rs.Fields("Nome_Pedreira").Value
        
        rs.MoveNext
    Wend
    
    ' Libera espaço na memoria
    Set rs = Nothing
    'Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorId = tipoMaterial
    ' Libera espaço na memoria
    Set tipoMaterial = Nothing
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
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
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
Function listarTiposMateriais() As Collection
    ' String para consultas
    Dim strSql As String ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    
    ' String para consulta
    strSql = "SELECT * FROM Tipo_Material ORDER BY Nome_Tipo_Material;"
    
    'Abrindo conexão com banco
    Call conctarBanco
    ' Criação e atribuição dos objeto
    Set listaTipoMaterial = ObjectFactory.factoryLista(listaTipoMaterial)
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        ' Criação e atribuição do objeto
        Set tipoMaterial = ObjectFactory.factoryTipoMaterial(tipoMaterial)
        
        tipoMaterial.id = rs.Fields("Id_Tipo_Material").Value
        tipoMaterial.nome = rs.Fields("Nome_Tipo_Material").Value
        
        ' Adiciona na lista
        listaTipoMaterial.Add tipoMaterial
        
        ' Libera espaço para nova pesquisa se ouver
        Set tipoMaterial = Nothing
        
        rs.MoveNext
    Wend
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    ' Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Retorna pesquisa
    Set listarTiposMateriais = listaTipoMaterial
    
    ' Libera espaço
    Set listaTipoMaterial = Nothing
End Function
