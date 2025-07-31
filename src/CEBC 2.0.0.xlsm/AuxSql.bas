Attribute VB_Name = "AuxSql"
Option Explicit

' Cadastra os estoque com nomes das pedreiras
Function cadastrarEstoquePedreiras()

    'Metodos do metodo
    '--------------------------------------------------------------
    Dim nomePedreiraCollection As Collection
    Dim rs As ADODB.Recordset
    Dim sqlSelectPedreiras As String
    Dim sqlInsertEstoque As String
    Dim nomeEstoque As Variant
    Dim i As Integer
    
    'Criando Coleção para manipulação
    '--------------------------------------------------------------
    Set nomePedreiraCollection = New Collection
    Set rs = New ADODB.Recordset
    
    'Consulta select Pedreiras
    '--------------------------------------------------------------
    sqlSelectPedreiras = "SELECT * FROM Pedreiras ORDER BY Nome_Pedreira;"
    
    'Abrindo conexão com banco para pesquisar as Pedreiras
    Call conctarBanco
    
    rs.Open sqlSelectPedreiras, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    'Adicionando nome na coleção
    While Not rs.EOF
    
        nomePedreiraCollection.Add rs.Fields("Nome_Pedreira").Value
        
        rs.MoveNext
    Wend
        
    'Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Cadastras os estoque para os blocos
    '-----------------------------------------------------------------------
    For i = 1 To nomePedreiraCollection.Count
    
        nomeEstoque = nomePedreiraCollection.Item(i)
        
        If nomeEstoque = "AVULSO" Then

        ElseIf nomeEstoque = "IMPORTADO" Then
        
        Else
            'Comando para estoque
            sqlInsertEstoque = "INSERT INTO Estoque_chapas ( Nome_Empresa )VALUES ('" & nomeEstoque & "');"
            
            'Abrindo conexão com banco para cadastro bloco
            Call conctarBanco
            
            rs.Open sqlInsertEstoque, CONEXAO_BD, adOpenKeyset, adLockPessimistic
            
            'Fechando conexão com banco
            Call fecharConexaoBanco
        End If
    Next i
End Function

' Coloca sim ativo em todos os blocos
Function adicionarAtributoBlocos()

    'Metodos do metodo
    '--------------------------------------------------------------
    Dim rs As ADODB.Recordset
    Dim sqlInsertSim As String
    Dim id As Variant
    
    'Criando Coleção para manipulação
    '--------------------------------------------------------------
    Set rs = New ADODB.Recordset
    
    ' Comando para edicao
    sqlInsertSim = "UPDATE Blocos SET Ativo = 'SIM';"
    
    'Abrindo conexão com banco para cadastro bloco
    Call conctarBanco
    
    rs.Open sqlInsertSim, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    
    'Fechando conexão com banco
    Call fecharConexaoBanco
End Function
