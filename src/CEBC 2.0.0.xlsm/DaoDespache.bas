Attribute VB_Name = "DaoDespache"
Option Explicit

Private listaDespaches As Collection
Private despache As objDespache
Private chapa As objChapa
Private tamanho As objTamanho

' Cadastra e edita objeto
Function cadastrarEEditar(despache As objDespache)
    ' String para consultas
    Dim rs As ADODB.Recordset
    Dim rsAuxiliar As ADODB.Recordset
    Dim rsIDCadastro As ADODB.Recordset
    Dim chapaLista As objBloco
    Dim listaIdBlocoAtualizarStatus As Collection
    Dim strSql As String ' String para consultas
    Dim campos() As String
    Dim valoresCampos As String
    Dim cadastro As Boolean
    Dim idCadastro As Integer
    Dim i As Long

    ' Seta true em cadastro
    cadastro = True
    
    'Faz a consulta para saber se o código do bloco já exite
    strSql = "SELECT * FROM Carregos_despaches" _
        & " WHERE id_carrego = " & despache.id & ";"
    
    ' Abrindo conexão com banco
    Call conctarBanco
    
    ' Criando objetos
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    Set rsAuxiliar = ObjectFactory.factoryRsAuxiliar(rsAuxiliar)
    Set listaIdBlocoAtualizarStatus = ObjectFactory.factoryLista(listaIdBlocoAtualizarStatus)
    
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
    
    CONEXAO_BD.BeginTrans
    
    ' Direciona para os comandos certos de cadastro ou edição
    If cadastro = True Then ' Se cadastro
    
        Set rsIDCadastro = ObjectFactory.factoryRsAuxiliar(rsIDCadastro)
        
        ReDim campos(1 To 5)
        campos(1) = "('" & despache.dataDespache & "', "
        campos(2) = "'" & despache.despachado & "', "
        campos(3) = "" & despache.getMotorista.id & ", "
        campos(4) = "" & despache.qtdChapas & ", "
        campos(5) = "" & despache.getDestino.id & ");"
        
        ' Concatenando os valores
        For i = 1 To 5
            valoresCampos = valoresCampos & campos(i)
        Next i
        
        ' Concatenando comando SQL e cadastrando bloco no banco de dados
        strSql = "INSERT INTO Carregos_despaches ( data_carrego, despachado, fk_motorista, qtd_chapas, fk_destino )" _
                        & " VALUES " & valoresCampos

        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
        
        ' Busca o id cadastrado
        strSql = "SELECT TOP 1 Id_Carrego FROM Carregos_despaches ORDER BY Id_Carrego DESC;"

        rsIDCadastro.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
        
        While Not rsIDCadastro.EOF
    
            despache.id = rsIDCadastro.Fields("Id_Carrego").Value
          
            rsIDCadastro.MoveNext
        Wend
        
        If despache.despachado = "SIM" Then
            ' Atualiza o estoque
            Call saidaEstoqueChapa(despache.getListaChapas)
            
            ' Atualiza status bloco
            Call saldoEstoqueStatus(despache.getListaChapas)
            
            ' Cadastra os materiais despachados
            Call cadastrarEditarMotoristaMateriais(despache)
        End If
    Else
        ' Se edição
        strSql = "UPDATE Carregos_despaches SET data_carrego = '" & despache.dataDespache & "', despachado = '" & despache.despachado & ", " _
                        & "fk_motorista = " & despache.getMotorista.id & ", fk_destino = " & despache.getDestino.id _
                        & " qtd_chapas = " & despache.qtdChapas & " WHERE Id_Destino = " & despache.id & ";"
            
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    End If
    
    CONEXAO_BD.CommitTrans
    
    Set rs = Nothing
    'Fechando conexão com banco
    Call fecharConexaoBanco
End Function

' Saida no estoque de chapas
Sub saidaEstoqueChapa(listaChapas As Collection)
          
    ' Variareis no medoto
    Dim rs As ADODB.Recordset
    Dim rsAuxiliar As ADODB.Recordset
    Dim strSql As String
    Dim qtnEstoque As String
    Dim m2Estoque As String
    Dim m2Atualizado As String
    Dim estoqueAtualizado As Integer
    Dim i As Integer
    
    For i = 1 To listaChapas.Count
        Set chapa = listaChapas.Item(i)
        Set tamanho = chapa.getTamanhos.Item(1)
        
        ' Consulta select
        strSql = "SELECT * FROM Tamanhos_Chapas WHERE id_tamanho = " & tamanho.id & ";"
        
        ' Criando e abrindo Recordset para consulta
        Set rs = ObjectFactory.factoryRsAuxiliar(rs)
        Set rsAuxiliar = ObjectFactory.factoryRsAuxiliar(rsAuxiliar)
        
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
            
        ' Laço para pegar estoque atual
        While Not rs.EOF
    
            qtnEstoque = rs.Fields("qtd_estoque").Value
            m2Estoque = rs.Fields("qtd_m2").Value
          
            rs.MoveNext
        Wend
        
        ' Faz a subtração do estoque
        estoqueAtualizado = qtnEstoque - tamanho.qtdEstoque
        m2Atualizado = M_METODOS_GLOBAL.subtracaoM2(m2Estoque, tamanho.qtdM2)
        
        'Edição do estoque de chapa
        strSql = "UPDATE Tamanhos_Chapas SET qtd_estoque = " & estoqueAtualizado _
                & ", qtd_m2 = '" & m2Atualizado _
                & "' WHERE id_tamanho = " & tamanho.id & ";"
        
        rsAuxiliar.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
        
        Set rs = Nothing
        Set rsAuxiliar = Nothing
        Set chapa = Nothing
        Set tamanho = Nothing
    Next i
End Sub

' Atualiza o Status do bloco para fechado se já estiver acabado a quantidade de chapas do bloco
Function saldoEstoqueStatus(listaChapas As Collection)

    'Variaveis do medoto
    Dim rs As ADODB.Recordset
    Dim rsAuxiliar As ADODB.Recordset
    Dim listaQtdCollection As Collection
    Dim chapaLista As objChapa
    Dim tamanhoChapa As objTamanho
    Dim qtd As Variant
    Dim qtdList As Variant
    Dim strSql As String
    Dim sqlAlteraStatusBloco As String
    Dim sqlSeletIdStatus As String
    Dim idStatus As Integer
    Dim qtdEstoque As Integer
    Dim i As Integer

    ' Inicializa as variaveis
    qtdEstoque = 0
    idStatus = 4 ' FECHADO
    
    For i = 1 To listaChapas.Count
        Set chapaLista = listaChapas.Item(i)
        Set tamanhoChapa = chapaLista.getTamanhos.Item(1)
        
        ' Consulta select
        strSql = "SELECT * FROM Tamanhos_Chapas WHERE id_tamanho = " & tamanhoChapa.id & ";"
        
        ' Criando e abrindo Recordset para consulta
        Set rs = ObjectFactory.factoryRsAuxiliar(rs)
        Set rsAuxiliar = ObjectFactory.factoryRsAuxiliar(rsAuxiliar)
        
        rsAuxiliar.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
        
        ' Captura a qtd de chapas
        While Not rsAuxiliar.EOF
    
            qtd = rsAuxiliar.Fields("qtd_estoque").Value
    
            rsAuxiliar.MoveNext
        Wend
        
        qtdEstoque = qtdEstoque + qtd
        
        ' Se estiver zerado o estoque muda o Status do bloco para "FECHADO"
        If qtdEstoque = 0 Then
    
            ' Consulta o estoque por pedreira
            strSql = "UPDATE Blocos SET Fk_Status = " _
                     & idStatus & " WHERE Id_Bloco = '" & chapaLista.getBloco.idSistema & "';"
    
            rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    
        End If
        
        Set rs = Nothing
        Set rsAuxiliar = Nothing
        Set chapa = Nothing
        Set tamanho = Nothing
    Next i
End Function

' Cadastra historico de carregos
Function cadastrarEditarMotoristaMateriais(despache As objDespache)

    ' Variaveis do medoto
    Dim rsSegundoCadastro As ADODB.Recordset
    Dim rsAuxiliar As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim chapaDespache As objChapa
    Dim tamanhoDespache As objTamanho
    Dim blocoDespache As objBloco
    Dim strSql As String
    Dim campos() As String
    Dim valoresCampos As String
    Dim dataFormatada As String
    Dim cadastro As Boolean
    Dim excluirCadastrio As Boolean
    Dim i As Integer
    Dim j As Integer

    dataFormatada = M_METODOS_GLOBAL.ConverterFormatoData(despache.dataDespache) ' Formata data
    cadastro = True
    excluirCadastrio = True
    
    strSql = "SELECT * FROM Motoristas_Materiais" _
        & " WHERE fk_Carrego = " & despache.id & " AND fk_motorista = " & despache.getMotorista.id & ";"
        
    Set rsAuxiliar = ObjectFactory.factoryRsAuxiliar(rsAuxiliar)
    
    ' Abrindo Recordset para consulta
    rsAuxiliar.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    ' Retorno da consulta
    While Not rsAuxiliar.EOF
        ' Seta false porquê vai ser uma edição
        cadastro = False
        
        rsAuxiliar.MoveNext
    Wend
    
    For i = 1 To despache.getListaChapas.Count
        
        Set rs = ObjectFactory.factoryRsAuxiliar(rs)
        Set chapaDespache = despache.getListaChapas.Item(i)
        Set tamanhoDespache = chapaDespache.getTamanhos.Item(1)
        
        ' Direciona para os comandos certos de cadastro ou edição
        If cadastro = True Then ' Se cadastro
            ReDim campos(1 To 13)
            campos(1) = "(" & despache.id & ", "
            campos(2) = despache.getMotorista.id & ", "
            campos(3) = "'" & chapaDespache.idSistema & "', "
            campos(4) = tamanhoDespache.id & ", "
            campos(5) = tamanhoDespache.estoque.id & ", "
            campos(6) = "'" & chapaDespache.getBloco.idSistema & "', "
            campos(7) = "'" & despache.dataDespache & "', " ' campos(7) = "#" & dataFormatada & "#, "
            campos(8) = "'" & chapaDespache.nomeMaterial & "', "
            campos(9) = tamanhoDespache.qtdEstoque & ", "
            campos(10) = "'" & tamanhoDespache.qtdM2 & "', "
            campos(11) = "'" & tamanhoDespache.compremento & "', "
            campos(12) = "'" & tamanhoDespache.altura & "', "
            campos(13) = "'" & tamanhoDespache.espessura & "');"


            ' Concatenando os valores
            For j = 1 To 13
                valoresCampos = valoresCampos & campos(j)
            Next j
            
            ' Concatenando comando SQL e cadastrando bloco no banco de dados
            strSql = "INSERT INTO Motoristas_Materiais ( [fk_carrego], [fk_motorista], [fk_chapa], [fk_tamanho], " _
                            & "[fk_estoque], [fk_bloco], [data_carregamento], [material], [qtd_estoque], [qtd_m2], " _
                            & "[comp], [alt], [esp] ) VALUES " & valoresCampos
            
            CONEXAO_BD.Execute strSql
            
'            strSql = "INSERT INTO Motoristas_Materiais ( [fk_carrego], [fk_motorista], [fk_chapa], [fk_tamanho], [fk_estoque], [fk_bloco], [data_carregamento], [material], [qtd_estoque], [qtd_m2], [comp], [alt], [esp] ) VALUES ( 54, 12, 'TESTE-TESTE-PO', 2, 1, 'TESTE-TESTE-BL', #2025-07-28#, 'TESTE POLIDO', 5, '30,0000', '3,0000', '2,0000', '02' );"
              
'            rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic

'            Debug.Print strSql
'            MsgBox strSql
            
'On Error GoTo TrataErro
'    If CONEXAO_BD Is Nothing Then
'        MsgBox "Conexão não está aberta.", vbExclamation
'        Exit Function
'    End If
'    CONEXAO_BD.Execute strSql
'    MsgBox "Registro inserido com sucesso!", vbInformation
'    Exit Function
'
'TrataErro:
'    MsgBox "Erro ao inserir dados: " & Err.Description, vbCritical
        Else
            ' Se edição
            If excluirCadastrio = True Then
                strSql = "DELETE FROM Motoristas_Materiais " _
                            & "WHERE fk_carrego = " & despache.id & "AND fk_motorista = " & despache.getMotorista.id & ";"
    
                rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
                
                excluirCadastrio = False
            End If
            
            Set rsSegundoCadastro = ObjectFactory.factoryRsAuxiliar(rsSegundoCadastro)
            
            ReDim campos(1 To 13)
            campos(1) = "(" & despache.id & ", "
            campos(2) = despache.getMotorista.id & ", "
            campos(3) = "'" & chapaDespache.idSistema & "', "
            campos(4) = tamanhoDespache.id & ", "
            campos(5) = tamanhoDespache.estoque.id & ", "
            campos(6) = "'" & chapaDespache.getBloco.idSistema & "', "
            campos(7) = "'" & despache.dataDespache & "', "
            campos(8) = "'" & chapaDespache.nomeMaterial & "', "
            campos(9) = tamanhoDespache.qtdEstoque & ", "
            campos(10) = "'" & tamanhoDespache.qtdM2 & "', "
            campos(11) = "'" & tamanhoDespache.compremento & "', "
            campos(12) = "'" & tamanhoDespache.altura & "', "
            campos(13) = "'" & tamanhoDespache.espessura & "');"
            
            
            ' Concatenando os valores
            For j = 1 To 13
                valoresCampos = valoresCampos & campos(j)
            Next j
            
            ' Concatenando comando SQL e cadastrando bloco no banco de dados
            strSql = "INSERT INTO Motoristas_Materiais ( [fk_carrego], [fk_motorista], [fk_chapa], [fk_tamanho], " _
                            & "[fk_estoque], [fk_bloco], [data_carregamento], [material], [qtd_estoque], [qtd_m2], " _
                            & "[comp], [alt], [esp] ) VALUES " & valoresCampos
            
            rsSegundoCadastro.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
            
            Set rsSegundoCadastro = Nothing
        End If
        
        Set rs = Nothing
        Set rsAuxiliar = Nothing
        Set chapaDespache = Nothing
        Set tamanhoDespache = Nothing
        valoresCampos = ""
    Next i
End Function

' Pesquisa objeto por id
Function pesquisarPorId(id As Integer) As objDespache
    
    ' String para consultas
    Dim strSql As String ' String para consultas
    Dim sqlSelectPesquisarPorId As String ' String para consultas
    Dim fkObject As String ' fk para consultas extras
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    
    ' String para consulta
    strSql = "SELECT * FROM Carregos_despaches WHERE Id_Carrego = " & id & ";"
    
    'Abrindo conexão com banco
    Call conctarBanco
    
    ' Criação e atribuição dos objeto
    Set despache = ObjectFactory.factoryDespache(despache)
    
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        
        despache.id = rs.Fields("Id_Carrego").Value
        despache.dataDespache = rs.Fields("Data_carrego").Value
        despache.despachado = rs.Fields("despachado").Value
        despache.qtdChapas = rs.Fields("qtd_chapas").Value
        
        
        'Atribuições dos objetos
        ' fk para consulta
        fkObject = rs.Fields("fk_motorista").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Motoristas WHERE Id_Motorista = " & fkObject & ";"
        ' Setando Objeto
        despache.setMotorista retornarObjeto(moto, sqlSelectPesquisarPorId, _
                            "id_motorista", "nome_motorista", "ativo")
        
        ' fk para consulta
        fkObject = rs.Fields("fk_destino").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Destinos WHERE id_destino = " & fkObject & ";"
        ' Setando Objeto
        despache.setDestino retornarObjeto(tipoPolimento, sqlSelectPesquisarPorId, _
                            "id_destino", "nome_destino", "ativo")
                            
        rs.MoveNext
    Wend
    
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    
    ' Fechar conexão com banco
    Call fecharConexaoBanco
    
    Set pesquisarPorId = despache
    
    Set despache = Nothing
End Function

' Pesquisar por motorista materiais
Function listaMateriaisMotoristasSalvos(despache As objDespache)


End Function

' Pesquisa objeto
Function listarDespachesSalvos()
    
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim strSql As String ' String para consultas
    Dim mensagem As String
    Dim qtdLista As Integer
    Dim id As Integer
    Dim i As Integer
    
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    Set listaDespaches = ObjectFactory.factoryLista(listaDespaches)
    
    ' String para consulta
    strSql = "SELECT * FROM Carregos_despaches WHERE despachado = 'NAO' ORDER BY Id_Carrego;"
    
    'Abrindo conexão com banco
    Call conctarBanco
    
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        
        id = rs.Fields("Id_Carrego").Value
                            
        listaDespaches.Add id
        
        rs.MoveNext
    Wend
    
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    
    ' Fechar conexão com banco
    Call fecharConexaoBanco
    
    mensagem = "Carregamentos salvos: "
    qtdLista = listaDespaches.Count
    
    For i = 1 To listaDespaches.Count
        If i < qtdLista Then
            mensagem = mensagem & listaDespaches.Item(i) & ", "
        Else
            mensagem = mensagem & listaDespaches.Item(i) & "."
        End If
    Next i
    
    ' Mensagem de retorno
    MsgBox mensagem, vbInformation, "CARREGAMENTOS SALVOS"
    
    Set listaDespaches = Nothing
End Function

' Metodo auxiliar para montar o objeto bloco
Function retornarObjeto(objeto As Object, sqlSelect As String, StringIdBanco As String, StringNomeBanco As String, _
                    StringAtivoBanco As String) As Object
                    
    ' Variaveis do metodo
    Dim rsAuxiliar As ADODB.Recordset ' Recordset para consulta
    
    ' Criando e abrindo Recordset para consulta
    Set rsAuxiliar = ObjectFactory.factoryRsAuxiliar(rsAuxiliar)
    ' Abrindo Recordset para consulta
    rsAuxiliar.Open sqlSelect, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    ' Retorno da consulta
    While Not rsAuxiliar.EOF
        ' Atribuição dos atributos
        objeto.id = rsAuxiliar.Fields(StringIdBanco).Value
        objeto.nome = rsAuxiliar.Fields(StringNomeBanco).Value
        objeto.nome = rsAuxiliar.Fields(StringAtivoBanco).Value
        
        rsAuxiliar.MoveNext
    Wend
    ' Libera recurso Recordset
    rsAuxiliar.Close
    Set rsAuxiliar = Nothing
    ' Retorno
    Set retornarObjeto = objeto
End Function
