Attribute VB_Name = "daoChapa"
Option Explicit

Private listaChapas As Collection
Private chapa As objChapa
Private bloco As objBloco
Private polideira As objPolideira
Private tipoPolimento As objTipoPolimento
Private estoque As objEstoqueChapa

' Cadastra e edita objeto
Function cadastrarEEditar(chapa As objChapa)
    ' String para consultas
    Dim rsChapa As ADODB.Recordset ' Recordset para consulta principal
    Dim rsAuxiliar As ADODB.Recordset ' Recordset para consulta
    Dim fkObject As Variant ' fk para consultas extras
    Dim strSql As String ' String para consultas
    Dim campos() As String
    Dim valoresCamposChapa As String
    Dim cadastro As Boolean
    Dim i As Long
    
    ' Seta true em cadastro
    cadastro = True
    
    ' Faz a consulta para saber se o código do bloco já exite
    strSql = "SELECT * FROM Chapas WHERE Id_Chapa = '" & chapa.idSistema & "';"
    
    ' Abrindo conexão com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rsChapa = ObjectFactory.factoryRsAuxiliar(rsChapa)
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
        
        ' Realoca espaço da variavel
        ReDim campos(1 To 13)
        ' Colocando vingulas, Parenteses e  arpas simples os valores
        campos(1) = "('" & chapa.idSistema & "', "
        campos(2) = "'" & chapa.nomeMaterial & "', "
        campos(3) = "'" & chapa.custoPolimento & "', "
        campos(4) = "'" & chapa.custoTotal & "', "
        campos(5) = "'" & chapa.qtdEstoque & "', "
        campos(6) = "'" & chapa.qtdM2Bruto & "', "
        campos(7) = "'" & chapa.compBruto & "', "
        campos(8) = "'" & chapa.altBruto & "', "
        campos(9) = "'" & chapa.idPedreira & "', "
        campos(10) = "'" & chapa.tipoPolimento.id & "', "
        campos(11) = "'" & chapa.estoque.id & "', "
        campos(12) = "'" & chapa.polideira.id & "', "
        campos(13) = "'" & chapa.bloco.idSistema & "');"
        
        ' Concatenando os valores
        For i = 1 To 13
            valoresCamposChapa = valoresCamposChapa & campos(i)
        Next i
    
        ' Concatenando comando SQL e cadastrando bloco no banco de dados
        strSql = "INSERT INTO Chapas ( Id_Chapa, Descricao, Custo_Polimento, Custo_Total, Qtd_Estoque, Qtd_Bruto_M2, " _
                    & "Comp_Bruto, Alt_Bruto, Id_bloco_Pedreira, Fk_Polimento, Fk_Estoque, Fk_Polidoria, Fk_Bloco ) " _
                    & "VALUES " & valoresCamposChapa
        
        rsChapa.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    Else ' Se edição
        
        ' Edição do bloco com serraria e polideira
        strSql = "UPDATE Chapas SET Descricao = '" & chapa.nomeMaterial & "', " _
            & "Custo_Polimento = '" & chapa.custoPolimento & "', Custo_Total = '" & chapa.custoTotal & "', " _
            & "Qtd_Estoque = '" & chapa.qtdEstoque & "', Qtd_Bruto_M2 = '" & chapa.qtdM2Bruto & "', " _
            & "Comp_Bruto = '" & chapa.compBruto & "', Alt_Bruto = '" & chapa.altBruto & "', " _
            & "Id_bloco_Pedreira = '" & chapa.idPedreira & "', Fk_Polimento = '" & chapa.tipoPolimento.id & "', " _
            & "Fk_Estoque = '" & chapa.estoque.id & "', Fk_Polidoria = '" & chapa.polideira.id & "', " _
            & "Fk_Bloco = '" & chapa.bloco.idSistema & "' WHERE Id_Chapa = '" & chapa.idSistema & "';"
            
        rsChapa.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    End If
    ' Libera espaço da memoria
    Set rsChapa = Nothing
    'Fechando conexão com banco
    Call fecharConexaoBanco
End Function

' Exclui objeto
Function excluir()

End Function

' Pesquisa objeto por id
Function pesquisarPorId(id As String) As objChapa
    'Metodos do metodo
    ' String para consultas
    Dim sqlSelectPesquisarPorId As String ' String para consultas
    Dim fkObject As String ' fk para consultas extras
    Dim rsChapa As ADODB.Recordset ' Recordset para consulta principal
    
    ' Criação e atribuição dos objetos
    Set bloco = ObjectFactory.factoryBloco(bloco)
    Set polideira = ObjectFactory.factoryPolideira(polideira)
    Set tipoPolimento = ObjectFactory.factoryTipoPolimento(tipoPolimento)
    Set estoque = ObjectFactory.factoryEstoque(estoque)
    
    'Abrindo conexão com banco
    Call conctarBanco
    ' String para consulta
    sqlSelectPesquisarPorId = "SELECT * FROM Chapas " & "WHERE Id_Chapa = '" & id & "';"
    ' Criando e abrindo Recordset para consulta
    Set rsChapa = ObjectFactory.factoryRsAuxiliar(rsChapa)
    ' Consulta banco
    rsChapa.Open sqlSelectPesquisarPorId, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    ' Retorno da consulta
    While Not rsChapa.EOF
        ' Atribuição dos atributos
        chapa.idSistema = rsChapa.Fields("Id_Chapa").Value
        chapa.nomeMaterial = rsChapa.Fields("Descricao").Value
        chapa.custoPolimento = rsChapa.Fields("Custo_Polimento").Value
        chapa.custoTotal = rsChapa.Fields("Custo_Total").Value
        chapa.qtdEstoque = rsChapa.Fields("Qtd_Estoque").Value
        chapa.qtdM2Bruto = rsChapa.Fields("Qtd_Bruto_M2").Value
        chapa.compBruto = rsChapa.Fields("Comp_Bruto").Value
        chapa.altBruto = rsChapa.Fields("Alt_Bruto").Value
        chapa.idPedreira = rsChapa.Fields("Id_bloco_Pedreira").Value
        
        'Atribuições dos objetos em bloco
        ' fk para consulta
        fkObject = rsChapa.Fields("Fk_Polimento").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Tipo_Polimento WHERE Id_Polimento = " & fkObject & ";"
        ' Setando Objeto
        chapa.setTipoPolimento retornarObjeto(tipoPolimento, sqlSelectPesquisarPorId, "Id_Polimento", "Nome_Polimento")
        
        ' fk para consulta
        fkObject = rsChapa.Fields("Fk_Estoque").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Estoque_chapas WHERE Id_Estoque = " & fkObject & ";"
        ' Setando Objeto
        chapa.setEstoque retornarObjeto(estoque, sqlSelectPesquisarPorId, "Id_Estoque", "Nome_Empresa")
        
        ' fk para consulta
        fkObject = rsChapa.Fields("Fk_Estoque").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Polideiras WHERE Id_Estoque = " & fkObject & ";"
        ' Setando Objeto
        chapa.setPolideira retornarObjeto(estoque, sqlSelectPesquisarPorId, "Id_Polidoria", "Nome_Polidoria")
        
        ' fk para consulta
        fkObject = rsChapa.Fields("Fk_Bloco").Value
        ' Setando Objeto
        chapa.setBloco daoBloco.pesquisarPorId(fkObject)
        
        ' fk para consulta
        chapa.setTamanhos daoTamanho.pesquisarPorIdChapa(chapa.idSistema)
        
        rsChapa.MoveNext
    Wend
    ' Retorno
    Set pesquisarPorId = chapa
    
    ' Libera espaço na memoria
    Set tipoPolimento = Nothing
    Set polideira = Nothing
    Set estoque = Nothing
    Set bloco = Nothing
    Set chapa = Nothing
End Function

' Pesquisa objeto por id
Function pesquisarPorIdPedreira(idPedreira As String) As Boolean
    'Metodos do metodo
    ' String para consultas
    Dim sqlSelectPesquisarPorId As String ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim temCadastro As Boolean
    
    ' Seta não para verificação de cadastro
    temCadastro = False
    
    'Abrindo conexão com banco
    Call conctarBanco
    ' String para consulta
    sqlSelectPesquisarPorId = "SELECT * FROM Chapas " & "WHERE Id_bloco_Pedreira = '" & idPedreira & "';"
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    ' Consulta banco
    rs.Open sqlSelectPesquisarPorId, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    ' Retorno da consulta
    While Not rs.EOF
        ' Seta true em temCadastro porquê achou um cadastro
        temCadastro = True
        
        rs.MoveNext
    Wend
    
    ' Retorno
    pesquisarPorIdPedreira = temCadastro
End Function

' Pesquisa objeto por nome
Function pesquisarPorNome()

End Function

' Pesquisa objeto
Function listarChapasFilter(nomeMaterial As String, numeroBlocoPedreira As String, idBlocoSistema As String, _
                    polideira As String, tipoPolimento As String, estoqueZero As String)
    
    ' String para consultas
    Dim sqlSelectPesquisarPorId As String ' String para consultas auxiliar
    Dim strSql As String ' String para consultas
    Dim fkObject As Variant ' fk para consultas extras
    Dim rsChapa As ADODB.Recordset ' Recordset para consulta principal
    Dim strLike As String
    Dim strWhere As String
    Dim strOrderBY As String
    Dim idPolideira As String
    Dim idTipoPolimento As String
    Dim i As Long
    
        ' String SQL para a consulta
    strSql = "SELECT * FROM Chapas"
    strLike = ""
    strWhere = ""
    strOrderBY = " ORDER BY Descricao;"
    
        ' Construindo a cláusula LIKE
'    If descricaoBloco <> "" Then
'        strLike = "Descricao LIKE '*" & nomeMaterial & "*'"
'
'    End If
    
    ' Abrindo conexão com banco para pesquisar as Chapas
    Call conctarBanco

    ' Construindo a cláusula WHERE baseada nos filtros selecionados
    If estoqueZero = "NÃO" Then
        strWhere = "Quantidade_Estoque > 0"
    End If
    
    If numeroBlocoPedreira <> "" Then
        If strWhere <> "" Then
            strWhere = strWhere & " AND "
        End If
        strWhere = "Id_bloco_Pedreira = '" & numeroBlocoPedreira & "'"
    End If
    
    If idBlocoSistema <> "" Then
        If strWhere <> "" Then
            strWhere = strWhere & " AND "
        End If
        strWhere = strWhere & "Id_Chapa = '" & idBlocoSistema & "'"
    End If
    
    If polideira <> "" Then
        idPolideira = retornarIdObjeto( _
                        "SELECT * FROM Polideiras WHERE Nome_Polidoria = '" & polideira & "';", "Id_Polidoria")
        If strWhere <> "" Then
            strWhere = strWhere & " AND "
        End If
        strWhere = strWhere & "Fk_Polidoria = " & idPolideira
    End If

    If tipoPolimento <> "" Then
        idTipoPolimento = retornarIdObjeto( _
                        "SELECT * FROM Tipo_Polimento WHERE Nome_Polimento = '" & tipoPolimento & "';", "Id_Polimento")
        If strWhere <> "" Then
            strWhere = strWhere & " AND "
        End If
        strWhere = strWhere & "Fk_Polimento = " & idTipoPolimento
    End If
    
    ' Adicionar a cláusula WHERE à consulta
    If strWhere <> "" Then
        If strLike <> "" Then
                strSql = strSql & " WHERE " & strWhere & " AND " & strLike & strOrderBY
        Else
            strSql = strSql & " " & strWhere & strOrderBY
        End If
    Else
        If strLike <> "" Then
                strSql = strSql & " WHERE " & strLike & strOrderBY
        Else
            strSql = strSql & strOrderBY
        End If
    End If
    
    ' Criação e atribuição dos objeto
    Set listaChapas = ObjectFactory.factoryLista(listaChapas)
    
    ' Criando e abrindo Recordset para consulta
    Set rsChapa = ObjectFactory.factoryRsAuxiliar(rsChapa)
    ' Consulta banco
    rsChapa.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    ' Retorno da consulta
    While Not rsChapa.EOF
        ' Atribuição para comparação
        Set chapa = ObjectFactory.factoryChapa(chapa)
        chapa.nomeMaterial = rsChapa.Fields("Descricao").Value
        'Analisa se tem a descrição no texto digitado
        If InStr(1, chapa.nomeMaterial, nomeMaterial, vbTextCompare) > 0 Then
            ' Criação para atribuição
            Set bloco = ObjectFactory.factoryBloco(bloco)
            Set polideira = ObjectFactory.factoryPolideira(polideira)
            Set tipoPolimento = ObjectFactory.factoryTipoPolimento(tipoPolimento)
            Set estoque = ObjectFactory.factoryEstoqueChapas(estoque)
           
            ' Atribuição dos atributos
            chapa.idSistema = rsChapa.Fields("Id_Chapa").Value
            chapa.nomeMaterial = rsChapa.Fields("Descricao").Value
            chapa.custoPolimento = rsChapa.Fields("Custo_Polimento").Value
            chapa.custoTotal = rsChapa.Fields("Custo_Total").Value
            chapa.qtdEstoque = rsChapa.Fields("Qtd_Estoque").Value
            chapa.qtdM2Bruto = rsChapa.Fields("Qtd_Bruto_M2").Value
            chapa.compBruto = rsChapa.Fields("Id_Chapa").Value
            chapa.altBruto = rsChapa.Fields("Id_Chapa").Value
            chapa.idPedreira = rsChapa.Fields("Id_Chapa").Value
            
            ' Atribuições dos objetos em chapa
            ' fk para consulta
            fkObject = rsChapa.Fields("Fk_Polidoria").Value
            ' String para consulta
            sqlSelectPesquisarPorId = "SELECT * FROM Polideiras WHERE Id_Polidoria = " & fkObject & ";"
            ' Setando Objeto
            chapa.setPolideira retornarObjeto(polideira, sqlSelectPesquisarPorId, "Id_Polidoria", "Nome_Polidoria")
            
            ' fk para consulta
            fkObject = rsChapa.Fields("Fk_Polimento").Value
            ' String para consulta
            sqlSelectPesquisarPorId = "SELECT * FROM Tipo_Polimento WHERE Id_Polimento = " & fkObject & ";"
            ' Setando Objeto
            chapa.setTipoPolimento retornarObjeto(tipoPolimento, sqlSelectPesquisarPorId, "Id_Polimento", "Nome_Polimento")
        
            ' fk para consulta
            fkObject = rsChapa.Fields("Fk_Estoque").Value
            ' String para consulta
            sqlSelectPesquisarPorId = "SELECT * FROM Estoque_chapas WHERE Id_Estoque = " & fkObject & ";"
            ' Setando Objeto
            chapa.setEstoque retornarObjeto(estoque, sqlSelectPesquisarPorId, "Id_Estoque", "Nome_Empresa")
        
            ' fk para consulta
            fkObject = rsChapa.Fields("Fk_Bloco").Value
            ' Setando Objeto
            chapa.setBloco daoBloco.pesquisarPorId(fkObject)
            
            ' fk para consulta
            chapa.setTamanhos daoTamanho.pesquisarPorIdChapa(chapa.idSistema)

            ' Adiciona na lista
            listaChapas.Add chapa
            
            ' Libera espaço para nova pesquisa se ouver
            Set tipoPolimento = Nothing
            Set polideira = Nothing
            Set estoque = Nothing
            Set bloco = Nothing
            Set chapa = Nothing
            
            rsChapa.MoveNext
        Else
            rsChapa.MoveNext
        End If
    Wend
    
    ' Libera recurso Recordset
    rsChapa.Close
    Set rsChapa = Nothing
    ' Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Retorna pesquisa
    Set listarChapasFilter = listaChapas
    
    ' Libera espaço
    Set listaChapas = Nothing
End Function

' Metodo auxiliar para montar o objeto bloco
Function retornarObjeto(objeto As Object, sqlSelect As String, StringIdBanco As String, StringNomeBanco As String) As Object
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
        
        rsAuxiliar.MoveNext
    Wend
    ' Libera recurso Recordset
    rsAuxiliar.Close
    Set rsAuxiliar = Nothing
    ' Retorno
    Set retornarObjeto = objeto
End Function
