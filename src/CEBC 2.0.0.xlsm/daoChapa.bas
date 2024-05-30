Attribute VB_Name = "daoChapa"
Option Explicit

Private listaChapas As Collection
Private chapa As objChapa
Private bloco As objBloco
Private tipoPolimento As objTipoPolimento
Private tamanho As objTamanho

' Cadastra e edita objeto
Function cadastrarEEditar(chapa As objChapa)
    ' String para consultas
    Dim rsChapa As ADODB.Recordset ' Recordset para consulta principal
    Dim rsAuxiliar As ADODB.Recordset ' Recordset para consulta
    Dim fkObject As Variant ' fk para consultas extras
    Dim strSql As String ' String para consultas
    Dim campos() As String
    Dim valoresCampos As String
    Dim cadastro As Boolean
    Dim i As Long
    Dim j As Long
    
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
    ' Fecha conexão do Recordset
    rsAuxiliar.Close
    
    ' Direciona para os comandos certos de cadastro ou edição
    If cadastro = True Then ' Se cadastro
        
        ' Realoca espaço da variavel
        ReDim campos(1 To 6)
        ' Colocando vingulas, Parenteses e  arpas simples os valores
        campos(1) = "('" & chapa.idSistema & "', "
        campos(2) = "'" & chapa.nomeMaterial & "', "
        campos(3) = "'" & chapa.valorTotal & "', "
        campos(4) = "'" & chapa.numeroBlocoPedreira & "', "
        campos(5) = chapa.tipoPolimento.id & ", "
        campos(6) = "'" & chapa.bloco.idSistema & "');"
        
        ' Concatenando os valores
        For i = 1 To 6
            valoresCampos = valoresCampos & campos(i)
        Next i
    
        ' Concatenando comando SQL e cadastrando bloco no banco de dados
        strSql = "INSERT INTO Chapas ( id_chapa, Descricao, valor_Total, numero_bloco_pedreira, fk_tipo_polimento, fk_bloco ) " _
                    & "VALUES " & valoresCamposChapa
        
        rsChapa.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
        
        ' Cadastra os tamanhos
        Call daoTamanho.cadastrarEEditar(chapa.tamanhos, False)
        
    Else ' Se edição
        
        ' Edição do bloco com serraria e polideira
        strSql = "UPDATE Chapas SET descricao = '" & chapa.nomeMaterial & "', " _
                            & "valor_Total = '" & chapa.valorTotal & "', " _
                            & "numero_bloco_pedreira = '" & chapa.numeroBlocoPedreira & "'," _
                            & "fk_tipo_polimento = '" & chapa.tipoPolimento.id & "',  " _
                            & "fk_Bloco = '" & chapa.bloco.idSistema & "' " _
                            & "WHERE id_Chapa = '" & chapa.idSistema & "';"
            
        rsChapa.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
        
        ' Edita os tamanhos
        Call daoTamanho.cadastrarEEditar(chapa.tamanhos, False)
    End If
    
    ' Libera espaço da memoria
    Set rsChapa = Nothing
    Set rsAuxiliar = Nothing
    'Fechando conexão com banco
    Call fecharConexaoBanco
End Function

' Exclui objeto
Function excluir(id As String)

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
    Set tipoPolimento = ObjectFactory.factoryTipoPolimento(tipoPolimento)
    
    'Abrindo conexão com banco
    Call conctarBanco
    ' String para consulta
    sqlSelectPesquisarPorId = "SELECT * FROM Chapas " & "WHERE id_Chapa = '" & id & "';"
    ' Criando e abrindo Recordset para consulta
    Set rsChapa = ObjectFactory.factoryRsAuxiliar(rsChapa)
    Set chapa = ObjectFactory.factoryChapa(chapa)
    ' Consulta banco
    rsChapa.Open sqlSelectPesquisarPorId, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    ' Retorno da consulta
    While Not rsChapa.EOF
        ' Atribuição dos atributos
        chapa.idSistema = rsChapa.Fields("id_Chapa").Value
        chapa.nomeMaterial = rsChapa.Fields("descricao").Value
        chapa.valorTotal = rsChapa.Fields("valor_total").Value
        chapa.numeroBlocoPedreira = rsChapa.Fields("numero_bloco_pedreira").Value
        
        'Atribuições dos objetos em bloco
        ' fk para consulta
        fkObject = rsChapa.Fields("fk_tipo_polimento").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Tipo_Polimento WHERE id_polimento = " & fkObject & ";"
        ' Setando Objeto
        chapa.setTipoPolimento retornarObjeto(tipoPolimento, sqlSelectPesquisarPorId, "id_polimento", "nome_polimento")
        
        ' fk para consulta
        fkObject = rsChapa.Fields("fk_bloco").Value
        ' Setando Objeto
        chapa.setBloco daoBloco.pesquisarPorId(fkObject, False)
        
        ' fk para consulta
        'chapa.setTamanhos daoTamanho.pesquisarPorIdChapa(chapa.idSistema, False)
        
        rsChapa.MoveNext
    Wend
    ' Retorno
    Set pesquisarPorId = chapa
    
    ' Libera espaço na memoria
    Set tipoPolimento = Nothing
    Set bloco = Nothing
    Set chapa = Nothing
End Function

' Pesquisa objeto por id
Function pesquisarPorIdPedreira(numeroBlocoPedreira As String) As Boolean
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
    sqlSelectPesquisarPorId = "SELECT * FROM Chapas " & "WHERE numero_bloco_pedreira = '" & numeroBlocoPedreira & "';"
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
    
    ' Fechar conexão com banco
    Call fecharConexaoBanco
    ' Retorno
    pesquisarPorIdPedreira = temCadastro
End Function
    
' Pesquisa objeto por id
Function pesquisarPorFKBloco(idBloco As String) As Collection
    'Metodos do metodo
    ' String para consultas
    Dim sqlSelectPesquisarPorId As String ' String para consultas
    Dim fkObject As String ' fk para consultas extras
    Dim rsChapa As ADODB.Recordset ' Recordset para consulta principal
    
    ' Criação da lista
    Set listaChapas = ObjectFactory.factoryLista(listaChapas)
    
    'Abrindo conexão com banco
    Call conctarBanco
    ' String para consulta
    sqlSelectPesquisarPorId = "SELECT * FROM Chapas " & "WHERE fk_bloco = '" & idBloco & "';"
    ' Criando e abrindo Recordset para consulta
    Set rsChapa = ObjectFactory.factoryRsAuxiliar(rsChapa)
    ' Consulta banco
    rsChapa.Open sqlSelectPesquisarPorId, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    ' Retorno da consulta
    While Not rsChapa.EOF
        ' Criação e atribuição dos objetos
        Set chapa = ObjectFactory.factoryChapa(chapa)
        Set bloco = ObjectFactory.factoryBloco(bloco)
        Set tipoPolimento = ObjectFactory.factoryTipoPolimento(tipoPolimento)
        
        ' Atribuição dos atributos
        chapa.idSistema = rsChapa.Fields("id_Chapa").Value
        chapa.nomeMaterial = rsChapa.Fields("descricao").Value
        chapa.valorTotal = rsChapa.Fields("valor_total").Value
        chapa.numeroBlocoPedreira = rsChapa.Fields("numero_bloco_pedreira").Value
        
        'Atribuições dos objetos em bloco
        ' fk para consulta
        fkObject = rsChapa.Fields("fk_tipo_polimento").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Tipo_Polimento WHERE id_polimento = " & fkObject & ";"
        ' Setando Objeto
        chapa.setTipoPolimento retornarObjeto(tipoPolimento, sqlSelectPesquisarPorId, "id_polimento", "nome_polimento")
        
        ' fk para consulta
        fkObject = rsChapa.Fields("fk_bloco").Value
        ' Setando Objeto
        chapa.setBloco daoBloco.pesquisarPorId(fkObject, False)
        
        ' fk para consulta
        chapa.setTamanhos daoTamanho.pesquisarPorIdChapa(chapa.idSistema, False)
        
        ' Seta chapa na lista
        listaChapas.Add chapa
        
        ' Libera espaço na memoria
        Set tipoPolimento = Nothing
        Set bloco = Nothing
        Set chapa = Nothing
    
        rsChapa.MoveNext
    Wend
    
    ' Retorno
    Set pesquisarPorFKBloco = listaChapas
    ' Libera espaço na memoria
    Set listaChapas = Nothing
End Function

' Pesquisa objeto por uma lista de ids do bloco
Function pesquisarPorListaIdsPedreira(listaIdsParaPesquisa As Collection) As Collection
    'Metodos do metodo
    ' String para consultas
    Dim sqlSelectPesquisarPorId As String ' String para consultas
    Dim fkObject As String ' fk para consultas extras
    Dim rsChapa As ADODB.Recordset ' Recordset para consulta principal
    Dim listaChapasAvulsas As Collection
    Dim numeroBlocoPedreira As String
    Dim idLista As Variant
    
    ' Criação da lista para adição e retorno
    Set listaChapasAvulsas = ObjectFactory.factoryLista(listaChapasAvulsas)
    
    'Abrindo conexão com banco
    Call conctarBanco
    
    For Each idLista In listaIdsParaPesquisa
        ' Seta id para pesquisa
        numeroBlocoPedreira = idLista
        
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Chapas " & "WHERE numero_bloco_pedreira = '" & numeroBlocoPedreira & "' ORDER BY Descricao;"
        ' Criando e abrindo Recordset para consulta
        Set rsChapa = ObjectFactory.factoryRsAuxiliar(rsChapa)
        ' Consulta banco
        rsChapa.Open sqlSelectPesquisarPorId, CONEXAO_BD, adOpenKeyset, adLockReadOnly
        ' Retorno da consulta
        While Not rsChapa.EOF
            ' Criação e atribuição dos objetos
            Set chapa = ObjectFactory.factoryChapa(chapa)
            Set bloco = ObjectFactory.factoryBloco(bloco)
            Set tipoPolimento = ObjectFactory.factoryTipoPolimento(tipoPolimento)
        
            ' Atribuição dos atributos
            chapa.idSistema = rsChapa.Fields("id_Chapa").Value
            chapa.nomeMaterial = rsChapa.Fields("descricao").Value
            chapa.valorTotal = rsChapa.Fields("valor_total").Value
            chapa.numeroBlocoPedreira = rsChapa.Fields("numero_bloco_pedreira").Value
            
            'Atribuições dos objetos em bloco
            ' fk para consulta
            fkObject = rsChapa.Fields("fk_tipo_polimento").Value
            ' String para consulta
            sqlSelectPesquisarPorId = "SELECT * FROM Tipo_Polimento WHERE id_polimento = " & fkObject & ";"
            ' Setando Objeto
            chapa.setTipoPolimento retornarObjeto(tipoPolimento, sqlSelectPesquisarPorId, "id_polimento", "nome_polimento")
            
            ' fk para consulta
            fkObject = rsChapa.Fields("fk_bloco").Value
            ' Setando Objeto
            chapa.setBloco daoBloco.pesquisarPorId(fkObject, False)
            
            ' fk para consulta
            chapa.setTamanhos daoTamanho.pesquisarPorIdChapa(chapa.idSistema, False)
            
            ' Adciona a chapa na lista
            listaChapasAvulsas.Add chapa
            
            ' Libera espaço para da momeria
            Set chapa = Nothing
            Set bloco = Nothing
            Set tipoPolimento = Nothing
            
            rsChapa.MoveNext
        Wend
    Next idLista
    
    ' Libera recurso Recordset
    rsChapa.Close
    Set rsChapa = Nothing
    
    ' Fechar conexão com banco
    Call fecharConexaoBanco
    ' Retorno
    Set pesquisarPorListaIdsPedreira = listaChapasAvulsas
    ' Libera espaço
    Set listaChapasAvulsas = Nothing
End Function

' Pesquisa objeto por nome
Function pesquisarPorNome()

End Function

' Pesquisa objeto
Function listarChapasFilter(nomeMaterial As String, numeroBlocoPedreira As String, idBlocoSistema As String, _
                    nomePolideira As String, nomeTipoPolimento As String, estoqueZero As String)
    
    ' String para consultas
    Dim filterListaChapa As Collection
    Dim filterlistaTamanhos As Collection
    Dim tamanho As objTamanho
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
    Dim j As Long
    
    ' String SQL para a consulta
    strSql = "SELECT * FROM Chapas"
    strLike = ""
    strWhere = ""
    strOrderBY = "ORDER BY descricao;"
    
    ' Construindo a cláusula LIKE
    If nomeMaterial <> "" Then
        strLike = "descricao LIKE '%%" & nomeMaterial & "%%'"
    End If
    
    ' Abrindo conexão com banco para pesquisar as Chapas
    Call conctarBanco

    ' Construindo a cláusula WHERE baseada nos filtros selecionados
    If estoqueZero = "NÃO" Then
        strWhere = "estoque_zero = 'NAO'"
    End If

    If numeroBlocoPedreira <> "" Then
        If strWhere <> "" Then
            strWhere = strWhere & " AND "
        End If
        strWhere = strWhere & "numero_bloco_pedreira = '" & numeroBlocoPedreira & "'"
    End If

    If idBlocoSistema <> "" Then
        If strWhere <> "" Then
            strWhere = strWhere & " AND "
        End If
        strWhere = strWhere & "fk_bloco = '" & idBlocoSistema & "'"
    End If

    If nomeTipoPolimento <> "" Then
        idTipoPolimento = retornarIdObjeto( _
                        "SELECT * FROM Tipo_Polimento WHERE Nome_Polimento = '" & nomeTipoPolimento & "';", "Id_Polimento")
        If strWhere <> "" Then
            strWhere = strWhere & " AND "
        End If
        strWhere = strWhere & "fk_tipo_polimento = " & idTipoPolimento
    End If

    ' Adicionar a cláusula WHERE à consulta
    If strWhere <> "" Then
        If strLike <> "" Then
                strSql = strSql & " WHERE " & strWhere & " AND " & strLike & " " & strOrderBY
        Else
            strSql = strSql & " WHERE " & strWhere & " " & strOrderBY
        End If
    Else
        If strLike <> "" Then
                strSql = strSql & " WHERE " & strLike & " " & strOrderBY
        Else
            strSql = strSql & " " & strOrderBY
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
        ' Criação para atribuição
        Set bloco = ObjectFactory.factoryBloco(bloco)
        Set tipoPolimento = ObjectFactory.factoryTipoPolimento(tipoPolimento)
        Set tamanho = ObjectFactory.factoryTamanho(tamanho)
        
        ' Atribuição dos atributos
        chapa.idSistema = rsChapa.Fields("id_Chapa").Value
        chapa.nomeMaterial = rsChapa.Fields("descricao").Value
        chapa.valorTotal = rsChapa.Fields("valor_total").Value
        chapa.numeroBlocoPedreira = rsChapa.Fields("numero_bloco_pedreira").Value
        
        'Atribuições dos objetos em bloco
        ' fk para consulta
        fkObject = rsChapa.Fields("fk_tipo_polimento").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Tipo_Polimento WHERE id_polimento = " & fkObject & ";"
        ' Setando Objeto
        chapa.setTipoPolimento retornarObjeto(tipoPolimento, sqlSelectPesquisarPorId, "id_polimento", "nome_polimento")
        
        ' fk para consulta
        fkObject = rsChapa.Fields("Fk_bloco").Value
        ' Setando Objeto
        chapa.setBloco daoBloco.pesquisarPorId(fkObject, False)
        
        ' fk para consulta
        chapa.setTamanhos daoTamanho.pesquisarPorIdChapa(chapa.idSistema, False)
        
        For Each tamanho In chapa.getTamanhos
            tamanho.setChapa chapa
        Next tamanho
        
        ' Adiciona na lista
        listaChapas.Add chapa
        
        ' Libera espaço para nova pesquisa se ouver
        Set tipoPolimento = Nothing
        Set bloco = Nothing
        Set chapa = Nothing
        Set tamanho = Nothing
        
        rsChapa.MoveNext
    Wend
    
    ' Libera recurso Recordset
    rsChapa.Close
    Set rsChapa = Nothing
    ' Fechar conexão com banco
    Call fecharConexaoBanco
    
    If nomePolideira <> "" Then
        ' Cria a lista e objetos
        Set filterListaChapa = ObjectFactory.factoryLista(filterListaChapa)
        Set filterlistaTamanhos = ObjectFactory.factoryLista(filterlistaTamanhos)
        Set chapa = ObjectFactory.factoryChapa(chapa)
        Set tamanho = ObjectFactory.factoryTamanho(tamanho)
        
        ' Laço para filtro nas chapas
        For i = 1 To listaChapas.Count
            ' Seta chapa
            Set chapa = listaChapas.Item(i)
            ' Laço nos tamanhos da chapa
            For j = 1 To chapa.tamanhos.Count
                ' Seta tamanho
                Set tamanho = chapa.tamanhos.Item(j)
                
                ' Filtra o nome da Polideira
                If nomePolideira = tamanho.polideira.nome Then
                    ' Filtra a qtd estoque
                    If estoqueZero = "NÃO" Then
                        If tamanho.qtdEstoque > 0 Then
                            filterlistaTamanhos.Add tamanho
                        End If
                    Else
                        filterlistaTamanhos.Add tamanho
                    End If
                End If
                Set tamanho = Nothing
            Next j
            
            ' Troca lista com tamanhos pela lista filtrada se tiver dados
            ' Verifica se tem algum dado a pesquisa
            If filterlistaTamanhos.Count > 0 Then
                chapa.setTamanhos filterlistaTamanhos
            End If
            ' Adiciona na lista
            filterListaChapa.Add chapa
            
            ' Libera espaço na memoria
            Set chapa = Nothing
        Next i
    End If

    ' Retorna pesquisa
    If nomePolideira <> "" Then
        Set listarChapasFilter = filterListaChapa
    Else
        Set listarChapasFilter = listaChapas
    End If
    
    
    ' Libera espaço
    Set listaChapas = Nothing
    
    If nomePolideira <> "" Then
        ' Libera espaço
        Set filterListaChapa = Nothing
        Set chapa = Nothing
        Set tamanho = Nothing
    End If
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

