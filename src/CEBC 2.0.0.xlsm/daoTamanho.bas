Attribute VB_Name = "daoTamanho"
Option Explicit

Private listaTamanhos As Collection
Private tipoMaterial As objTipoMaterial
Private tamanho As objTamanho
Private polideira As objPolideira
Private estoque As objEstoqueChapa
Private chapa As objChapa

' Cadastra e edita objeto
Function cadastrarEEditar(lista As Collection, fecharConexao As Boolean)
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim rsAuxiliar As ADODB.Recordset ' Recordset para consulta
    Dim strSql As String ' String para consultas
    Dim campos() As String
    Dim valoresCampos As String
    Dim cadastro As Boolean
    Dim i As Long
    Dim j As Long

    ' Seta true em cadastro
    cadastro = True
    
    ' Abrindo conexão com banco
    Call conctarBanco
    
    ' Loop através dos itens da coleção
    For i = 1 To lista.Count
    
            ' Criando e abrindo Recordset para consulta
        Set rs = ObjectFactory.factoryRsAuxiliar(rs)
        Set rsAuxiliar = ObjectFactory.factoryRsAuxiliar(rsAuxiliar)
        
        ' Seta o ojeto
        Set tamanho = lista(i)
        
        ' String para consulta
        strSql = "SELECT * FROM Tamanhos_Chapas " & "WHERE id_tamanho = " & tamanho.id & "" _
                & " AND fK_chapa = '" & tamanho.chapa.idSistema & "';"
        
        ' Consulta banco
        rsAuxiliar.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
        
        ' Retorno da consulta
        While Not rsAuxiliar.EOF
            ' Irá ser um edição
            cadastro = False
            rsAuxiliar.MoveNext
        Wend
        
        ' Direciona para os comandos certos de cadastro ou edição
        If cadastro = True Then ' Se cadastro
            ' Realoca espaço da variavel
            ReDim campos(1 To 10)
            ' Colocando vingulas, Parenteses e  arpas simples os valores
            campos(1) = "('" & tamanho.chapa.idSistema & "', "
            campos(2) = "'" & tamanho.compremento & "', "
            campos(3) = "'" & tamanho.altura & "', "
            campos(4) = "'" & tamanho.qtdEstoque & "', "
            campos(5) = "'" & tamanho.qtdM2 & "', "
            campos(6) = "'" & tamanho.valorPolimento & "', "
            campos(7) = "'" & tamanho.espessura & "', "
            campos(8) = tamanho.tipoMaterial.id & ", "
            campos(9) = tamanho.polideira.id & ", "
            campos(10) = tamanho.estoque.id & ");"
            
            ' Concatenando os valores
            For j = 1 To 10
                valoresCampos = valoresCampos & campos(j)
            Next j
            
            ' Concatenando comando SQL e cadastrando bloco no banco de dados
            strSql = "INSERT INTO Tamanhos_Chapas ( fK_chapa, comp_chapa, alt_chapa, qtd_estoque, qtd_m2, " _
                        & "valor_polimento, esp_chapa, fk_tipo_material, fk_polidoria, fk_estoque ) " _
                        & "VALUES " & valoresCampos
                    
            rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
            
            ' Limpa variavel para proximo cadastro
            valoresCampos = ""
        Else
            ' Se edição
            
            strSql = "UPDATE Tamanhos_Chapas SET fK_chapa = '" & tamanho.chapa.idSistema & "', " _
                & "comp_chapa = '" & tamanho.compremento & "', alt_chapa = '" & tamanho.altura & "', " _
                & "qtd_estoque = " & tamanho.qtdEstoque & ", qtd_m2 = '" & tamanho.qtdM2 & "', " _
                & "valor_Polimento = " & tamanho.valorPolimento & ", esp_chapa = '" & tamanho.espessura & "', " _
                & "fk_Tipo_Material = " & tamanho.tipoMaterial.id & ", fk_polidoria = " & tamanho.polideira.id & ", " _
                & "fk_estoque = " & tamanho.estoque.id & " WHERE id_tamanho = " & tamanho.id & ";"
                        
            ' Abrindo Recordset para consulta
            rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
            ' Retorna o valor
            cadastro = True
        End If
        
            ' Libera recurso Recordset
            rsAuxiliar.Close
            Set rsAuxiliar = Nothing
            Set rs = Nothing
    Next i

    ' Libara espaço na memoria
    Set tamanho = Nothing
    
    ' Fecha a conexão se não for pesquisa de chapa quem chamou esse metodo
    If fecharConexao = True Then
        ' Fechar conexão com banco
        Call fecharConexaoBanco
    End If
End Function

' Exclui objeto
Function excluir()

End Function

' Pesquisa objeto por id do tamanho
Function pesquisarPorIdTamanho(idTamanho As Variant) As objTamanho
    
    ' Metodos do metodo
    ' String para consultas
    Dim sqlSelectPesquisarPorId As String ' String para consultas
    Dim fkObject As String ' fk para consultas extras
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    
    'Abrindo conexão com banco
    Call conctarBanco
    ' String para consulta
    sqlSelectPesquisarPorId = "SELECT * FROM Tamanhos_Chapas " & "WHERE id_tamanho = " & idTamanho & ";"
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    ' Consulta banco
    rs.Open sqlSelectPesquisarPorId, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    ' Retorno da consulta
    While Not rs.EOF
        ' Criado os objetos
        Set tipoMaterial = ObjectFactory.factoryTipoMaterial(tipoMaterial)
        Set tamanho = ObjectFactory.factoryTamanho(tamanho)
        Set polideira = ObjectFactory.factoryPolideira(polideira)
        Set estoque = ObjectFactory.factoryEstoqueChapas(estoque)
        
        ' Atribuição dos atributos
        tamanho.id = rs.Fields("id_tamanho").Value
        tamanho.compremento = rs.Fields("comp_chapa").Value
        tamanho.altura = rs.Fields("alt_chapa").Value
        tamanho.qtdEstoque = rs.Fields("qtd_estoque").Value
        tamanho.qtdM2 = rs.Fields("qtd_m2").Value
        tamanho.valorPolimento = rs.Fields("valor_polimento")
        tamanho.espessura = rs.Fields("esp_chapa").Value
        
        ' Atribuições do objetos em tamanho
        ' fk para consulta
        fkObject = rs.Fields("fk_tipo_material").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Tipo_Material WHERE Id_Tipo_Material = " & fkObject & ";"
        ' Setando Objeto
        tamanho.setTipoMaterial retornarObjeto(tipoMaterial, sqlSelectPesquisarPorId, "Id_Tipo_Material", "Nome_Tipo_Material")
        
        ' fk para consulta
        fkObject = rs.Fields("fk_estoque").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Estoque_chapas WHERE Id_Estoque = " & fkObject & ";"
        ' Setando Objeto
        tamanho.setEstoque retornarObjeto(estoque, sqlSelectPesquisarPorId, "Id_Estoque", "Nome_Empresa")
        
        ' Só pesquisa se existir objeto
        If IsNull(fkObject = rs.Fields("fk_polidoria").Value) Then
            ' Seta objeto sem dados
            tamanho.setPolideira ObjectFactory.factoryPolideira(polideira)
        Else
            fkObject = rs.Fields("fk_polidoria").Value
            ' String para consulta
            sqlSelectPesquisarPorId = "SELECT * FROM Polideiras WHERE Id_Polidoria = " & fkObject & ";"
            ' Setando Objeto
            tamanho.setPolideira retornarObjeto(polideira, sqlSelectPesquisarPorId, "Id_Polidoria", "Nome_Polidoria")
        End If
        
'        ' fk para consulta
'        fkObject = rs.Fields("fK_chapa").Value
'        Set chapa = daoChapa.pesquisarPorId(fkObject)
'        tamanho.setChapa chapa
        
        rs.MoveNext
    Wend
    
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    
    ' Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorIdTamanho = tamanho
    
    ' Liberando espaço na memoria
    Set tipoMaterial = Nothing
    Set tamanho = Nothing
    Set polideira = Nothing
    Set estoque = Nothing
    Set chapa = Nothing
End Function

' Pesquisa objeto por id da chapa
Function pesquisarPorIdChapa(idchapa As Variant, conexaoFechar As Boolean) As Collection
    
    ' Metodos do metodo
    ' String para consultas
    Dim sqlSelectPesquisarPorId As String ' String para consultas
    Dim fkObject As String ' fk para consultas extras
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    
    'Abrindo conexão com banco
    Call conctarBanco
    ' String para consulta
    sqlSelectPesquisarPorId = "SELECT * FROM Tamanhos_Chapas " & "WHERE fK_chapa = '" & idchapa & "';"
    ' Seta a lista
    Set listaTamanhos = ObjectFactory.factoryLista(listaTamanhos)
    ' Criando e abrindo Recordset para consulta
    Set rs = ObjectFactory.factoryRsAuxiliar(rs)
    ' Consulta banco
    rs.Open sqlSelectPesquisarPorId, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    ' Retorno da consulta
    While Not rs.EOF
        ' Criado os objetos
        Set tipoMaterial = ObjectFactory.factoryTipoMaterial(tipoMaterial)
        Set tamanho = ObjectFactory.factoryTamanho(tamanho)
        Set polideira = ObjectFactory.factoryPolideira(polideira)
        Set estoque = ObjectFactory.factoryEstoqueChapas(estoque)
        
        ' Atribuição dos atributos
        tamanho.id = rs.Fields("id_tamanho").Value
        tamanho.compremento = rs.Fields("comp_chapa").Value
        tamanho.altura = rs.Fields("alt_chapa").Value
        tamanho.qtdEstoque = rs.Fields("qtd_estoque").Value
        tamanho.qtdM2 = rs.Fields("qtd_m2").Value
        tamanho.valorPolimento = rs.Fields("valor_polimento")
        tamanho.espessura = rs.Fields("esp_chapa").Value
        
        ' Atribuições do objetos em tamanho
        ' fk para consulta
        fkObject = rs.Fields("fk_tipo_material").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Tipo_Material WHERE Id_Tipo_Material = " & fkObject & ";"
        ' Setando Objeto
        tamanho.setTipoMaterial retornarObjeto(tipoMaterial, sqlSelectPesquisarPorId, "Id_Tipo_Material", "Nome_Tipo_Material")
        
        ' fk para consulta
        fkObject = rs.Fields("fk_estoque").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Estoque_chapas WHERE Id_Estoque = " & fkObject & ";"
        ' Setando Objeto
        tamanho.setEstoque retornarObjeto(estoque, sqlSelectPesquisarPorId, "Id_Estoque", "Nome_Empresa")
        
        ' Só pesquisa se existir objeto
        If IsNull(fkObject = rs.Fields("fk_polidoria").Value) Then
            ' Seta objeto sem dados
            tamanho.setPolideira ObjectFactory.factoryPolideira(polideira)
        Else
            fkObject = rs.Fields("fk_polidoria").Value
            ' String para consulta
            sqlSelectPesquisarPorId = "SELECT * FROM Polideiras WHERE Id_Polidoria = " & fkObject & ";"
            ' Setando Objeto
            tamanho.setPolideira retornarObjeto(polideira, sqlSelectPesquisarPorId, "Id_Polidoria", "Nome_Polidoria")
        End If
        
'        ' fk para consulta
'        fkObject = rs.Fields("fK_chapa").Value
'        Set chapa = daoChapa.pesquisarPorId(fkObject)
'        tamanho.setChapa chapa
        
        ' Setando tamanhos
        listaTamanhos.Add tamanho
        
        ' Liberando espaço na memoria
        Set tipoMaterial = Nothing
        Set tamanho = Nothing
        Set polideira = Nothing
        Set estoque = Nothing
        Set chapa = Nothing
        
        rs.MoveNext
    Wend
    
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    ' Fecha a conexão se não for pesquisa de chapa quem chamou esse metodo
    If conexaoFechar = True Then
        ' Fechar conexão com banco
        Call fecharConexaoBanco
    End If
    ' Retorno
    Set pesquisarPorIdChapa = listaTamanhos
    ' Liberando memoria
    Set listaTamanhos = Nothing
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
