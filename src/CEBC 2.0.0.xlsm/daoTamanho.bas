Attribute VB_Name = "daoTamanho"
Option Explicit

Private listaTamanhos As Collection
Private tipoMaterial As objTipoMaterial
Private tamanho As objTamanho
' Cadastra e edita objeto
Function cadastrarEEditar(lista As Collection)
    ' String para consultas
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    Dim rsAuxiliar As ADODB.Recordset ' Recordset para consulta
    Dim strSql As String ' String para consultas
    Dim cadastro As Boolean
    Dim i As Long

    ' Seta true em cadastro
    cadastro = True
    
    For Each idLista In idsParaPesquisa
        ' Seta id para pesquisa
        id = idLista
        
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Blocos " & "WHERE Id_Bloco = '" & id & "';"
        ' Criando e abrindo Recordset para consulta
        Set rsBloco = ObjectFactory.factoryRsAuxiliar(rsBloco)
        ' Consulta banco
        rsBloco.Open sqlSelectPesquisarPorId, CONEXAO_BD, adOpenKeyset, adLockReadOnly
        ' Retorno da consulta
        While Not rsBloco.EOF
            ' Criação e atribuição dos objetos
            Set bloco = ObjectFactory.factoryBloco(bloco)
            Set pedreira = ObjectFactory.factoryPedreira(pedreira)
            Set serraria = ObjectFactory.factorySerraria(serraria)
            Set polideira = ObjectFactory.factoryPolideira(polideira)
            Set status = ObjectFactory.factoryStatus(status)
            Set tipoMaterial = ObjectFactory.factoryTipoMaterial(tipoMaterial)
            Set estoque = ObjectFactory.factoryEstoque(estoque)
            
            ' Atribuição dos atributos
            bloco.idSistema = rsBloco.Fields("Id_Bloco").Value
            bloco.nomeMaterial = rsBloco.Fields("Descricao").Value
            bloco.observacao = rsBloco.Fields("Observacao").Value
            bloco.numeroBlocoPedreira = rsBloco.Fields("Id_bloco_Pedreira").Value
            bloco.dataCadastro = rsBloco.Fields("Data_cadastro").Value
            bloco.qtdM3 = rsBloco.Fields("Quantidade_M3").Value
            bloco.qtdM2Serrada = rsBloco.Fields("Quantidade_Serrada_M2").Value
            
            
            'Atribuições dos objetos em bloco
            ' fk para consulta
            fkObject = rsBloco.Fields("Fk_Pedreira").Value
            ' String para consulta
            sqlSelectPesquisarPorId = "SELECT * FROM Pedreiras WHERE Id_Pedreira = " & fkObject & ";"
            ' Setando Objeto
            bloco.setPedreira retornarObjeto(pedreira, sqlSelectPesquisarPorId, "Id_Pedreira", "Nome_Pedreira")
            
            
            ' Libera espaço para nova pesquisa se ouver
            Set bloco = Nothing
            Set pedreira = Nothing
            Set serraria = Nothing
            Set polideira = Nothing
            Set status = Nothing
            Set tipoMaterial = Nothing
            Set estoque = Nothing
            
            rsBloco.MoveNext
        Wend
        
        ' Libera recurso Recordset
        rsBloco.Close
        Set rsBloco = Nothing
    Next idLista
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ' Faz a consulta para saber se o código do bloco já exite
    strSql = "SELECT * FROM Polideiras" _
        & " WHERE Id_Polidoria = " & polideira.id & ";"
    
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
        ' Concatenando comando SQL e cadastrando bloco no banco de dados
        strSql = "INSERT INTO Polideiras ( Nome_Polidoria )VALUES ('" & polideira.nome & "');"
        
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    Else
        ' Se edição
        strSql = "UPDATE Polideiras SET Nome_Pedreira = '" & polideira.nome & "', " _
            & "' WHERE Id_Polidoria = '" & polideira.id & "';"
            
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
    End If
    
    Set rs = Nothing
    ' Fechando conexão com banco
    Call fecharConexaoBanco
End Function

' Exclui objeto
Function excluir()

End Function

' Pesquisa objeto por id
Function pesquisarPorIdChapa(idChapa As String) As Collection
    ' Metodos do metodo
    ' String para consultas
    Dim sqlSelectPesquisarPorId As String ' String para consultas
    Dim fkObject As String ' fk para consultas extras
    Dim rs As ADODB.Recordset ' Recordset para consulta principal
    
    'Abrindo conexão com banco
    Call conctarBanco
    ' String para consulta
    sqlSelectPesquisarPorId = "SELECT * FROM Tamanhos_Chapas " & "WHERE id_Chapa = '" & idChapa & "';"
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
        
        ' Atribuição dos atributos
        tamanho.id = rs.Fields("id_tamanho").Value
        tamanho.idChapa = rs.Fields("id_Chapa").Value
        tamanho.compremento = rs.Fields("comp_chapa").Value
        tamanho.altura = rs.Fields("alt_chapa").Value
        tamanho.qtdEstoque = rs.Fields("qtd_estoque").Value
        tamanho.qtdM2 = rs.Fields("qtd_m2").Value
        tamanho.espessura = rs.Fields("esp_chapa").Value
        
        ' Atribuições do objetos em tamanho
        ' fk para consulta
        fkObject = rsBloco.Fields("Fk_Tipo_Material").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Tipo_Material WHERE Id_Tipo_Material = " & fkObject & ";"
        ' Setando Objeto
        tamanho.setTipoMaterial retornarObjeto(tipoMaterial, sqlSelectPesquisarPorId, "Id_Tipo_Material", "Nome_Tipo_Material")
        
        ' Setando tamanhos
        listaTamanhos.Add tamanho
        
        ' Liberando espaço na memoria
        Set tipoMaterial = Nothing
        Set tamanho = Nothing
        
        rs.MoveNext
    Wend
    
    ' Libera recurso Recordset
    rs.Close
    Set rs = Nothing
    ' Fechar conexão com banco
    Call fecharConexaoBanco
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
