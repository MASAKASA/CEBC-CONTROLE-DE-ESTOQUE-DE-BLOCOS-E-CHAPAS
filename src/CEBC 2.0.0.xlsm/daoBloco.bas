Attribute VB_Name = "daoBloco"
Option Explicit

Private listaBlocos As Collection
Private bloco As objBloco
Private pedreira As objPedreira
Private serraria As objSerraria
Private polideira As objPolideira
Private tipoMaterial As objTipoMaterial
Private status As objStatus
Private estoque As objEstoque

' Cadastra e edita objeto
Function cadastrarEEditar(bloco As objBloco)


End Function

' Exclui objeto
Function excluir(id As String)

End Function

' Pesquisa objeto por id
Function pesquisarPorId(id As String) As objBloco
    'Metodos do metodo
    ' String para consultas
    Dim sqlSelectPesquisarPorId As String ' String para consultas
    Dim fkObject As String ' fk para consultas extras
    Dim rsBloco As ADODB.Recordset ' Recordset para consulta principal
    
    ' Criação e atribuição dos objetos
    Set bloco = ObjectFactory.factoryBloco(bloco)
    Set pedreira = ObjectFactory.factoryPedreira(pedreira)
    Set serraria = ObjectFactory.factorySerraria(serraria)
    Set polideira = ObjectFactory.factoryPolideira(polideira)
    Set status = ObjectFactory.factoryStatus(status)
    Set tipoMaterial = ObjectFactory.factoryTipoMaterial(tipoMaterial)
    Set estoque = ObjectFactory.factoryEstoque(estoque)
    
   'Abrindo conexão com banco
    Call conctarBanco
    
    ' String para consulta
    sqlSelectPesquisarPorId = "SELECT * FROM Blocos " & "WHERE Id_Bloco = '" & id & "' ORDER BY Descricao;"
    ' Criando e abrindo Recordset para consulta
    Set rsBloco = ObjectFactory.factoryRsBloco(rsBloco)
    ' Consulta banco
    rsBloco.Open sqlSelectPesquisarPorId, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    ' Retorno da consulta
    While Not rsBloco.EOF
        ' Atribuição dos atributos
        bloco.idSistema = rsBloco.Fields("Id_Bloco").Value
        bloco.nomeMaterial = rsBloco.Fields("Descricao").Value
        bloco.observacao = rsBloco.Fields("Observacao").Value
        bloco.numeroBlocoPedreira = rsBloco.Fields("Id_bloco_Pedreira").Value
        bloco.dataCadastro = rsBloco.Fields("Data_cadastro").Value
        bloco.qtdM3 = rsBloco.Fields("Quantidade_M3").Value
        bloco.qtdM2Serrada = rsBloco.Fields("Quantidade_Serrada_M2").Value
        bloco.qtdM2Polimento = rsBloco.Fields("Quantidade_Polimento_M2").Value
        bloco.qtdChapas = rsBloco.Fields("Total_chapas").Value
        bloco.nota = rsBloco.Fields("Tem_Nota").Value
        bloco.consultarCustoMedio = rsBloco.Fields("Custo_Medio").Value
        bloco.compBrutoBloco = rsBloco.Fields("Comp_Bruto_Bloco").Value
        bloco.altBrutoBloco = rsBloco.Fields("Alt_Bruto_Bloco").Value
        bloco.largBrutoBloco = rsBloco.Fields("Larg_Bruto_Bloco").Value
        bloco.compLiquidoBloco = rsBloco.Fields("Comp_Liquida_Bloco").Value
        bloco.altLiquidoBloco = rsBloco.Fields("Alt_Liquida_Bloco").Value
        bloco.largLiquidoBloco = rsBloco.Fields("Larg_Liquida_Bloco").Value
        bloco.compBrutoChapaBruta = rsBloco.Fields("Comp_Bruto_Chapa_Bruta").Value
        bloco.altBrutoChapaBruta = rsBloco.Fields("Alt_Bruto_Chapa_Bruta").Value
        bloco.compLiquidoChapaBruta = rsBloco.Fields("Comp_Liquido_Chapa_Bruta").Value
        bloco.altLiquidoChapaBruta = rsBloco.Fields("Alt_Liquido_Chapa_Bruta").Value
        bloco.compBrutoChapaPolida = rsBloco.Fields("Comp_Bruto_Chapa_Polida").Value
        bloco.altBrutoChapaPolida = rsBloco.Fields("Comp_Bruto_Chapa_Polida").Value
        bloco.compLiquidoChapaPolida = rsBloco.Fields("Comp_Liquido_Chapa_Polida").Value
        bloco.altLiquidoChapaPolida = rsBloco.Fields("Alt_Liquido_Chapa_Polida").Value
        bloco.valorBloco = rsBloco.Fields("Preço_Bloco").Value
        bloco.precoM3Bloco = rsBloco.Fields("Valor_M3").Value
        bloco.freteBloco = rsBloco.Fields("Valor_Frete").Value
        bloco.valorMetroSerrada = rsBloco.Fields("Valor_Serrada").Value
        bloco.valorMetroPolimento = rsBloco.Fields("Valor_Polimento").Value
        bloco.valoresAdicionais = rsBloco.Fields("Valores_Adicionais").Value
        bloco.valorTotalSerrada = rsBloco.Fields("Valor_Serrada").Value
        bloco.valorTotalPolimento = rsBloco.Fields("Valor_Polimento").Value
        bloco.custoMaterial = rsBloco.Fields("Custo_Material").Value
        bloco.valorTotalBloco = rsBloco.Fields("Custo_Total").Value
        
        'Atribuições dos objetos em bloco
        ' fk para consulta
        fkObject = rsBloco.Fields("Fk_Pedreira").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Pedreiras WHERE Id_Pedreira = " & fkObject & ";"
        ' Setando Objeto
        bloco.setPedreira retornarObjeto(pedreira, sqlSelectPesquisarPorId, "Id_Pedreira", "Nome_Pedreira")
        
        ' fk para consulta
        fkObject = rsBloco.Fields("Fk_Serraria").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Serrarias WHERE Id_Serraria = " & fkObject & ";"
        ' Setando Objeto
        bloco.setSerraria retornarObjeto(serraria, sqlSelectPesquisarPorId, "Id_Serraria", "Nome_Serraria")
        
        ' fk para consulta
        fkObject = rsBloco.Fields("Fk_Polideira").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Polideiras WHERE Id_Polidoria = " & fkObject & ";"
        ' Setando Objeto
        bloco.setPolideira retornarObjeto(polideira, sqlSelectPesquisarPorId, "Id_Polidoria", "Nome_Polidoria")
        
        ' fk para consulta
        fkObject = rsBloco.Fields("Fk_Status").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Status WHERE Id_Status = " & fkObject & ";"
        ' Setando Objeto
        bloco.setStatus retornarObjeto(status, sqlSelectPesquisarPorId, "Id_Status", "Nome_Status")
    
        ' fk para consulta
        fkObject = rsBloco.Fields("Fk_Tipo_Material").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Tipo_Material WHERE Id_Tipo_Material = " & fkObject & ";"
        ' Setando Objeto
        bloco.setTipoMaterial retornarObjeto(tipoMaterial, sqlSelectPesquisarPorId, "Id_Tipo_Material", "Nome_Tipo_Material")
        
        ' fk para consulta
        fkObject = rsBloco.Fields("Fk_Estoque").Value
        ' String para consulta
        sqlSelectPesquisarPorId = "SELECT * FROM Estoque_blocos WHERE Id_Estoque = " & fkObject & ";"
        ' Setando Objeto
        bloco.setEstoque retornarObjeto(estoque, sqlSelectPesquisarPorId, "Id_Estoque", "Empresa")
        
        rsBloco.MoveNext
    Wend
    
    ' Libera recurso Recordset
    rsBloco.Close
    Set rsBloco = Nothing
    ' Fechar conexão com banco
    Call fecharConexaoBanco
    
    ' Retorno
    Set pesquisarPorId = bloco
End Function

' Pesquisa objeto por nome
Function pesquisarPorNome(nome As String) As Collection
    Set listaBlocos = New Collection
    
    Set pesquisarPorNome = listaBlocos
End Function

' Pesquisa objeto
Function listarBlocosFilter() As Collection
    ' Chama Serviço
    MsgBox "Retorna pesquiar"
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
