Attribute VB_Name = "daoBloco"
Option Explicit

Private bloco As objBloco
Dim pedreira As objPedreira
Dim serraria As objSerraria
Dim polideira As objPolideira
Dim tipoMaterial As objTipoMaterial
Dim status As objStatus

' Cadastra e edita objeto
Function cadastrarEEditar(bloco As objBloco)


End Function

' Exclui objeto
Function excluir()

End Function

' Pesquisa objeto por id
Function pesquisarPorId() As objBloco
    ' Criação e atribuição dos objetos
    Set bloco = New objBloco
    Set pedreira = New objPedreira
    Set serraria = New objSerraria
    Set polideira = New objPolideira
    Set status = New objStatus
    Set tipoMaterial = New objTipoMaterial
    
    ' Pesquisa e atribuições dos objetos
    pedreira.carregarPedreiraManipulacao "01", "MINERAÇÃO VISTA LINDA"
    serraria.carregarSerrariaManipulacao "01", "ELSON BABISQUE"
    polideira.carregarPolideiraManipulacao "01", "SÃO ROQUE"
    status.carregarStatusManipulacao "01", "EM PROCESSO"
    tipoMaterial.carregarTipoMaterialManipulacao "01", "COMERCIAL SATAND"
    
    ' Atribuição dos atributos
    bloco.idSistema = "37766-50793-MOON-LIGHT-BL"
    bloco.nomeMaterial = "BLOCO MARMORE BRANCO CLASSICO"
    bloco.observacao = "BLOCO COM CHAPAS QUEBRADAS SERÃO REPOSTAS POSTERIORMENTE"
    bloco.numeroBlocoPedreira = "37766-50793-MOON-LIGHT-BL"
    bloco.estoque = "CASA DO GRANITO"
    bloco.dataCadastro = "22/02/2024"
    bloco.qtdM3 = "12,255"
    bloco.qtdM2Serrada = "359,5448"
    bloco.qtdM2Polimento = "284,4578"
    bloco.qtdChapas = "71"
    bloco.nota = "SIM"
    bloco.consultarCustoMedio = "SIM"
    bloco.compBrutoBloco = "3,8000"
    bloco.altBrutoBloco = "2,8000"
    bloco.largBrutoBloco = "2,8000"
    bloco.compLiquidoBloco = "3,5000"
    bloco.altLiquidoBloco = "2,5000"
    bloco.largLiquidoBloco = "2,5000"
    bloco.compBrutoChapaBruta = "3,300"
    bloco.altBrutoChapaBruta = "2,300"
    bloco.compLiquidoChapaBruta = "3,000"
    bloco.altBrutoChapaBruta = "2,000"
    bloco.compBrutoChapaPolida = "2,9000"
    bloco.altBrutoChapaPolida = "1,9000"
    bloco.compLiquidoChapaPolida = "2,5000"
    bloco.altBrutoChapaPolida = "1,5000"
    bloco.valorBloco = "6000,00"
    bloco.precoM3Bloco = "600,00"
    bloco.freteBloco = "1500,00"
    bloco.valorMetroSerrada = "11,00"
    bloco.valorMetroPolimento = "22,00"
    bloco.valoresAdicionais = "5000,00"
    bloco.valorTotalSerrada = "11000,00"
    bloco.valorTotalPolimento = "9000,00"
    bloco.custoMaterial = "90,00"
    bloco.qtdM2Polimento = "284,4578"
    bloco.valorTotalBloco = "12.500,00"
    
    bloco.setPedreira pedreira
    bloco.setSerraria serraria
    bloco.setPolideira polideira
    bloco.setStatus status
    bloco.setTipoMaterial tipoMaterial
    
    Set pesquisarPorId = bloco
End Function

' Pesquisa objeto por nome
Function pesquisarPorNome()

End Function

' Pesquisa objeto
Function listarBlocosFilter()
    ' Chama Serviço
    MsgBox "Retorna pesquiar"
End Function
