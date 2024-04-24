VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objBloco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public idSistema As String
Public nomeMaterial As String
Public observacao As String
Public numeroBlocoPedreira As String
Public dataCadastro As String
Public qtdM3 As String
Public qtdM2Serrada As String
Public qtdM2Polimento As String
Public qtdChapas As String
Public nota As String
Public consultarCustoMedio As String
Public compBrutoBloco As String
Public altBrutoBloco As String
Public largBrutoBloco As String
Public compLiquidoBloco As String
Public altLiquidoBloco As String
Public largLiquidoBloco As String
Public compBrutoChapaBruta As String
Public altBrutoChapaBruta As String
Public compLiquidoChapaBruta As String
Public altLiquidoChapaBruta As String
Public compBrutoChapaPolida As String
Public altBrutoChapaPolida As String
Public compLiquidoChapaPolida As String
Public altLiquidoChapaPolida As String
Public valorBloco As String
Public precoM3Bloco As String
Public freteBloco As String
Public valorMetroSerrada As String
Public valorMetroPolimento As String
Public valoresAdicionais As String
Public valorTotalSerrada As String
Public valorTotalPolimento As String
Public custoMaterial As String
Public valorTotalBloco As String
Public estoque As objEstoque
Public status As objStatus
Public tipoMaterial As objTipoMaterial
Public pedreira As objPedreira
Public serraria As objSerraria
Public polideira As objPolideira
' Construtor para cadastro
Public Sub carregarBlocoCadastro(dataCadastro As String, idSistema As String, pedreira As objPedreira, serraria As objSerraria, _
                        numeroBlocoPedreira As String, nomeMaterial As String, tipoMaterial As objTipoMaterial, nota As String, _
                        status As objStatus, observacao As String, compBrutoBloco As String, altBrutoBloco As String, largBrutoBloco As String, _
                        compLiquidoBloco As String, altLiquidoBloco As String, largLiquidoBloco As String, estoque As objEstoque, _
                        valoresAdicionais As String, freteBloco As String, precoM3Bloco As String, qtdM3 As String, valorBloco As String)
    ' Atribuições
    Me.dataCadastro = dataCadastro
    Me.idSistema = idSistema
    Set Me.pedreira = pedreira
    Set Me.serraria = serraria
    Set Me.estoque = estoque
    Me.numeroBlocoPedreira = numeroBlocoPedreira
    Me.nomeMaterial = nomeMaterial
    Set Me.tipoMaterial = tipoMaterial
    Me.nota = nota
    Set Me.status = status
    Me.observacao = observacao
    Me.compBrutoBloco = compBrutoBloco
    Me.altBrutoBloco = altBrutoBloco
    Me.largBrutoBloco = largBrutoBloco
    Me.compLiquidoBloco = compLiquidoBloco
    Me.altLiquidoBloco = altLiquidoBloco
    Me.largLiquidoBloco = largLiquidoBloco
    Me.valoresAdicionais = valoresAdicionais
    Me.freteBloco = freteBloco
    Me.precoM3Bloco = precoM3Bloco
    Me.qtdM3 = qtdM3
    Me.valorBloco = valorBloco
End Sub

' Construtor para manipulação
Public Sub carregarBlocoEdicao(idSistema As String, nomeMaterial As String, observacao As String, _
                        numeroBlocoPedreira As String, estoque As objEstoque, dataCadastro As String, _
                        qtdM3 As String, qtdM2Serrada As String, qtdM2Polimento As String, qtdChapas As String, _
                        nota As String, consultarCustoMedio As String, compBrutoBloco As String, _
                        altBrutoBloco As String, largBrutoBloco As String, compLiquidoBloco As String, _
                        altLiquidoBloco As String, largLiquidoBloco As String, compBrutoChapaBruta As String, _
                        altBrutoChapaBruta As String, compLiquidoChapaBruta As String, altLiquidoChapaBruta As String, _
                        compBrutoChapaPolida As String, altBrutoChapaPolida As String, compLiquidoChapaPolida As String, _
                        altLiquidoChapaPolida As String, valorBloco As String, precoM3Bloco As String, _
                        freteBloco As String, valorMetroSerrada As String, valorMetroPolimento As String, _
                        valoresAdicionais As String, valorTotalSerrada As String, valorTotalPolimento As String, _
                        custoMaterial As String, valorTotalBloco As String, status As objStatus, _
                        tipoMaterial As objTipoMaterial, pedreira As objPedreira, serraria As objSerraria, _
                        polideira As objPolideira)
    ' Atribuições
    Me.idSistema = idSistema
    Me.nomeMaterial = nomeMaterial
    Me.observacao = observacao
    Me.numeroBlocoPedreira = numeroBlocoPedreira
    Me.dataCadastro = dataCadastro
    Me.qtdM3 = qtdM3
    Me.qtdM2Serrada = qtdM2Serrada
    Me.qtdM2Polimento = qtdM2Polimento
    Me.qtdChapas = qtdChapas
    Me.nota = nota
    Me.consultarCustoMedio = consultarCustoMedio
    Me.compBrutoBloco = compBrutoBloco
    Me.altBrutoBloco = altBrutoBloco
    Me.largBrutoBloco = largBrutoBloco
    Me.compLiquidoBloco = compLiquidoBloco
    Me.altLiquidoBloco = altLiquidoBloco
    Me.largLiquidoBloco = largLiquidoBloco
    Me.compBrutoChapaBruta = compBrutoChapaBruta
    Me.altBrutoChapaBruta = altBrutoChapaBruta
    Me.compLiquidoChapaBruta = compLiquidoChapaBruta
    Me.altLiquidoChapaBruta = altLiquidoChapaBruta
    Me.compBrutoChapaPolida = compBrutoChapaPolida
    Me.altBrutoChapaPolida = altBrutoChapaPolida
    Me.compLiquidoChapaPolida = compLiquidoChapaPolida
    Me.altLiquidoChapaPolida = altLiquidoChapaPolida
    Me.valorBloco = valorBloco
    Me.precoM3Bloco = precoM3Bloco
    Me.freteBloco = freteBloco
    Me.valorMetroSerrada = valorMetroSerrada
    Me.valorMetroPolimento = valorMetroPolimento
    Me.valoresAdicionais = valoresAdicionais
    Me.valorTotalSerrada = valorTotalSerrada
    Me.valorTotalPolimento = valorTotalPolimento
    Me.custoMaterial = custoMaterial
    Me.valorTotalBloco = valorTotalBloco
    Set Me.estoque = estoque
    Set Me.status = status
    Set Me.tipoMaterial = tipoMaterial
    Set Me.pedreira = pedreira
    Set Me.serraria = serraria
    Set Me.polideira = polideira
End Sub
' Gets e sets
Public Function getStatus() As objStatus
    getStatus = status
End Function

Public Function setStatus(status As objStatus)
    Set Me.status = status
End Function
Public Function getPedreira() As objPedreira
    getPedreira = pedreira
End Function

Public Function setPedreira(pedreira As objPedreira)
    Set Me.pedreira = pedreira
End Function

Public Function getSerraria() As objSerraria
    getSerraria = serraria
End Function

Public Function setSerraria(serraria As objSerraria)
    Set Me.serraria = serraria
End Function

Public Function getPolideira() As objPolideira
    getPolideira = polideira
End Function

Public Function setPolideira(polideira As objPolideira)
    Set Me.polideira = polideira
End Function

Public Function getTipoMaterial() As objTipoMaterial
    getTipoMaterial = tipoMaterial
End Function

Public Function setTipoMaterial(tipoMaterial As objTipoMaterial)
    Set Me.tipoMaterial = tipoMaterial
End Function

Public Function getEstoque() As objEstoque
    getEstoque = estoque
End Function

Public Function setEstoque(estoque As objEstoque)
    Set Me.estoque = estoque
End Function








