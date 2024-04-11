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
Public tipoMaterial As String
Public observacao As String
Public numeroBlocoPedreira As String
Public estoque As String
Public dataCadastro As String
Public qtdM3 As String
Public qtdM2Serrada As String
Public qtdM2Polimento As String
Public qtdChapas As String
Public status As String
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
Public pedreira As objPedreira
Public serraria As objSerraria
Public polideira As objPolideira
' Construtor do bloco para cadastro
Public Sub carregarBlocoCadastro(dataCadastro As String, idSistema As String, pedreira As objPedreira, serraria As objSerraria, _
                        numeroBlocoPedreira As String, nomeMaterial As String, tipoMaterial As String, nota As String, status As String, _
                        observacao As String, compBrutoBloco As String, altBrutoBloco As String, largBrutoBloco As String, _
                        compLiquidoBloco As String, altLiquidoBloco As String, largLiquidoBloco As String, valoresAdicionais As String, _
                        freteBloco As String, precoM3Bloco As String, qtdM3 As String, valorBloco As String)
    ' Atribuições
    Me.dataCadastro = dataCadastro
    Me.idSistema = idSistema
    Set Me.pedreira = pedreira
    Set Me.serraria = serraria
    Me.numeroBlocoPedreira = numeroBlocoPedreira
    Me.nomeMaterial = nomeMaterial
    Me.tipoMaterial = tipoMaterial
    Me.nota = nota
    Me.status = status
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

' Construtor do bloco para edição
Public Sub carregarBlocoEdicao(idSistema As String, nomeMaterial As String, tipoMaterial As String, _
                        observacao As String, numeroBlocoPedreira As String, estoque As String, _
                        dataCadastro As String, qtdM3 As String, qtdM2Serrada As String, qtdM2Polimento As String, _
                        qtdChapas As String, status As String, nota As String, consultarCustoMedio As String, _
                        compBrutoBloco As String, altBrutoBloco As String, largBrutoBloco As String, _
                        compLiquidoBloco As String, altLiquidoBloco As String, largLiquidoBloco As String, _
                        compBrutoChapaBruta As String, altBrutoChapaBruta As String, compLiquidoChapaBruta As String, _
                        altLiquidoChapaBruta As String, compBrutoChapaPolida As String, altBrutoChapaPolida As String, _
                        compLiquidoChapaPolida As String, altLiquidoChapaPolida As String, valorBloco As String, _
                        precoM3Bloco As String, freteBloco As String, valorMetroSerrada As String, valorMetroPolimento As String, _
                        valoresAdicionais As String, valorTotalSerrada As String, valorTotalPolimento As String, _
                        custoMaterial As String, valorTotalBloco As String, pedreira As objPedreira, _
                        serraria As objSerraria, polideira As objPolideira)
    ' Atribuições
    Me.idSistema = idSistema
    Me.nomeMaterial = nomeMaterial
    Me.tipoMaterial = tipoMaterial
    Me.observacao = observacao
    Me.numeroBlocoPedreira = numeroBlocoPedreira
    Me.estoque = estoque
    Me.dataCadastro = dataCadastro
    Me.qtdM3 = qtdM3
    Me.qtdM2Serrada = qtdM2Serrada
    Me.qtdM2Polimento = qtdM2Polimento
    Me.qtdChapas = qtdChapas
    Me.status = status
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
    Set Me.pedreira = pedreira
    Set Me.serraria = serraria
    Set Me.polideira = polideira
End Sub













