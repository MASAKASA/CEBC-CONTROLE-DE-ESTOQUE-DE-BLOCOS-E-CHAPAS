VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents efeitoFrame As MSForms.Frame
Attribute efeitoFrame.VB_VarHelpID = -1

' Efeito fontes do menu
Public menuFontSizeSmall As Integer

' Efeito botoes texto
Public botoesTextoFontSizeSmall As Integer
'Efeito de passagem por cima do mouse
Private Sub efeitoFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    ' Atribuicoes
    menuFontSizeSmall = 20
    botoesTextoFontSizeSmall = 16
    
    'Efeitos paras os label botões com texto
    Select Case efeitoFrame.name
        ' Efeitos para os botões do menu
        Case "fTiraEfeitoBotoesMenu"
            formControle.Controls("btnLMenuHome").Font.Size = menuFontSizeSmall
            formControle.Controls("btnLMenuBloco").Font.Size = menuFontSizeSmall
            formControle.Controls("btnLMenuChapa").Font.Size = menuFontSizeSmall
            formControle.Controls("btnLMenuDespachar").Font.Size = menuFontSizeSmall
            formControle.Controls("btnLMenuCarrago").Font.Size = menuFontSizeSmall
            formControle.Controls("btnLMenuUsuarios").Font.Size = menuFontSizeSmall
            formControle.Controls("btnLMenuCadastros").Font.Size = menuFontSizeSmall
            
        ' Tela cadastrar bloco
        Case "fTiraEfeitoBotoesPedreiraBloco"
            formControle.Controls("btnLImgCadastrarPedreira").top = 24
            formControle.Controls("btnLImgCadastrarSerrariaCB").top = 78
            formControle.Controls("btnLImgCadastroTipoMaterial").top = 24
        Case "fTiraEfeitoBotoesAcoesBloco"
            formControle.Controls("btnLTxtCadastrarBloco").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtLimparCadastroBloco").Font.Size = botoesTextoFontSizeSmall
        
        ' Tela despache
        Case "fTiraEfeitoBotoesCabecalho"
            formControle.Controls("btnLImgCadastrarMotoristaDespache").top = 30
            formControle.Controls("btnLImgCadastrarDestinoDespache").top = 78
        Case "fTiraEfeitoBotoesMaterial"
            formControle.Controls("btnLTxtAdicionar").Font.Size = botoesTextoFontSizeSmall
        Case "fTiraEfeitoBotoesAcoesDespachar"
            formControle.Controls("btnLTxtDespachar").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtLimparDespache").Font.Size = botoesTextoFontSizeSmall
            
        ' Tela estoque m³
        Case "fTiraEfeitoBotoesFiltrosEstoqueM3"
            formControle.Controls("btnLTxtPesquisarBlocos").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtLimparFiltrosBlocos").Font.Size = botoesTextoFontSizeSmall
        Case "fTiraEfeitoBotoesExportarBlocosM3"
            formControle.Controls("btnLImgExportarEstoqueM3").top = 18
        Case "fTiraEfeitoBotoesAcoesListaBlocos"
            formControle.Controls("btnLTxtNovoBloco").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtEditarBloco").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtADDEstoque").Font.Size = botoesTextoFontSizeSmall
            
        ' Tela estoque m²
        Case "fTiraEfeitoBotoesFiltrosEstoqueM2"
            formControle.Controls("btnLTxtPesquisarChapas").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtLimparFiltrosChapas").Font.Size = botoesTextoFontSizeSmall
        Case "fTiraEfeitoBotoesExportarChapasM2"
            formControle.Controls("btnLImgExportarEstoqueM2").top = 6
        Case "fTiraEfeitoBotoesAcoesListChapas"
            formControle.Controls("btnLTxtNovoAvulso").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtEditarChapa").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtTrocaEstoque").Font.Size = botoesTextoFontSizeSmall
            
        ' Tela cadastro chapas avulso
        Case "fTiraEfeitoBotoesMaterialAvulso"
            formControle.Controls("btnLImgCadastrarMaterialAvulso").top = 24
            formControle.Controls("btnLImgCadastrarPolimentoAvulso").top = 78
        Case "fTiraEfeitoBotoesAcoesAvulso"
            formControle.Controls("btnLTxtCadastrarChapaAvulso").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtLimparCadastroChapaAvulso").Font.Size = botoesTextoFontSizeSmall
            
        ' Tela carregos
        Case "fTiraEfeitoBotoesFiltrosCarregos"
            formControle.Controls("btnLTxtPesquisarCarregos").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtLimparListas").Font.Size = botoesTextoFontSizeSmall
        Case "fTiraEfeitoBotoesExportarCarrego"
            formControle.Controls("btnLImgExportarCarregoPDF").top = 24
        Case "fTiraEfeitoBotoesAcoesListCarrego"
            formControle.Controls("btnLTxtEditarCarrego").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtVoltarCArrego").Font.Size = botoesTextoFontSizeSmall
            
        ' Tela editar bloco
        Case "fTiraEfeitoBotoesAcoesEdicaoBloco"
            formControle.Controls("btnLTxtSalvarEdicaoBloco").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtVoltarEdicaoBloco").Font.Size = botoesTextoFontSizeSmall
            
        ' Tela lançamento e edição chapa
        Case "fTiraEfeitoBotoesMaterialChapa"
            formControle.Controls("btnLImgCadastrarPolideiraChapa").top = 24
            formControle.Controls("btnLImgCadastrarTipoPolideiraChapa").top = 78
            formControle.Controls("btnLImgCadastrarTipoMaterialChapa").top = 132
        Case "fTiraEfeitoBotoesTamanhoChapa"
            formControle.Controls("btnLImgCadastrarTipoMaterialChapaTamanhos").top = 24
            formControle.Controls("btnLTxtAdicionarTamanhoChapa").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtEditarTamanhoChapa").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtTirarDaLista").Font.Size = botoesTextoFontSizeSmall
        Case "fTiraEfeitoBotoesAcoesEdicaoChapa"
            formControle.Controls("btnLTxtSalvarChapa").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtVoltarChapa").Font.Size = botoesTextoFontSizeSmall
            
        ' Tela Troca estoque
        Case "fTiraEfeitoBotoesTrocaPor"
            formControle.Controls("btnLTxtAdicionarTrocaEstoque").Font.Size = botoesTextoFontSizeSmall
        Case "fTiraEfeitoBotoesAcoesTrocaEstoque"
            formControle.Controls("btnLTxtTrocarEstoque").Font.Size = botoesTextoFontSizeSmall
            formControle.Controls("btnLTxtVoltarTrocaEstoque").Font.Size = botoesTextoFontSizeSmall
    End Select
End Sub
