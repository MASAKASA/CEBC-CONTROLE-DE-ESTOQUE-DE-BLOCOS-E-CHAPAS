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
'Efeito de passagem por cima do mouse
Private Sub efeitoFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    'Efeitos paras os label botões com texto
    Select Case efeitoFrame.name
        ' Efeitos para os botões do menu
        Case "fTiraEfeitoBotoesMenu"
            formControle.Controls("btnLMenuHome").Font.Size = 20
            formControle.Controls("btnLMenuBloco").Font.Size = 20
            formControle.Controls("btnLMenuChapa").Font.Size = 20
            formControle.Controls("btnLMenuDespachar").Font.Size = 20
            formControle.Controls("btnLMenuCarrago").Font.Size = 20
            formControle.Controls("btnLMenuUsuarios").Font.Size = 20
            formControle.Controls("btnLMenuCadastros").Font.Size = 20
            
        ' Tela cadastrar bloco
        Case "fTiraEfeitoBotoesPedreiraBloco"
            formControle.Controls("btnLImgCadastrarPedreira").top = 24
            formControle.Controls("btnLImgCadastrarSerrariaCB").top = 78
            formControle.Controls("btnLImgCadastroTipoMaterial").top = 24
        Case "fTiraEfeitoBotoesAcoesBloco"
            formControle.Controls("btnLTxtCadastrarBloco").Font.Size = 16
            formControle.Controls("btnLTxtLimparCadastroBloco").Font.Size = 16
        
        ' Tela despache
        Case "fTiraEfeitoBotoesPesquisaEAdicicao"
            formControle.Controls("btnLImgCadastrarMotoristaDespache").top = 72
            formControle.Controls("btnLImgCadastrarDestinoDespache").top = 72
            formControle.Controls("btnLTxtAdicionar").Font.Size = 16
        Case "fTiraEfeitoBotoesExportarDespachePDF"
            formControle.Controls("btnLImgExportarListaDespache").top = 6
        Case "fTiraEfeitoBotoesAcoesDespachar"
            formControle.Controls("btnLTxtDespachar").Font.Size = 16
            formControle.Controls("btnLTxtLimparDespache").Font.Size = 16
            
        ' Tela estoque m³
        Case "fTiraEfeitoBotoesFiltrosEstoqueM3"
            formControle.Controls("btnLTxtPesquisarBlocos").Font.Size = 16
            formControle.Controls("btnLTxtLimparFiltrosBlocos").Font.Size = 16
        Case "fTiraEfeitoBotoesExportarBlocosM3"
            formControle.Controls("btnLImgExportarEstoqueM3").top = 18
        Case "fTiraEfeitoBotoesAcoesListaBlocos"
            formControle.Controls("btnLTxtNovoBloco").Font.Size = 16
            formControle.Controls("btnLTxtEditarBloco").Font.Size = 16
            formControle.Controls("btnLTxtADDEstoque").Font.Size = 16
            
        ' Tela estoque m²
        Case "fTiraEfeitoBotoesFiltrosEstoqueM2"
            formControle.Controls("btnLTxtPesquisarChapas").Font.Size = 16
            formControle.Controls("btnLTxtLimparFiltrosChapas").Font.Size = 16
        Case "fTiraEfeitoBotoesExportarChapasM2"
            formControle.Controls("btnLImgExportarEstoqueM2").top = 6
        Case "fTiraEfeitoBotoesAcoesListChapas"
            formControle.Controls("btnLTxtNovoAvulso").Font.Size = 16
            formControle.Controls("btnLTxtEditarChapa").Font.Size = 16
            formControle.Controls("btnLTxtTrocaEstoque").Font.Size = 16
            
        ' Tela cadastro chapas avulso
        Case "fTiraEfeitoBotoesMaterialAvulso"
            formControle.Controls("btnLImgCadastrarMaterialAvulso").top = 24
            formControle.Controls("btnLImgCadastrarPolimentoAvulso").top = 78
        Case "fTiraEfeitoBotoesAcoesAvulso"
            formControle.Controls("btnLTxtCadastrarChapaAvulso").Font.Size = 16
            formControle.Controls("btnLTxtLimparCadastroChapaAvulso").Font.Size = 16
            
            
            
     
        
    End Select
End Sub
