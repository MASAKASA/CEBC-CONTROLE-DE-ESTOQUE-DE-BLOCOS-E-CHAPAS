VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formControle 
   Caption         =   "CONTROLE DE BLOCOS E CHAPAS 2.0.0"
   ClientHeight    =   13410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   24645
   OleObjectBlob   =   "formControle.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "formControle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim botoesMenu() As clsLabel
Dim botoesImg() As clsLabel
Dim botoesText() As clsLabel
Dim frameEfeito() As clsFrame
'Inicializa��o do formControle
Private Sub UserForm_Initialize()

    Dim obj As Object
    Dim i As Long
    Dim j As Long
    Dim l As Long
    Dim m As Long
    Dim nameObj As String
    Dim nameObjInicio As String
    
    ReDim botoesMenu(1 To Me.Controls.Count)
    ReDim botoesImg(1 To Me.Controls.Count)
    ReDim botoesText(1 To Me.Controls.Count)
    ReDim frameEfeito(1 To Me.Controls.Count)
    
    For Each obj In Me.Controls
        
        nameObj = obj.name
        nameObjInicio = Mid(nameObj, 1, 7)
        
        If nameObjInicio = "btnLImg" Then
            i = i + 1
            Set botoesImg(i) = New clsLabel
            Set botoesImg(i).efeitoBotoesImagem = obj
        End If
        
        If nameObjInicio = "btnLTxt" Then
            j = j + 1
            Set botoesText(j) = New clsLabel
            Set botoesText(j).efeitoBotoesTexto = obj
        End If
        
        If nameObjInicio = "fTiraEf" Then
            l = l + 1
            Set frameEfeito(l) = New clsFrame
            Set frameEfeito(l).efeitoFrame = obj
        End If
        
        If nameObjInicio = "btnLMen" Then
            m = m + 1
            Set botoesMenu(m) = New clsLabel
            Set botoesMenu(m).efeitoBotoesMenu = obj
        End If
        
    Next obj
    
    Set obj = Nothing
    
    ReDim Preserve botoesImg(1 To i)
    ReDim Preserve botoesText(1 To j)
    ReDim Preserve frameEfeito(1 To l)
    ReDim Preserve botoesMenu(1 To m)
    
    
End Sub

'-----------------------------------------------------------------MENU DO SISTEMA-----------------------------------
'                                                                 ---------------
'Efeito para clique nas label btnLMenuHome do menu
Private Sub btnLMenuHome_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 10
End Sub
'Efeito para clique nas label btnLMenuBloco do menu
Private Sub btnLMenuBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 3
End Sub
'Efeito para clique nas label btnLMenuChapa do menu
Private Sub btnLMenuChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 4
End Sub
'Efeito para clique nas label btnLMenuDespachar do menu
Private Sub btnLMenuDespachar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 2
End Sub
'Efeito para clique nas label btnLMenuCarrago do menu
Private Sub btnLMenuCarrago_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 6
End Sub
'Efeito para clique nas label btnLMenuUsuarios do menu
Private Sub btnLMenuUsuarios_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 7
End Sub
'Efeito para clique nas label btnLMenuCadastros do menu
Private Sub btnLMenuCadastros_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 7
End Sub

'-----------------------------------------------------------------TELA CADASTRO DE BLOCOS-----------------------------------
'                                                                 -----------------------
'Bot�o btnLImgCadastrarPedreira tela cadastrar bloco
Private Sub btnLImgCadastrarPedreira_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar pedreira"
End Sub
'Bot�o btnLImgCadastrarSerrariaCB tela cadastrar bloco
Private Sub btnLImgCadastrarSerrariaCB_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar serraria"
End Sub
'Bot�o btnLImgCadastroTipoMaterial tela cadastrar bloco
Private Sub btnLImgCadastroTipoMaterial_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar tipo material"
End Sub
'Bot�o btnLTxtCadastrarBloco tela cadastrar bloco
Private Sub btnLTxtCadastrarBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar bloco"
End Sub
'Bot�o btnLTextLimparCadastroBloco tela cadastrar bloco
Private Sub btnLTxtLimparCadastroBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o limpar campos tela cadastro de blocos"
End Sub

'-----------------------------------------------------------------TELA DESPACHE-----------------------------------
'                                                                 -------------
'Efeito de label nome do pdf tela despache
Private Sub lDigiteNomeArquivoExemplo_Click()
    lDigiteNomeArquivo.Visible = True
    lDigiteNomeArquivoExemplo.Visible = False
    txtNomeArquivo.SetFocus
End Sub
'Efeito e coloca em caixa alta o texto em txtNomeArquivo tela despache
Private Sub txtNomeArquivo_Change()
    lDigiteNomeArquivo.Visible = True
    lDigiteNomeArquivoExemplo.Visible = False
    
    If txtNomeArquivo.Value = "" Then
        lDigiteNomeArquivo.Visible = False
        lDigiteNomeArquivoExemplo.Visible = True
    End If
    
    txtNomeArquivo.Value = UCase(txtNomeArquivo.Value)
End Sub
'Efeito ao sair da caixa txtNomeArquivo de texto tela despache
Private Sub txtNomeArquivo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtNomeArquivo.Value = "" Then
        lDigiteNomeArquivo.Visible = False
        lDigiteNomeArquivoExemplo.Visible = True
    End If
End Sub
'Bot�o btnLImgCadastrarMotoristaDespache tela despache
Private Sub btnLImgCadastrarMotoristaDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastro motorista tela despache"
End Sub
'Efeito de passagem do mouse bot�o btnLImgCadastrarDestinoDespache tela despache
Private Sub btnLImgCadastrarDestinoDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastro destino tela despache"
End Sub
'Bot�o btnLTxtAdicionar tela despache
Private Sub btnLTxtAdicionar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o adicionar chapa tela despache"
End Sub
'Bot�o btnLImgExportarListaDespache tela despache
Private Sub btnLImgExportarListaDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o exportar pdf tela despache"
End Sub
'Bot�o btnLTxtDespachar tela despache
Private Sub btnLTxtDespachar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o despachar tela despache"
End Sub
'Bot�o btnLTxtLimparDespache tela despache
Private Sub btnLTxtLimparDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o limpar dados tela despache"
End Sub

'-----------------------------------------------------------------TELA ESTOQUE M�-----------------------------------
'                                                                 ---------------
'Efeito de label nome do pdf tela estoque m�
Private Sub lDigiteNomeArquivoM3Explemplo_Click()
    lDigiteNomeArquivoM3.Visible = True
    lDigiteNomeArquivoM3Explemplo.Visible = False
    txtNomeArquivoEstoqueBlocos.SetFocus
End Sub
'Efeito e coloca em caixa alta o texto em txtNomeArquivoEstoqueBlocos tela estoque m�
Private Sub txtNomeArquivoEstoqueBlocos_Change()
    lDigiteNomeArquivoM3.Visible = True
    lDigiteNomeArquivoM3Explemplo.Visible = False

    If txtNomeArquivoEstoqueBlocos.Value = "" Then
        lDigiteNomeArquivoM3.Visible = False
        lDigiteNomeArquivoM3Explemplo.Visible = True
    End If

    txtNomeArquivoEstoqueBlocos.Value = UCase(txtNomeArquivoEstoqueBlocos.Value)
End Sub
'Efeito ao sair da caixa txtNomeArquivoEstoqueBlocos de texto tela estoque m�
Private Sub txtNomeArquivoEstoqueBlocos_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtNomeArquivoEstoqueBlocos.Value = "" Then
        lDigiteNomeArquivoM3.Visible = False
        lDigiteNomeArquivoM3Explemplo.Visible = True
    End If
End Sub
'Bot�o btnLTxtPesquisarBlocos tela estoque m�
Private Sub btnLTxtPesquisarBlocos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o pesquiasr tela estoque m�"
End Sub
'Bot�o btnLTxtLimparFiltrosBlocos tela estoque m�
Private Sub btnLTxtLimparFiltrosBlocos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o limpar filtros tela estoque m�"
End Sub
'Bot�o btnLImgExportarEstoqueM3 tela estoque m�
Private Sub btnLImgExportarEstoqueM3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o exportar estoque m� tela estoque m�"
End Sub
'Bot�o btnLTxtNovoBloco tela estoque m�
Private Sub btnLTxtNovoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o novo bloco tela estoque m�"
End Sub
'Bot�o btnLTxtEditarBloco tela estoque m�
Private Sub btnLTxtEditarBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o editar bloco tela estoque m�"
End Sub
'Bot�o btnLTxtADDEstoque tela estoque m�
Private Sub btnLTxtADDEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o adicionar chapas ao estoque tela estoque m�"
End Sub

'-----------------------------------------------------------------TELA ESTOQUE M�-----------------------------------
'                                                                 ---------------
'Efeito de label nome do pdf tela estoque m�
Private Sub lDigiteNomeArquivoM2Explemplo_Click()
    lDigiteNomeArquivoM2.Visible = True
    lDigiteNomeArquivoM2Explemplo.Visible = False
    txtNomeArquivoEstoqueChapas.SetFocus
End Sub
'Efeito e coloca em caixa alta o texto em txtNomeArquivoEstoqueChapas tela estoque m�
Private Sub txtNomeArquivoEstoqueChapas_Change()
    lDigiteNomeArquivoM2.Visible = True
    lDigiteNomeArquivoM2Explemplo.Visible = False

    If txtNomeArquivoEstoqueChapas.Value = "" Then
        lDigiteNomeArquivoM2.Visible = False
        lDigiteNomeArquivoM2Explemplo.Visible = True
    End If

    txtNomeArquivoEstoqueChapas.Value = UCase(txtNomeArquivoEstoqueChapas.Value)
End Sub
'Efeito ao sair da caixa txtNomeArquivoEstoqueChapas de texto tela estoque m�
Private Sub txtNomeArquivoEstoqueChapas_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtNomeArquivoEstoqueChapas.Value = "" Then
        lDigiteNomeArquivoM2.Visible = False
        lDigiteNomeArquivoM2Explemplo.Visible = True
    End If
End Sub
'Bot�o btnLTxtPesquisarChapas tela estoque m�
Private Sub btnLTxtPesquisarChapas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o pesquisar chapa tela estoque m�"
End Sub
'Bot�o btnLTxtLimparFiltrosChapas tela estoque m�
Private Sub btnLTxtLimparFiltrosChapas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o limpar filtros tela estoque m�"
End Sub
'Bot�o btnLImgExportarEstoqueM2 tela estoque m�
Private Sub btnLImgExportarEstoqueM2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o esportar estoque m� tela estoque m�"
End Sub
'Bot�o btnLTxtNovoAvulso tela estoque m�
Private Sub btnLTxtNovoAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar avulso tela estoque m�"
End Sub
'Bot�o btnLTxtEditarChapa tela estoque m�
Private Sub btnLTxtEditarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o editar chapa tela estoque m�"
End Sub
'Bot�o btnLTxtTrocaEstoque tela estoque m�
Private Sub btnLTxtTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o troca material no estoque tela estoque m�"
End Sub

'-----------------------------------------------------------------TELA CADASTRO AVULSO-----------------------------------
'                                                                 --------------------
'Bot�o btnLImgCadastrarMaterialAvulso tela cadastro avulso
Private Sub btnLImgCadastrarMaterialAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar tipo material tela cadastro avulso"
End Sub
'Bot�o btnLImgCadastrarPolimentoAvulso tela cadastro avulso
Private Sub btnLImgCadastrarPolimentoAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar tipo polimento tela cadastro avulso"
End Sub
'Bot�o btnLTxtCadastrarChapaAvulso tela cadastro avulso
Private Sub btnLTxtCadastrarChapaAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar chapas avulsos tela cadastro avulso"
End Sub
'Bot�o btnLTxtLimparCadastroChapaAvulso tela cadastro avulso
Private Sub btnLTxtLimparCadastroChapaAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o limpar cadastro chapas avulsos tela cadastro avulso"
End Sub

'-----------------------------------------------------------------TELA CARREGOS-----------------------------------
'                                                                 -------------
'Bot�o btnLTxtPesquisarCarregos tela carregos
Private Sub btnLTxtPesquisarCarregos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o pesquisar por carregos tela carregos"
End Sub
'Bot�o btnLTxtLimparListas tela carregos
Private Sub btnLTxtLimparListas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o limpar dados filtro tela carregos"
End Sub
'Bot�o btnLImgExportarCarregoPDF tela carregos
Private Sub btnLImgExportarCarregoPDF_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o exportar carregos em pdf tela carregos"
End Sub
'Bot�o btnLTxtEditarCarrego tela carregos
Private Sub btnLTxtEditarCarrego_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o editar carrego tela carregos"
End Sub
'Bot�o btnLTxtVoltarCArrego tela carregos
Private Sub btnLTxtVoltarCArrego_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o voltar tela carregos"
End Sub
