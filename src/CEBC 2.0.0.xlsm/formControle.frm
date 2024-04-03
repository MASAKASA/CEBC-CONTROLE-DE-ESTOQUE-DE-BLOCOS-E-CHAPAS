VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formControle 
   Caption         =   "CONTROLE DE BLOCOS E CHAPAS 2.0.0"
   ClientHeight    =   13410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   24675
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
'Inicialização do formControle
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
    Me.MultiPageCEBC.Value = 11
End Sub
'Efeito para clique nas label btnLMenuCadastros do menu
Private Sub btnLMenuCadastros_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 1
End Sub
'Efeito para clique nas label btnLMenuLogs do menu
Private Sub btnLMenuLogs_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 12
End Sub

'-----------------------------------------------------------------TELA CADASTRO DE BLOCOS-----------------------------------
'                                                                 -----------------------
'Botão btnLImgCadastrarPedreira tela cadastrar bloco
Private Sub btnLImgCadastrarPedreira_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar pedreira"
End Sub
'Botão btnLImgCadastrarSerrariaCB tela cadastrar bloco
Private Sub btnLImgCadastrarSerrariaCB_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar serraria"
End Sub
'Botão btnLImgCadastroTipoMaterial tela cadastrar bloco
Private Sub btnLImgCadastroTipoMaterial_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar tipo material"
End Sub
'Botão btnLTxtCadastrarBloco tela cadastrar bloco
Private Sub btnLTxtCadastrarBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar bloco"
End Sub
'Botão btnLTextLimparCadastroBloco tela cadastrar bloco
Private Sub btnLTxtLimparCadastroBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço limpar campos tela cadastro de blocos"
End Sub

'-----------------------------------------------------------------TELA DESPACHE-----------------------------------
'                                                                 -------------
'Botão btnLImgCadastrarMotoristaDespache tela despache
Private Sub btnLImgCadastrarMotoristaDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastro motorista tela despache"
End Sub
'Botão btnLImgCadastrarDestinoDespache tela despache
Private Sub btnLImgCadastrarDestinoDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastro destino tela despache"
End Sub
'Botão btnLTxtAdicionar tela despache
Private Sub btnLTxtAdicionar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço adicionar chapa tela despache"
End Sub
'Botão btnLTxtDespachar tela despache
Private Sub btnLTxtDespachar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço despachar tela despache"
End Sub
'Botão btnLTxtLimparDespache tela despache
Private Sub btnLTxtLimparDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço limpar dados tela despache"
End Sub

'-----------------------------------------------------------------TELA ESTOQUE M³-----------------------------------
'                                                                 ---------------
'Efeito de label nome do pdf tela estoque m³
Private Sub lDigiteNomeArquivoM3Explemplo_Click()
    lDigiteNomeArquivoM3.Visible = True
    lDigiteNomeArquivoM3Explemplo.Visible = False
    txtNomeArquivoEstoqueBlocos.SetFocus
End Sub
'Efeito e coloca em caixa alta o texto em txtNomeArquivoEstoqueBlocos tela estoque m³
Private Sub txtNomeArquivoEstoqueBlocos_Change()
    lDigiteNomeArquivoM3.Visible = True
    lDigiteNomeArquivoM3Explemplo.Visible = False

    If txtNomeArquivoEstoqueBlocos.Value = "" Then
        lDigiteNomeArquivoM3.Visible = False
        lDigiteNomeArquivoM3Explemplo.Visible = True
    End If

    txtNomeArquivoEstoqueBlocos.Value = UCase(txtNomeArquivoEstoqueBlocos.Value)
End Sub
'Efeito ao sair da caixa txtNomeArquivoEstoqueBlocos de texto tela estoque m³
Private Sub txtNomeArquivoEstoqueBlocos_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtNomeArquivoEstoqueBlocos.Value = "" Then
        lDigiteNomeArquivoM3.Visible = False
        lDigiteNomeArquivoM3Explemplo.Visible = True
    End If
End Sub
'Botão btnLTxtPesquisarBlocos tela estoque m³
Private Sub btnLTxtPesquisarBlocos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço pesquiasr tela estoque m³"
End Sub
'Botão btnLTxtLimparFiltrosBlocos tela estoque m³
Private Sub btnLTxtLimparFiltrosBlocos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço limpar filtros tela estoque m³"
End Sub
'Botão btnLImgExportarEstoqueM3 tela estoque m³
Private Sub btnLImgExportarEstoqueM3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço exportar estoque m³ tela estoque m³"
End Sub
'Botão btnLTxtNovoBloco tela estoque m³
Private Sub btnLTxtNovoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço novo bloco tela estoque m³"
End Sub
'Botão btnLTxtEditarBloco tela estoque m³
Private Sub btnLTxtEditarBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço editar bloco tela estoque m³"
End Sub
'Botão btnLTxtADDEstoque tela estoque m³
Private Sub btnLTxtADDEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço adicionar chapas ao estoque tela estoque m³"
End Sub

'-----------------------------------------------------------------TELA ESTOQUE M²-----------------------------------
'                                                                 ---------------
'Efeito de label nome do pdf tela estoque m²
Private Sub lDigiteNomeArquivoM2Explemplo_Click()
    lDigiteNomeArquivoM2.Visible = True
    lDigiteNomeArquivoM2Explemplo.Visible = False
    txtNomeArquivoEstoqueChapas.SetFocus
End Sub
'Efeito e coloca em caixa alta o texto em txtNomeArquivoEstoqueChapas tela estoque m²
Private Sub txtNomeArquivoEstoqueChapas_Change()
    lDigiteNomeArquivoM2.Visible = True
    lDigiteNomeArquivoM2Explemplo.Visible = False

    If txtNomeArquivoEstoqueChapas.Value = "" Then
        lDigiteNomeArquivoM2.Visible = False
        lDigiteNomeArquivoM2Explemplo.Visible = True
    End If

    txtNomeArquivoEstoqueChapas.Value = UCase(txtNomeArquivoEstoqueChapas.Value)
End Sub
'Efeito ao sair da caixa txtNomeArquivoEstoqueChapas de texto tela estoque m²
Private Sub txtNomeArquivoEstoqueChapas_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtNomeArquivoEstoqueChapas.Value = "" Then
        lDigiteNomeArquivoM2.Visible = False
        lDigiteNomeArquivoM2Explemplo.Visible = True
    End If
End Sub
'Botão btnLTxtPesquisarChapas tela estoque m²
Private Sub btnLTxtPesquisarChapas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço pesquisar chapa tela estoque m²"
End Sub
'Botão btnLTxtLimparFiltrosChapas tela estoque m²
Private Sub btnLTxtLimparFiltrosChapas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço limpar filtros tela estoque m²"
End Sub
'Botão btnLImgExportarEstoqueM2 tela estoque m²
Private Sub btnLImgExportarEstoqueM2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço esportar estoque m² tela estoque m²"
End Sub
'Botão btnLTxtNovoAvulso tela estoque m²
Private Sub btnLTxtNovoAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar avulso tela estoque m²"
End Sub
'Botão btnLTxtEditarChapa tela estoque m²
Private Sub btnLTxtEditarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço editar chapa tela estoque m²"
End Sub
'Botão btnLTxtTrocaEstoque tela estoque m²
Private Sub btnLTxtTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço troca material no estoque tela estoque m²"
End Sub

'-----------------------------------------------------------------TELA CADASTRO AVULSO-----------------------------------
'                                                                 --------------------
'Botão btnLImgCadastrarMaterialAvulso tela cadastro avulso
Private Sub btnLImgCadastrarMaterialAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar tipo material tela cadastro avulso"
End Sub
'Botão btnLImgCadastrarPolimentoAvulso tela cadastro avulso
Private Sub btnLImgCadastrarPolimentoAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar tipo polimento tela cadastro avulso"
End Sub
'Botão btnLTxtCadastrarChapaAvulso tela cadastro avulso
Private Sub btnLTxtCadastrarChapaAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar chapas avulsos tela cadastro avulso"
End Sub
'Botão btnLTxtLimparCadastroChapaAvulso tela cadastro avulso
Private Sub btnLTxtLimparCadastroChapaAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço limpar cadastro chapas avulsos tela cadastro avulso"
End Sub

'-----------------------------------------------------------------TELA CARREGOS-----------------------------------
'                                                                 -------------
'Botão btnLTxtPesquisarCarregos tela carregos
Private Sub btnLTxtPesquisarCarregos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço pesquisar por carregos tela carregos"
End Sub
'Botão btnLTxtLimparListas tela carregos
Private Sub btnLTxtLimparListas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço limpar dados filtro tela carregos"
End Sub
'Botão btnLImgExportarCarregoPDF tela carregos
Private Sub btnLImgExportarCarregoPDF_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço exportar carregos em pdf tela carregos"
End Sub
'Botão btnLTxtEditarCarrego tela carregos
Private Sub btnLTxtEditarCarrego_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço editar carrego tela carregos"
End Sub
'Botão btnLTxtVoltarCArrego tela carregos
Private Sub btnLTxtVoltarCArrego_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço voltar tela carregos"
End Sub

'-----------------------------------------------------------------TELA EDITAR BLOCO-----------------------------------
'                                                                 -----------------
'Botão btnLTxtSalvarEdicaoBloco tela carregos
Private Sub btnLTxtSalvarEdicaoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço editar bloco tela editar bloco"
End Sub
'Botão btnLTxtVoltarEdicaoBloco tela carregos
Private Sub btnLTxtVoltarEdicaoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço voltar tela editar bloco"
End Sub

'-----------------------------------------------------------------TELA LANÇAMENTO E EDIÇÃO CHAPA-----------------------------------
'                                                                 ------------------------------
'Botão btnLImgCadastrarPolideiraChapa tela lançamento e edição chapa
Private Sub btnLImgCadastrarPolideiraChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar Polideira tela lançamento e edição chapa"
End Sub
'Botão btnLImgCadastrarTipoPolideiraChapa tela lançamento e edição chapa
Private Sub btnLImgCadastrarTipoPolideiraChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar tipo polimento tela lançamento e edição chapa"
End Sub
'Botão btnLImgCadastrarTipoMaterialChapa tela lançamento e edição chapa
Private Sub btnLImgCadastrarTipoMaterialChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço tipo material tela lançamento e edição chapa"
End Sub
'Botão btnLImgCadastrarTipoMaterialChapaTamanhos tela lançamento e edição chapa
Private Sub btnLImgCadastrarTipoMaterialChapaTamanhos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço tipo material tela lançamento e edição chapa"
End Sub
'Botão btnLTxtAdicionarTamanhoChapa tela lançamento e edição chapa
Private Sub btnLTxtAdicionarTamanhoChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço adicionar tamanhos tela lançamento e edição chapa"
End Sub
'Botão btnLTxtEditarTamanhoChapa tela lançamento e edição chapa
Private Sub btnLTxtEditarTamanhoChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço editar tamanho chapa tela lançamento e edição chapa"
End Sub
'Botão btnLTxtTirarDaLista tela lançamento e edição chapa
Private Sub btnLTxtTirarDaLista_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço tira tamanho da lista tela lançamento e edição chapa"
End Sub
'Botão btnLTxtSalvarChapa tela lançamento e edição chapa
Private Sub btnLTxtSalvarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço salva alteração da chapa tela lançamento e edição chapa"
End Sub
'Botão btnLTxtVoltarChapa tela lançamento e edição chapa
Private Sub btnLTxtVoltarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço voltar tela lançamento e edição chapa"
End Sub


'-----------------------------------------------------------------TELA TROCA ESTOQUE-----------------------------------
'                                                                 ------------------
'Botão btnLTxtAdicionarTrocaEstoque tela troca estoque
Private Sub btnLTxtAdicionarTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço adicionar material para troca tela troca estoque"
End Sub
'Botão btnLTxtTrocarEstoque tela troca estoque
Private Sub btnLTxtTrocarEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço troca de estoque tela troca estoque"
End Sub
'Botão btnLTxtVoltarTrocaEstoque tela troca estoque
Private Sub btnLTxtVoltarTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço voltar tela troca estoque"
End Sub


'-----------------------------------------------------------------TELA CADASTROS DIVERSOS-----------------------------------
'                                                                 -----------------------
'Botão btnLTxtSalvarPedreira tela cadastros diversos
Private Sub btnLTxtSalvarPedreira_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar pedreira tela cadastros diversos"
End Sub
'Botão btnLTxtSalvarSerraria tela cadastros diversos
Private Sub btnLTxtSalvarSerraria_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar serraria tela cadastros diversos"
End Sub
'Botão btnLTxtSalvarPolideira tela cadastros diversos
Private Sub btnLTxtSalvarPolideira_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar polideira tela cadastros diversos"
End Sub
'Botão btnLTxtSalvarTipoMaterial tela cadastros diversos
Private Sub btnLTxtSalvarTipoMaterial_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar tipo material tela cadastros diversos"
End Sub
'Botão btnLTxtSalvarTipoPolimento tela cadastros diversos
Private Sub btnLTxtSalvarTipoPolimento_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar tipo polimento tela cadastros diversos"
End Sub
'Botão btnLTxtSalvarMotorista tela cadastros diversos
Private Sub btnLTxtSalvarMotorista_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar motorista tela cadastros diversos"
End Sub
'Botão btnLTxtSalvarMotorista tela cadastros diversos
Private Sub btnLTxtSalvarDestino_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar destino tela cadastros diversos"
End Sub
