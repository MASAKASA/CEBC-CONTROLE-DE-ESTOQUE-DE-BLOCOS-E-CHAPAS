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

' Variaveis para manipulação com os botões e frames
Dim botoesMenu() As clsLabel
Dim botoesImg() As clsLabel
Dim botoesText() As clsLabel
Dim frameEfeito() As clsFrame

' Variaveis para manipulação
Dim status() As String
Dim listaObjeto As Collection

' Variaveis de objetos
Dim bloco As objBloco
Dim chapa As objChapa
Dim pedreira As objPedreira
Dim polideira As objPolideira
Dim serraria As objSerraria

' Inicialização do formControle
Private Sub UserForm_Initialize()
    ' Variaveis para o metodo
    Dim obj As Object
    Dim i As Long
    Dim j As Long
    Dim l As Long
    Dim m As Long
    Dim nameObj As String
    Dim nameObjInicio As String
    
    ' Resevando espaço em memoria para manipulação das variaveis
    ReDim botoesMenu(1 To Me.Controls.Count)
    ReDim botoesImg(1 To Me.Controls.Count)
    ReDim botoesText(1 To Me.Controls.Count)
    ReDim frameEfeito(1 To Me.Controls.Count)
    ReDim status(1 To 6)
    
    ' Atribuições da variaveis
    status(1) = "PEDREIRA"
    status(2) = "SERRADA"
    status(3) = "ESTOQUE"
    status(4) = "FECHADO"
    status(5) = "CHAPAS BRUTAS"
    status(6) = "EM PROCESSO"
    
    ' Separa os botões e frames
    For Each obj In Me.Controls
        
        ' Atribuições das variaveis para manipulações
        nameObj = obj.name
        nameObjInicio = Mid(nameObj, 1, 7)
        
        ' Captura os botões no menu
        If nameObjInicio = "btnLMen" Then
            i = i + 1
            Set botoesMenu(i) = New clsLabel
            Set botoesMenu(i).efeitoBotoesMenu = obj
        End If
        
        ' Captura os botões com imagens
        If nameObjInicio = "btnLImg" Then
            j = j + 1
            Set botoesImg(j) = New clsLabel
            Set botoesImg(j).efeitoBotoesImagem = obj
        End If
        
        ' Captura os botões com textos
        If nameObjInicio = "btnLTxt" Then
            l = l + 1
            Set botoesText(l) = New clsLabel
            Set botoesText(l).efeitoBotoesTexto = obj
        End If
        
        ' Captura os frames para efeitos com botões
        If nameObjInicio = "fTiraEf" Then
            m = m + 1
            Set frameEfeito(m) = New clsFrame
            Set frameEfeito(m).efeitoFrame = obj
        End If
    Next obj
    
    ' Limpando a variavel
    Set obj = Nothing
    
    ' Redefinição dos espaço em memoria das variaveis
    ReDim Preserve botoesMenu(1 To i)
    ReDim Preserve botoesImg(1 To j)
    ReDim Preserve botoesText(1 To l)
    ReDim Preserve frameEfeito(1 To m)
        
    ' Retira os nomes de cima da multPage
    Me.MultiPageCEBC.Style = fmTabStyleNone
End Sub

'-----------------------------------------------------------------MENU DO SISTEMA-----------------------------------
'                                                                 ---------------
' Efeito para clique nas label btnLMenuHome do menu
Private Sub btnLMenuHome_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 0
End Sub
' Efeito para clique nas label btnLMenuBloco do menu
Private Sub btnLMenuBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ' Carregar os comboBox da tela
    Call carregarPedreiras(Me.cbPedreiraBlocoPesquisa)
    Call carregarSerrarias(Me.cbSerrariaBlocoPesquisa)
    Call carregarTemNota(Me.cbTemNota)
    
    ' Carregar a list
    Call carregarList(Me.ListEstoqueM3)
    
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 1
End Sub
' Efeito para clique nas label btnLMenuChapa do menu
Private Sub btnLMenuChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Carregar os comboBox da tela
    Call carregarPolideiras(Me.cbPolideiraChapaPesquisa)
    Call carregarTiposPolimento(Me.cbTipoPolimentoPesquisa)
    
    ' Carregar a list
    Call carregarList(Me.ListEstoqueChapas)
    
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 4
End Sub
' Efeito para clique nas label btnLMenuDespachar do menu
Private Sub btnLMenuDespachar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 8
End Sub
' Efeito para clique nas label btnLMenuCarrago do menu
Private Sub btnLMenuCarrago_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 9
End Sub
' Efeito para clique nas label btnLMenuCadastros do menu
Private Sub btnLMenuCadastros_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 10
End Sub
' Efeito para clique nas label btnLMenuUsuarios do menu
Private Sub btnLMenuUsuarios_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 11
End Sub

'-----------------------------------------------------------------TELA ESTOQUE M³-----------------------------------
'                                                                 ---------------
' Efeito de label nome do pdf tela estoque m³
Private Sub lDigiteNomeArquivoM3Explemplo_Click()
    lDigiteNomeArquivoM3.Visible = True
    lDigiteNomeArquivoM3Explemplo.Visible = False
    txtNomeArquivoEstoqueBlocos.SetFocus
End Sub
' Efeito e coloca em caixa alta o texto em txtNomeArquivoEstoqueBlocos tela estoque m³
Private Sub txtNomeArquivoEstoqueBlocos_Change()
    lDigiteNomeArquivoM3.Visible = True
    lDigiteNomeArquivoM3Explemplo.Visible = False

    If txtNomeArquivoEstoqueBlocos.Value = "" Then
        lDigiteNomeArquivoM3.Visible = False
        lDigiteNomeArquivoM3Explemplo.Visible = True
    End If

    txtNomeArquivoEstoqueBlocos.Value = UCase(txtNomeArquivoEstoqueBlocos.Value)
End Sub
' Efeito ao sair da caixa txtNomeArquivoEstoqueBlocos de texto tela estoque m³
Private Sub txtNomeArquivoEstoqueBlocos_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtNomeArquivoEstoqueBlocos.Value = "" Then
        lDigiteNomeArquivoM3.Visible = False
        lDigiteNomeArquivoM3Explemplo.Visible = True
    End If
End Sub
' txtDataInicioBlocoPesquisa tela estoque m³
Private Sub txtDataInicioBlocoPesquisa_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Deixa só a digitação de numero
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
    ' Coloca as barras para formatação
    If Len(txtDataInicioBlocoPesquisa.Value) = 2 Or Len(txtDataInicioBlocoPesquisa.Value) = 5 Then
    
        txtDataInicioBlocoPesquisa.Value = txtDataInicioBlocoPesquisa.Value & "/"
    End If
End Sub
' txtDataFinalBlocoPesquisa tela estoque m³
Private Sub txtDataFinalBlocoPesquisa_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Deixa só a digitação de numero
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
    ' Coloca as barras para formatação
    If Len(txtDataFinalBlocoPesquisa.Value) = 2 Or Len(txtDataFinalBlocoPesquisa.Value) = 5 Then
    
        txtDataFinalBlocoPesquisa.Value = txtDataFinalBlocoPesquisa.Value & "/"
    End If
End Sub
' txtMaterialBlocoPesquisa tela estoque m³
Private Sub txtMaterialBlocoPesquisa_Change()
    'Coloca tudo em caixa alta
    txtMaterialBlocoPesquisa.Value = UCase(txtMaterialBlocoPesquisa.Value)
End Sub
' txtIdBlocoPesquisa tela estoque m³
Private Sub txtIdBlocoPesquisa_Change()
    'Coloca tudo em caixa alta
    txtIdBlocoPesquisa.Value = UCase(txtIdBlocoPesquisa.Value)
End Sub
' Atelho para seleção dos status, obPedreiraESerrada tela estoque m³
Private Sub obPedreiraESerrada_Click()
    chbPedreida.Value = True
    chbSerraria.Value = True
    chbChapasBrutas.Value = False
    chbEmProcesso.Value = False
    chbEstoque.Value = False
    chbFechado.Value = False
End Sub
' Atelho para seleção dos status, obEmEstoque tela estoque m³
Private Sub obEmEstoque_Click()
    chbPedreida.Value = False
    chbSerraria.Value = False
    chbChapasBrutas.Value = True
    chbEmProcesso.Value = True
    chbEstoque.Value = True
    chbFechado.Value = False
End Sub
' Atelho para seleção dos status, obFechado tela estoque m³
Private Sub obFechado_Click()
    chbPedreida.Value = False
    chbSerraria.Value = False
    chbChapasBrutas.Value = False
    chbEmProcesso.Value = False
    chbEstoque.Value = False
    chbFechado.Value = True
End Sub
' Atelho para seleção dos status, opPedreiraSerradaEmProcesso tela estoque m³
Private Sub opPedreiraSerradaEmProcesso_Click()
    chbPedreida.Value = True
    chbSerraria.Value = True
    chbChapasBrutas.Value = True
    chbEmProcesso.Value = True
    chbEstoque.Value = True
    chbFechado.Value = False
End Sub
' Atelho para seleção dos status, opTodos tela estoque m³
Private Sub opTodos_Click()
    chbPedreida.Value = True
    chbSerraria.Value = True
    chbChapasBrutas.Value = True
    chbEmProcesso.Value = True
    chbEstoque.Value = True
    chbFechado.Value = True
End Sub
' Botão btnLTxtPesquisarBlocos tela estoque m³
Private Sub btnLTxtPesquisarBlocos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call daoBloco.listarBlocosFilter
End Sub
' Botão btnLTxtLimparFiltrosBlocos tela estoque m³
Private Sub btnLTxtLimparFiltrosBlocos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    Call limparCamposPesquisaEstoqueM3
End Sub
' Botão btnLImgExportarEstoqueM3 tela estoque m³
Private Sub btnLImgExportarEstoqueM3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço exportar estoque m³, tela estoque m³"
End Sub
' Botão btnLTxtNovoBloco tela estoque m³
Private Sub btnLTxtNovoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Coloca data atual na txtDataCadastro na tela cadastro de bloco
    txtDataCadastro.Value = Date
    
    ' Chama metodo para carregar comboBox
    Call carregarPedreiras(Me.cbPedreira)
    Call carregarSerrarias(Me.cbSerrariaCB)
    Call carregarTiposMateriais(Me.cbTipoMaterial)
    Call carregarTemNota(Me.cbNotaC)
    
    ' Chama metodo para carregar lista e blocos cadastros do dia atual
    Call carregarList(Me.listCadastradosHoje)
    
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 2
End Sub
' Botão btnLTxtEditarBloco tela estoque m³
Private Sub btnLTxtEditarBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage para tela editar bloco
    Me.MultiPageCEBC.Value = 3
    
    ' Chama serviço para pesquisa do bloco
    Set listaObjeto = daoBloco.pesquisarPorId("01") ' Me.ListEstoqueM3.list(Me.ListEstoqueM3.ListIndex, 0)) ' Envia o id do bloco
    
    ' Carregar os comboBox da tela
    Call carregarTiposMateriais(Me.cbTipoMaterialEditar)
    Call carregarPedreiras(Me.cbPedreiraEditar)
    Call carregarSerrarias(Me.cbSerrariaEditar)
    Call carregarPolideiras(Me.cbPolideiraEditar)
    Call carregarEstoque(Me.cbEstoqueEditar)
    Call carregarTemNota(Me.cbNotaBlocoEditar)
    Call carregarCustoMedio(Me.cbCustoMedioEditar)
    
    ' Carrega os dados na tela editar bloco
    Call carregarDadosBlocoTelaEdicaoBloco(listaObjeto)
End Sub
' Botão btnLTxtADDEstoque tela estoque m³
Private Sub btnLTxtADDEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ' Botão chapa
    formControle.Controls("btnLMenuChapa").BackColor = RGB(200, 230, 255)
    formControle.Controls("btnLMenuChapa").Font.Size = 32
    formControle.Controls("btnLMenuChapa").Font.Size = 20
    formControle.Controls("btnLMenuChapa").Left = 15
    formControle.Controls("btnLMenuChapa").Width = 172
    formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignCenter
                
    ' Botão Menu
    formControle.Controls("btnLMenuBloco").BackColor = RGB(0, 100, 200)
    formControle.Controls("btnLMenuBloco").Left = 2
    formControle.Controls("btnLMenuBloco").Width = 189
    formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignLeft
    
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 6
End Sub

'-----------------------------------------------------------------TELA CADASTRO DE BLOCOS-----------------------------------
'                                                                 -----------------------
' txtIdBloco tela cadastro de bloco
Private Sub txtIdBloco_Change()
    ' Coloca tudo em caixa alta
    txtIdBloco.Value = UCase(txtIdBloco.Value)
    
    ' Cria o código para o sistema
    txtIdBlocoSistema.Value = txtIdBloco & "-" & Util.ExtrairUltimaPalavra(txtNomeBloco.Value) & "-BL"
    
    ' Deixa em branco o codigo se as variaveis forem vazias
    If txtIdBloco.Value = "" And txtNomeBloco.Value = "" Then
        txtIdBlocoSistema.Value = ""
    End If
End Sub
' txtNomeBloco tela cadastro de bloco
Private Sub txtNomeBloco_Change()
    ' Coloca tudo em caixa alta
    txtNomeBloco.Value = UCase(txtNomeBloco.Value)
    
    ' Cria o código para o sistema
    txtIdBlocoSistema.Value = txtIdBloco & "-" & Util.ExtrairUltimaPalavra(txtNomeBloco.Value) & "-BL"
    
    ' Deixa em branco o codigo se as variaveis forem vazias
    If txtIdBloco.Value = "" And txtNomeBloco.Value = "" Then
        txtIdBlocoSistema.Value = ""
    End If
End Sub
' txtNomeBloco tela cadastro de bloco
Private Sub txtObsBlocoCB_Change()
    ' Coloca tudo em caixa alta
    txtObsBlocoCB.Value = UCase(txtObsBlocoCB.Value)
End Sub
' txtComprimentoBloco tela cadastro de bloco
Private Sub txtComprimentoBloco_Change()
    ' Define o resultado no TextBox
    txtComprimentoBloco.Value = Util.formatarMetros(txtComprimentoBloco.Value)
    
    ' Seta o valor no comprimento bruto
    txtCompBrutoBloco.Value = txtComprimentoBloco.Value
    
    ' Move o cursor para o final do TextBox
    txtComprimentoBloco.SelStart = Len(txtComprimentoBloco.Value)
    
    ' Retorna valor calculado e formatado
    txtTotalM3.Value = Util.formatarComPontos(Format(Util.calcularM3(txtComprimentoBloco.Value, _
            txtAlturaBloco.Value, txtLarguraBloco.Value), "0.0000"))

    ' Retorna valor calculado e formatado
    txtValorTotalBloco.Value = Util.formatarComPontos(Format(Util.calcularValorBloco(txtValorM3.Value, _
            txtTotalM3.Value), "0.00"))
End Sub
' txtAlturaBloco tela cadastro de bloco
Private Sub txtAlturaBloco_Change()
    ' Define o resultado no TextBox
    txtAlturaBloco.Value = Util.formatarMetros(txtAlturaBloco.Value)
    
    ' Seta o valor na altura bruto
    txtAlturaBlocoBruto.Value = txtAlturaBloco.Value
    
    ' Move o cursor para o final do TextBox
    txtAlturaBloco.SelStart = Len(txtAlturaBloco.Value)
    
    ' Retorna valor calculado e formatado
    txtTotalM3.Value = Util.formatarComPontos(Format(Util.calcularM3(txtComprimentoBloco.Value, _
            txtAlturaBloco.Value, txtLarguraBloco.Value), "0.0000"))

    ' Retorna valor calculado e formatado
    txtValorTotalBloco.Value = Util.formatarComPontos(Format(Util.calcularValorBloco(txtValorM3.Value, _
            txtTotalM3.Value), "0.00"))
End Sub
' txtLarguraBloco tela cadastro de bloco
Private Sub txtLarguraBloco_Change()
    ' Define o resultado no TextBox
    txtLarguraBloco.Value = Util.formatarMetros(txtLarguraBloco.Value)
    
    ' Seta o valor na altura bruto
    txtLarguraBlocoBruto.Value = txtLarguraBloco.Value
    
    ' Move o cursor para o final do TextBox
    txtLarguraBloco.SelStart = Len(txtLarguraBloco.Value)
    
    ' Retorna valor calculado e formatado
    txtTotalM3.Value = Util.formatarComPontos(Format(Util.calcularM3(txtComprimentoBloco.Value, _
            txtAlturaBloco.Value, txtLarguraBloco.Value), "0.0000"))

    ' Retorna valor calculado e formatado
    txtValorTotalBloco.Value = Util.formatarComPontos(Format(Util.calcularValorBloco(txtValorM3.Value, _
            txtTotalM3.Value), "0.00"))
End Sub
' txtCompBrutoBloco tela cadastro de bloco
Private Sub txtCompBrutoBloco_Change()
    ' Define o resultado no TextBox
    txtCompBrutoBloco.Value = Util.formatarMetros(txtCompBrutoBloco.Value)
    
    ' Move o cursor para o final do TextBox
    txtCompBrutoBloco.SelStart = Len(txtCompBrutoBloco.Value)
End Sub
' txtAlturaBlocoBruto tela cadastro de bloco
Private Sub txtAlturaBlocoBruto_Change()
    ' Define o resultado no TextBox
    txtAlturaBlocoBruto.Value = Util.formatarMetros(txtAlturaBlocoBruto.Value)
    
    ' Move o cursor para o final do TextBox
    txtAlturaBlocoBruto.SelStart = Len(txtAlturaBlocoBruto.Value)
End Sub
' txtLarguraBlocoBruto tela cadastro de bloco
Private Sub txtLarguraBlocoBruto_Change()
    ' Define o resultado no TextBox
    txtLarguraBlocoBruto.Value = Util.formatarMetros(txtLarguraBlocoBruto.Value)
    
    ' Move o cursor para o final do TextBox
    txtLarguraBlocoBruto.SelStart = Len(txtLarguraBlocoBruto.Value)
End Sub
' txtAdicionais tela cadastro de bloco
Private Sub txtAdicionais_Change()
    ' Define o resultado no TextBox
    txtAdicionais.Value = Util.formatarValor(txtAdicionais.Value)
    
    ' Move o cursor para o final do TextBox
    txtAdicionais.SelStart = Len(txtAdicionais.Value)
End Sub
' txtValorFreteBloco tela cadastro de bloco
Private Sub txtValorFreteBloco_Change()
    ' Define o resultado no TextBox
    txtValorFreteBloco.Value = Util.formatarValor(txtValorFreteBloco.Value)

    ' Move o cursor para o final do TextBox
    txtValorFreteBloco.SelStart = Len(txtValorFreteBloco.Value)
End Sub
' txtValorM3 tela cadastro de bloco
Private Sub txtValorM3_Change()
    ' Define o resultado no TextBox
    txtValorM3.Value = Util.formatarValor(txtValorM3.Value)
    
    ' Move o cursor para o final do TextBox
    txtValorM3.SelStart = Len(txtValorM3.Value)

    ' Retorna valor calculado e formatado
    txtValorTotalBloco.Value = Util.formatarComPontos(Format(Util.calcularValorBloco(txtValorM3.Value, _
            txtTotalM3.Value), "0.00"))
End Sub
' Botão btnLImgCadastrarPedreira tela cadastrar bloco
Private Sub btnLImgCadastrarPedreira_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar pedreira, tela cadastrar bloco"
End Sub
' Botão btnLImgCadastrarSerrariaCB tela cadastrar bloco
Private Sub btnLImgCadastrarSerrariaCB_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar serraria, tela cadastrar bloco"
End Sub
'Botão btnLImgCadastroTipoMaterial tela cadastrar bloco
Private Sub btnLImgCadastroTipoMaterial_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço cadastrar tipo material, tela cadastrar bloco"
End Sub
' Botão btnLTxtCadastrarBloco tela cadastrar bloco
Private Sub btnLTxtCadastrarBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Variaveis do medoto
    Dim nomeStatus As String
    
    ' Captura do status
    If obPedreiraCB.Value = True Then
        nomeStatus = status(1)
    Else
        nomeStatus = status(2)
    End If
    
    ' Criação e atribuição do objeto pedreira
    Set pedreira = New objPedreira
    pedreira.carregarPedreira cbPedreira.Value
    
    ' Criação e atribuição do objeto serraria
    Set serraria = New objSerraria
    serraria.carregarSerraria cbSerrariaCB.Value
    
    ' Criação e atribuição do objeto bloco
    Set bloco = New objBloco
    bloco.carregarBlocoCadastro txtDataCadastro.Value, txtIdBlocoSistema.Value, pedreira, serraria, txtIdBloco.Value, _
                                txtNomeBloco.Value, cbTipoMaterial.Value, cbNotaC.Value, nomeStatus, txtObsBlocoCB.Value, _
                                txtCompBrutoBloco.Value, txtAlturaBlocoBruto.Value, txtLarguraBlocoBruto.Value, _
                                txtComprimentoBloco.Value, txtAlturaBloco.Value, txtLarguraBloco.Value, txtAdicionais.Value, _
                                txtValorFreteBloco.Value, txtValorM3.Value, txtTotalM3.Value, txtValorTotalBloco.Value
    
    ' Chama Serviço
    MsgBox bloco.idSistema & " - " & bloco.nomeMaterial
End Sub
' Botão btnLTxtVoltarCadastroBloco tela cadastrar bloco
Private Sub btnLTxtVoltarCadastroBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage - tela estoque ³
    Me.MultiPageCEBC.Value = 1
End Sub
' Botão btnLTextLimparCadastroBloco tela cadastrar bloco
Private Sub btnLTxtLimparCadastroBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    Call limparCamposCadastroBlocos
End Sub

'-----------------------------------------------------------------TELA EDITAR BLOCO-----------------------------------
'                                                                 -----------------
' Carrega os campos com os dados do bloco tela editar bloco
Private Sub carregarDadosBlocoTelaEdicaoBloco(listaObjeto As Collection)
    
    ' Exibir o resultado da pesquisa
    For Each bloco In listaObjeto
     
        ' Descrição e dimensões finais
        txtIdBlocoEditar.Value = bloco.idSistema
        txtMaterialEditar.Value = bloco.nomeMaterial
        Call selecaoItem("cbTipoMaterialEditar", bloco.tipoMaterial.nome)
        txtObsEditar.Value = bloco.observacao
        Call selecaoItem("cbPedreiraEditar", bloco.pedreira.nome)
        Call selecaoItem("cbSerrariaEditar", bloco.serraria.nome)
        Call selecaoItem("cbPolideiraEditar", bloco.polideira.nome)
        txtNBlocoPedreiraEditar.Value = bloco.numeroBlocoPedreira
        Call selecaoItem("cbEstoqueEditar", bloco.estoque)
        txtDataCadastroEditar.Value = bloco.dataCadastro
        txtQtdM3blocoEditar.Value = bloco.qtdM3
        txtQtdM2SerradaEditar.Value = bloco.qtdM2Serrada
        txtQtdM2PolimentoEditar.Value = bloco.qtdM2Polimento
        txtTotalChapaBlocoEditar.Value = bloco.qtdChapas
        cbStatusBlocoEditar.Value = bloco.status.nome
        Call selecaoItem("cbNotaBlocoEditar", bloco.nota)
        Call selecaoItem("cbCustoMedioEditar", bloco.consultarCustoMedio)
        
        ' Dimensões bloco e médias chapas
        txtCompBrutaBlocoEditar.Value = bloco.compBrutoBloco
        txtAltBrutaBlocoEditar.Value = bloco.altBrutoBloco
        txtLArgBrutaBlocoEditar.Value = bloco.largBrutoBloco
        txtCompLiquidoBlocoEditar.Value = bloco.compLiquidoBloco
        txtAltLiquidoBlocoEditar.Value = bloco.altLiquidoBloco
        txtLArgLiquidoBlocoEditar.Value = bloco.largLiquidoBloco
        txtCompBrutaBrutoChapaEditar.Value = bloco.compBrutoChapaBruta
        txtAltBrutaBrutoChapaEditar.Value = bloco.altBrutoChapaBruta
        txtCompBrutaliquidoChapaEditar.Value = bloco.compLiquidoChapaBruta
        txtAltBrutaLiquidoChapaEditar.Value = bloco.altBrutoChapaBruta
        txtCompPolidaBrutoChapaEditar.Value = bloco.compBrutoChapaPolida
        txtAltPolidaBrutoChapaEditar.Value = bloco.altBrutoChapaPolida
        txtCompPolidaLiquidoChapaEditar.Value = bloco.compLiquidoChapaPolida
        txtAltPolidaLiquidaChapaEditar.Value = bloco.altBrutoChapaPolida
        
        ' Valores
        txtValoBlocoEditar.Value = bloco.valorBloco
        txtPrecoBlocoEditar.Value = bloco.precoM3Bloco
        txtFreteBlocoEditar.Value = bloco.freteBloco
        txtValorSerradaEditar.Value = bloco.valorMetroSerrada
        txtValorPolimentoEditar.Value = bloco.valorMetroPolimento
        txtValorADDImpostosEditar.Value = bloco.valoresAdicionais
        txtTotalSerradaEditar.Value = bloco.valorTotalSerrada
        txtTotalPolimentoEditar.Value = bloco.valorTotalPolimento
        
        ' Custos
        txtCustoMaterialBlocoEditar.Value = bloco.custoMaterial
        txtTotalM2PolimentoBlocoEditar.Value = bloco.qtdM2Polimento
        txtTotalBlocoEditar.Value = bloco.valorTotalBloco
        
        ' Se Status do bloco for finalizado deixar visivel lBlocoFinalizado e cbAbrirBlocoEditar e desabilitar todos os campos
        If bloco.status.nome = "FECHADO" Then
            cbAbrirBlocoEditar.Visible = True
            lBlocoFinalizado.Visible = True
        End If
    Next bloco
End Sub
' Habilita e desabilita campos para edição tela editar bloco
Private Sub cbAbrirBlocoEditar_Click()
    If cbAbrirBlocoEditar.Value = True Then
        Call habilitaCamposBlocoEditar
    Else
        Call desabilitaCamposBlocoEditar
    End If
End Sub
' Botão btnLTxtSalvarEdicaoBloco tela editar bloco
Private Sub btnLTxtSalvarEdicaoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Serviço
    MsgBox "Chama Serviço editar bloco, tela editar bloco"
End Sub
' Botão btnLTxtVoltarEdicaoBloco tela editar bloco
Private Sub btnLTxtVoltarEdicaoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 1
End Sub

'-----------------------------------------------------------------TELA ESTOQUE M²-----------------------------------
'                                                                 ---------------
' Efeito de label nome do pdf tela estoque m²
Private Sub lDigiteNomeArquivoM2Explemplo_Click()
    lDigiteNomeArquivoM2.Visible = True
    lDigiteNomeArquivoM2Explemplo.Visible = False
    txtNomeArquivoEstoqueChapas.SetFocus
End Sub
' Efeito e coloca em caixa alta o texto em txtNomeArquivoEstoqueChapas tela estoque m²
Private Sub txtNomeArquivoEstoqueChapas_Change()
    lDigiteNomeArquivoM2.Visible = True
    lDigiteNomeArquivoM2Explemplo.Visible = False

    If txtNomeArquivoEstoqueChapas.Value = "" Then
        lDigiteNomeArquivoM2.Visible = False
        lDigiteNomeArquivoM2Explemplo.Visible = True
    End If

    txtNomeArquivoEstoqueChapas.Value = UCase(txtNomeArquivoEstoqueChapas.Value)
End Sub
' Efeito ao sair da caixa txtNomeArquivoEstoqueChapas de texto tela estoque m²
Private Sub txtNomeArquivoEstoqueChapas_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtNomeArquivoEstoqueChapas.Value = "" Then
        lDigiteNomeArquivoM2.Visible = False
        lDigiteNomeArquivoM2Explemplo.Visible = True
    End If
End Sub
' txtMaterialChapaPesquisa tela estoque m²
Private Sub txtMaterialChapaPesquisa_Change()
    ' Coloca tudo em caixa alta
    txtMaterialChapaPesquisa.Value = UCase(txtMaterialChapaPesquisa.Value)
End Sub
' txtIdBlocoChapaPesquisa tela estoque m²
Private Sub txtIdBlocoChapaPesquisa_Change()
    ' Coloca tudo em caixa alta
    txtIdBlocoChapaPesquisa.Value = UCase(txtIdBlocoChapaPesquisa.Value)
End Sub
' txtIdchapaEstoque tela estoque m²
Private Sub txtIdchapaEstoque_Change()
    ' Coloca tudo em caixa alta
    txtIdchapaEstoque.Value = UCase(txtIdchapaEstoque.Value)
End Sub
' Botão btnLTxtPesquisarChapas tela estoque m²
Private Sub btnLTxtPesquisarChapas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço pesquisar chapa, tela estoque m²"
End Sub
' Botão btnLTxtLimparFiltrosChapas tela estoque m²
Private Sub btnLTxtLimparFiltrosChapas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    Call limparCamposPesquisaEstoqueM2
End Sub
' Botão btnLImgExportarEstoqueM2 tela estoque m²
Private Sub btnLImgExportarEstoqueM2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço esportar estoque m², tela estoque m²"
End Sub
'Botão btnLTxtNovoAvulso tela estoque m²
Private Sub btnLTxtNovoAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Coloca data atual na txtDataCadastroChapaAvulsa na tela cadastro chapa avulso
    txtDataCadastroChapaAvulsa.Value = Date
    
    ' Chama metodo para carregar comboBox
    Call carregarTiposMateriais(Me.cbTipoMaterialL)
    Call carregarTiposPolimento(Me.cbTipoPolimentoL)
    Call carregarTemNota(Me.cbTemNotaAvulso)
    
    ' Chama metodo para carregar lista
    Call carregarList(Me.ListMaterias)
    
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 5
End Sub
' Botão btnLTxtEditarChapa tela estoque m²
Private Sub btnLTxtEditarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 6
    
    ' Chama serviço para pesquisa da chapa
    
    
    ' Carrega os ComboBox da tela
    Call carregarPolideiras(cbPolideiraChapa)
    Call carregarTiposPolimento(cbTipoPolimentoChapa)
    Call carregarTiposMateriais(cbTipoMaterialChapaC)
    Call carregarEstoque(cbEstoqueChapaC)
    Call carregarTiposMateriais(cbTiposMateriaisChapas)
    
    ' Carrega os dados na tela editar chapa
    Call carregarDadosChapaTelaEdicaoChapa ' Irá enviar o objeto chapa para poder carregar os campos
End Sub
' Botão btnLTxtTrocaEstoque tela estoque m²
Private Sub btnLTxtTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 7
End Sub

'-----------------------------------------------------------------TELA CADASTRO AVULSO-----------------------------------
'                                                                 --------------------
' txtIdBlocoAvulso tela cadastro avulso
Private Sub txtIdBlocoAvulso_Change()
    ' Coloca tudo em caixa alta
    txtIdBlocoAvulso.Value = UCase(txtIdBlocoAvulso.Value)
    
    ' Cria o código para o sistema
    txtIdBlocoAvulsoSistema.Value = txtIdBlocoAvulso & "-" & Util.ExtrairUltimaPalavra(txtMaterialAvulso.Value) & "-BL"
    
    ' Deixa em branco o codigo se as variaveis forem vazias
    If txtIdBlocoAvulso.Value = "" And txtMaterialAvulso.Value = "" Then
        txtIdBlocoAvulsoSistema.Value = ""
    End If
End Sub
' txtMaterialAvulso tela cadastro avulso
Private Sub txtMaterialAvulso_Change()
    ' Coloca tudo em caixa alta
    txtMaterialAvulso.Value = UCase(txtMaterialAvulso.Value)
    
    ' Cria o código para o sistema
    txtIdBlocoAvulsoSistema.Value = txtIdBlocoAvulso & "-" & Util.ExtrairUltimaPalavra(txtMaterialAvulso.Value) & "-BL"
    
    ' Deixa em branco o codigo se as variaveis forem vazias
    If txtIdBlocoAvulso.Value = "" And txtMaterialAvulso.Value = "" Then
        txtIdBlocoAvulsoSistema.Value = ""
    End If
End Sub
' txtObsBlocoL tela cadastro avulso
Private Sub txtObsBlocoL_Change()
    ' Coloca tudo em caixa alta
    txtObsBlocoL.Value = UCase(txtObsBlocoL.Value)
End Sub
' txtComprimentoChapaAvulsa tela cadastro avulso
Private Sub txtComprimentoChapaAvulsa_Change()
   'Define o resultado no TextBox
    txtComprimentoChapaAvulsa.Value = Util.formatarMetros(txtComprimentoChapaAvulsa)
    
    ' Seta o valor no comprimento bruto
    txtCompChapasBrutasAvulso.Value = txtComprimentoChapaAvulsa.Value

    'Move o cursor para o final do TextBox
    txtComprimentoChapaAvulsa.SelStart = Len(txtComprimentoChapaAvulsa.Value)
    
    'Retorna valor calculado e formatado
    txtTotalM2Avulso.Value = Util.formatarComPontos(Format(Util.calcularM2(txtComprimentoChapaAvulsa.Value, _
        txtAlturaChapaAvulsa.Value, txtQuantidadeChapasAvulsas.Value), "0.0000"))
        
    'Seta o custo do material m²
    txtCustoSimplesM2Avulso.Value = Util.formatarComPontos(Format(Util.custoMaterialM2(txtTotalBlocoAvulso.Value, _
            txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", txtTotalM2Avulso.Value), "0.00"))
End Sub
' txtAlturaChapaAvulsa tela cadastro avulso
Private Sub txtAlturaChapaAvulsa_Change()
   'Define o resultado no TextBox
    txtAlturaChapaAvulsa.Value = Util.formatarMetros(txtAlturaChapaAvulsa)
    
    ' Seta o valor na altura bruto
    txtAlturaChapasBrutasAvulso.Value = txtAlturaChapaAvulsa.Value

    'Move o cursor para o final do TextBox
    txtAlturaChapaAvulsa.SelStart = Len(txtAlturaChapaAvulsa.Value)
    
    'Retorna valor calculado e formatado
    txtTotalM2Avulso.Value = Util.formatarComPontos(Format(Util.calcularM2(txtComprimentoChapaAvulsa.Value, _
        txtAlturaChapaAvulsa.Value, txtQuantidadeChapasAvulsas.Value), "0.0000"))
        
    'Seta o custo do material m²
    txtCustoSimplesM2Avulso.Value = Util.formatarComPontos(Format(Util.custoMaterialM2(txtTotalBlocoAvulso.Value, _
            txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", txtTotalM2Avulso.Value), "0.00"))
End Sub
' txtQuantidadeChapasAvulsas tela cadastro avulso
Private Sub txtQuantidadeChapasAvulsas_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Deixa só a digitação de numero
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub
' txtQuantidadeChapasAvulsas tela cadastro avulso
Private Sub txtQuantidadeChapasAvulsas_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Se não tiver valor deixa com 0
    If txtQuantidadeChapasAvulsas.Value = "" Then
        txtQuantidadeChapasAvulsas.Value = 0
    End If
    
    'Retorna valor calculado e formatado
    txtTotalM2Avulso.Value = Util.formatarComPontos(Format(Util.calcularM2(txtComprimentoChapaAvulsa.Value, _
        txtAlturaChapaAvulsa.Value, txtQuantidadeChapasAvulsas.Value), "0.0000"))
    'Seta o custo do material m²
    txtCustoSimplesM2Avulso.Value = Util.formatarComPontos(Format(Util.custoMaterialM2(txtTotalBlocoAvulso.Value, _
            txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", txtTotalM2Avulso.Value), "0.00"))
End Sub
' txtCompChapasBrutasAvulso tela cadastro avulso
Private Sub txtCompChapasBrutasAvulso_Change()
   'Define o resultado no TextBox
    txtCompChapasBrutasAvulso.Value = Util.formatarMetros(txtCompChapasBrutasAvulso)

    'Move o cursor para o final do TextBox
    txtCompChapasBrutasAvulso.SelStart = Len(txtCompChapasBrutasAvulso.Value)
End Sub
' txtAlturaChapasBrutasAvulso tela cadastro avulso
Private Sub txtAlturaChapasBrutasAvulso_Change()
   'Define o resultado no TextBox
    txtAlturaChapasBrutasAvulso.Value = Util.formatarMetros(txtAlturaChapasBrutasAvulso)

    'Move o cursor para o final do TextBox
    txtAlturaChapasBrutasAvulso.SelStart = Len(txtAlturaChapasBrutasAvulso.Value)
End Sub
' txtAdicionaisAvulso tela cadastro avulso
Private Sub txtAdicionaisAvulso_Change()
    ' Define o resultado no TextBox
    txtAdicionaisAvulso.Value = Util.formatarValor(txtAdicionaisAvulso.Value)
    
    ' Move o cursor para o final do TextBox
    txtAdicionaisAvulso.SelStart = Len(txtAdicionaisAvulso.Value)
    
    'Seta o custo do material m²
    txtCustoSimplesM2Avulso.Value = Util.formatarComPontos(Format(Util.custoMaterialM2(txtTotalBlocoAvulso.Value, _
            txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", txtTotalM2Avulso.Value), "0.00"))
End Sub
' txtValorFreteAvulso tela cadastro avulso
Private Sub txtValorFreteAvulso_Change()
    ' Define o resultado no TextBox
    txtValorFreteAvulso.Value = Util.formatarValor(txtValorFreteAvulso.Value)

    ' Move o cursor para o final do TextBox
    txtValorFreteAvulso.SelStart = Len(txtValorFreteAvulso.Value)
    
    'Seta o custo do material m²
    txtCustoSimplesM2Avulso.Value = Util.formatarComPontos(Format(Util.custoMaterialM2(txtTotalBlocoAvulso.Value, _
            txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", txtTotalM2Avulso.Value), "0.00"))
End Sub
' txtValorBlocoAvulso tela cadastro avulso
Private Sub txtValorMetroAvulso_Change()
    ' Define o resultado no TextBox
    txtValorMetroAvulso.Value = Util.formatarValor(txtValorMetroAvulso.Value)
    
    ' Move o cursor para o final do TextBox
    txtValorMetroAvulso.SelStart = Len(txtValorMetroAvulso.Value)

    ' Retorna valor calculado e formatado
    txtTotalBlocoAvulso.Value = Util.formatarComPontos(Format(Util.calcularValorBloco(txtTotalM2Avulso.Value, _
            txtValorMetroAvulso.Value), "0.00"))
            
    'Se m² for diferente de 0 calcula o custo do material
    'Seta o custo do material m²
    txtCustoSimplesM2Avulso.Value = Util.formatarComPontos(Format(Util.custoMaterialM2(txtTotalBlocoAvulso.Value, _
            txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", txtTotalM2Avulso.Value), "0.00"))
End Sub
' txtTotalM2Avulso tela cadastro avulso
Private Sub txtTotalM2Avulso_Change()
    'Se m² for diferente de 0 calcula o custo do material
    'Seta o custo do material m²
    txtCustoSimplesM2Avulso.Value = Util.formatarComPontos(Format(Util.custoMaterialM2(txtTotalBlocoAvulso.Value, _
            txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", txtTotalM2Avulso.Value), "0.00"))
End Sub
' Botão btnLImgCadastrarMaterialAvulso tela cadastro avulso
Private Sub btnLImgCadastrarMaterialAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar tipo material, tela cadastro avulso"
End Sub
' Botão btnLImgCadastrarPolimentoAvulso tela cadastro avulso
Private Sub btnLImgCadastrarPolimentoAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar tipo polimento, tela cadastro avulso"
End Sub
' Botão btnLTxtCadastrarChapaAvulso tela cadastro avulso
Private Sub btnLTxtCadastrarChapaAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar chapas avulsos, tela cadastro avulso"
End Sub
' Botão btnLTxtVoltarCadatradoChapasAvulso tela cadastro avulso
Private Sub btnLTxtVoltarCadatradoChapasAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage - tela estoque m²
    Me.MultiPageCEBC.Value = 4
End Sub
' Botão btnLTxtLimparCadastroChapaAvulso tela cadastro avulso
Private Sub btnLTxtLimparCadastroChapaAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    Call limparCamposCadastroAvulso
End Sub

'-----------------------------------------------------------------TELA LANÇAMENTO E EDIÇÃO CHAPA-----------------------------------
'                                                                 ------------------------------
' Carrega os campos com os dados da chapa tela lançamento e edição chapa
Private Sub carregarDadosChapaTelaEdicaoChapa() ' Irá receber o objeto Chapa para poder carregar os campos e algum campos do bloco
    ' Dados bloco
    txtIdBlocoPedreiraChapa.Value = "20745-MOON-LIGHT-BL"
    txtDecricaoBlocoChapa.Value = "BLOCO BRANCO DALLAS MOON-LIGHT"
    txtQtdDisponivelChapaBloco.Value = "71"
    txtNBlocoPedreiraChapa.Value = "20745"
    txtTipoMaterialChapa.Value = "EXTRA"
    
    ' Dados chapa
    txtIdChapaSistema.Value = "20745-MOON-LIGHT-PO"
    txtDescricaoChapa.Value = "BRANCO DALLAS MOON-LIGHT POLIDO"
    txtEstoqueChapa.Value = "71"
    
    ' Dimensões e custos
    Call selecaoItem("cbPolideiraChapa", "SÃO ROQUE")
    Call selecaoItem("cbTipoPolimentoChapa", "POLIDO")
    Call selecaoItem("cbTipoMaterialChapaC", "COMERCIAL SATAND")
    Call selecaoItem("cbEstoqueChapaC", "CASA DO GRANITO")
    txtCompLiquidoChapa.Value = "3,0000"
    txtAlturaLiquidoChapa.Value = "2,0000"
    txtQtdChapaC.Value = "71"
    txtCompBrutoChapa.Value = "3,5000"
    txtAlturaBrutaChapa.Value = "2,5000"
    txtEspChapa.Value = "02"
    txtQtsM2Chapa.Value = "426,0000"
    txtCustoChapa.Value = "71,50"
    txtTotalChapas.Value = "30.459,00"
    
    ' Tamanhos diferentes
    ' Carrega lista com tamanhos das chapas
    Call carregarListTamanhosChapas(ListTamanhosChapas) ' Irá enviar id chapa para carregamento
   
    ' Se Status do bloco for finalizado deixar visivel lBlocoFinalizadoChapa e cbAbrirParaEdicao e desabilitar todos os campos
    
End Sub
' Botão btnLImgCadastrarPolideiraChapa tela lançamento e edição chapa
Private Sub btnLImgCadastrarPolideiraChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar Polideira, tela lançamento e edição chapa"
End Sub
' Botão btnLImgCadastrarTipoPolideiraChapa tela lançamento e edição chapa
Private Sub btnLImgCadastrarTipoPolideiraChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar tipo polimento, tela lançamento e edição chapa"
End Sub
' Botão btnLImgCadastrarTipoMaterialChapa tela lançamento e edição chapa
Private Sub btnLImgCadastrarTipoMaterialChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço tipo material, tela lançamento e edição chapa"
End Sub
' Botão btnLImgCadastrarTipoMaterialChapaTamanhos tela lançamento e edição chapa
Private Sub btnLImgCadastrarTipoMaterialChapaTamanhos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço tipo material, tela lançamento e edição chapa"
End Sub
' Botão btnLTxtAdicionarTamanhoChapa tela lançamento e edição chapa
Private Sub btnLTxtAdicionarTamanhoChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço adicionar tamanhos, tela lançamento e edição chapa"
End Sub
' Botão btnLTxtEditarTamanhoChapa tela lançamento e edição chapa
Private Sub btnLTxtEditarTamanhoChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço editar tamanho chapa, tela lançamento e edição chapa"
End Sub
' Botão btnLTxtTirarDaLista tela lançamento e edição chapa
Private Sub btnLTxtTirarDaLista_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço tira tamanho da lista, tela lançamento e edição chapa"
End Sub
' Botão btnLTxtSalvarChapa tela lançamento e edição chapa
Private Sub btnLTxtSalvarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço salva alteração da chapa, tela lançamento e edição chapa"
End Sub
' Botão btnLTxtVoltarChapa tela lançamento e edição chapa
Private Sub btnLTxtVoltarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço voltar, tela lançamento e edição chapa"
End Sub

'-----------------------------------------------------------------TELA TROCA ESTOQUE-----------------------------------
'                                                                 ------------------
' Botão btnLTxtAdicionarTrocaEstoque tela troca estoque
Private Sub btnLTxtAdicionarTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço adicionar material para troca, tela troca estoque"
End Sub
' Botão btnLTxtTrocarEstoque tela troca estoque
Private Sub btnLTxtTrocarEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço troca de estoque, tela troca estoque"
End Sub
' Botão btnLTxtVoltarTrocaEstoque tela troca estoque
Private Sub btnLTxtVoltarTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço voltar, tela troca estoque"
End Sub

'-----------------------------------------------------------------TELA DESPACHE-----------------------------------
'                                                                 -------------
' Botão btnLImgCadastrarMotoristaDespache tela despache
Private Sub btnLImgCadastrarMotoristaDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastro motorista, tela despache"
End Sub
' Botão btnLImgCadastrarDestinoDespache tela despache
Private Sub btnLImgCadastrarDestinoDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastro destino, tela despache"
End Sub
' Botão btnLTxtAdicionar tela despache
Private Sub btnLTxtAdicionar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço adicionar chapa, tela despache"
End Sub
' Botão btnLTxtDespachar tela despache
Private Sub btnLTxtDespachar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço despachar, tela despache"
End Sub
' Botão btnLTxtLimparDespache tela despache
Private Sub btnLTxtLimparDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço limpar dados, tela despache"
End Sub

'-----------------------------------------------------------------TELA CARREGOS-----------------------------------
'                                                                 -------------
' Botão btnLTxtPesquisarCarregos tela carregos
Private Sub btnLTxtPesquisarCarregos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço pesquisar por carregos, tela carregos"
End Sub
' Botão btnLTxtLimparListas tela carregos
Private Sub btnLTxtLimparListas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço limpar dados filtro, tela carregos"
End Sub
' Botão btnLImgExportarCarregoPDF tela carregos
Private Sub btnLImgExportarCarregoPDF_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço exportar carregos em pdf, tela carregos"
End Sub
' Botão btnLTxtEditarCarrego tela carregos
Private Sub btnLTxtEditarCarrego_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço editar carrego, tela carregos"
End Sub
' Botão btnLTxtVoltarCArrego tela carregos
Private Sub btnLTxtVoltarCArrego_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço voltar, tela carregos"
End Sub

'-----------------------------------------------------------------TELA CADASTROS DIVERSOS-----------------------------------
'                                                                 -----------------------
' Botão btnLTxtSalvarPedreira tela cadastros diversos
Private Sub btnLTxtSalvarPedreira_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar pedreira, tela cadastros diversos"
End Sub
' Botão btnLTxtSalvarSerraria tela cadastros diversos
Private Sub btnLTxtSalvarSerraria_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar serraria, tela cadastros diversos"
End Sub
' Botão btnLTxtSalvarPolideira tela cadastros diversos
Private Sub btnLTxtSalvarPolideira_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar polideira, tela cadastros diversos"
End Sub
' Botão btnLTxtSalvarTipoMaterial tela cadastros diversos
Private Sub btnLTxtSalvarTipoMaterial_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar tipo material, tela cadastros diversos"
End Sub
' Botão btnLTxtSalvarTipoPolimento tela cadastros diversos
Private Sub btnLTxtSalvarTipoPolimento_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar tipo polimento, tela cadastros diversos"
End Sub
' Botão btnLTxtSalvarMotorista tela cadastros diversos
Private Sub btnLTxtSalvarMotorista_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar motorista, tela cadastros diversos"
End Sub
' Botão btnLTxtSalvarMotorista tela cadastros diversos
Private Sub btnLTxtSalvarDestino_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar destino, tela cadastros diversos"
End Sub

'-----------------------------------------------------------------TELA USUARIO-----------------------------------
'                                                                 ------------
' Botão btnLTxtSalvarUsuario tela usuarios
Private Sub btnLTxtSalvarUsuario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço cadastrar ou editar usuário, tela usuarios"
End Sub
' Botão btnLTxtListUsuario tela usuarios
Private Sub btnLTxtListUsuario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço carrega lista com usuários, tela usuarios"
End Sub
' Botão btnLTxtListUsuarioLog tela usuarios
Private Sub btnLTxtListUsuarioLog_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço carrega lista com log dos usuários, tela usuarios"
End Sub

'-----------------------------------------------------------------DESABILITA  E HABILITA CAMPOS-----------------------------------
'                                                                 -----------------------------
' Desabilita campos da tela editar bloco
Private Sub desabilitaCamposBlocoEditar()
    txtIdBlocoEditar.Visible = False
    txtMaterialEditar.Visible = False
    cbTipoMaterialEditar.Visible = False
    txtObsEditar.Visible = False
    cbPedreiraEditar.Visible = False
    cbSerrariaEditar.Visible = False
    cbPolideiraEditar.Visible = False
    txtNBlocoPedreiraEditar.Visible = False
    cbEstoqueEditar.Visible = False
    txtDataCadastroEditar.Visible = False
    txtQtdM3blocoEditar.Visible = False
    txtQtdM2SerradaEditar.Visible = False
    txtQtdM2PolimentoEditar.Visible = False
    txtTotalChapaBlocoEditar.Visible = False
    cbStatusBlocoEditar.Visible = False
    cbNotaBlocoEditar.Visible = False
    cbCustoMedioEditar.Visible = False
    
    ' Dimensões bloco e médias chapas
    txtCompBrutaBlocoEditar.Visible = False
    txtAltBrutaBlocoEditar.Visible = False
    txtLArgBrutaBlocoEditar.Visible = False
    txtCompLiquidoBlocoEditar.Visible = False
    txtAltLiquidoBlocoEditar.Visible = False
    txtLArgLiquidoBlocoEditar.Visible = False
    txtCompBrutaBrutoChapaEditar.Visible = False
    txtAltBrutaBrutoChapaEditar.Visible = False
    txtCompBrutaliquidoChapaEditar.Visible = False
    txtAltBrutaLiquidoChapaEditar.Visible = False
    txtCompPolidaBrutoChapaEditar.Visible = False
    txtAltPolidaBrutoChapaEditar.Visible = False
    txtCompPolidaLiquidoChapaEditar.Visible = False
    txtAltPolidaLiquidaChapaEditar.Visible = False
    
    ' Valores
    txtValoBlocoEditar.Visible = False
    txtPrecoBlocoEditar.Visible = False
    txtFreteBlocoEditar.Visible = False
    txtValorSerradaEditar.Visible = False
    txtValorPolimentoEditar.Visible = False
    txtValorADDImpostosEditar.Visible = False
    txtTotalSerradaEditar.Visible = False
    txtTotalPolimentoEditar.Visible = False
    
    ' Custos
    txtCustoMaterialBlocoEditar.Visible = False
    txtTotalM2PolimentoBlocoEditar.Visible = False
    txtTotalBlocoEditar.Visible = False
End Sub
' Habilita campos da tela editar bloco
Private Sub habilitaCamposBlocoEditar()
    txtIdBlocoEditar.Visible = True
    txtMaterialEditar.Visible = True
    cbTipoMaterialEditar.Visible = True
    txtObsEditar.Visible = True
    cbPedreiraEditar.Visible = True
    cbSerrariaEditar.Visible = True
    cbPolideiraEditar.Visible = True
    txtNBlocoPedreiraEditar.Visible = True
    cbEstoqueEditar.Visible = True
    txtDataCadastroEditar.Visible = True
    txtQtdM3blocoEditar.Visible = True
    txtQtdM2SerradaEditar.Visible = True
    txtQtdM2PolimentoEditar.Visible = True
    txtTotalChapaBlocoEditar.Visible = True
    cbStatusBlocoEditar.Visible = True
    cbNotaBlocoEditar.Visible = True
    cbCustoMedioEditar.Visible = True
    
    ' Dimensões bloco e médias chapas
    txtCompBrutaBlocoEditar.Visible = True
    txtAltBrutaBlocoEditar.Visible = True
    txtLArgBrutaBlocoEditar.Visible = True
    txtCompLiquidoBlocoEditar.Visible = True
    txtAltLiquidoBlocoEditar.Visible = True
    txtLArgLiquidoBlocoEditar.Visible = True
    txtCompBrutaBrutoChapaEditar.Visible = True
    txtAltBrutaBrutoChapaEditar.Visible = True
    txtCompBrutaliquidoChapaEditar.Visible = True
    txtAltBrutaLiquidoChapaEditar.Visible = True
    txtCompPolidaBrutoChapaEditar.Visible = True
    txtAltPolidaBrutoChapaEditar.Visible = True
    txtCompPolidaLiquidoChapaEditar.Visible = True
    txtAltPolidaLiquidaChapaEditar.Visible = True
    
    ' Valores
    txtValoBlocoEditar.Visible = True
    txtPrecoBlocoEditar.Visible = True
    txtFreteBlocoEditar.Visible = True
    txtValorSerradaEditar.Visible = True
    txtValorPolimentoEditar.Visible = True
    txtValorADDImpostosEditar.Visible = True
    txtTotalSerradaEditar.Visible = True
    txtTotalPolimentoEditar.Visible = True
    
    ' Custos
    txtCustoMaterialBlocoEditar.Visible = False
    txtTotalM2PolimentoBlocoEditar.Visible = False
    txtTotalBlocoEditar.Visible = False
End Sub

'-----------------------------------------------------------------LIMPAR CAMPOS-----------------------------------
'                                                                 -------------
' Limpa os campos de pesquisa da tela estoque M³
Private Sub limparCamposPesquisaEstoqueM3()
    txtDataInicioBlocoPesquisa.Value = ""
    txtDataFinalBlocoPesquisa.Value = ""
    txtMaterialBlocoPesquisa.Value = ""
    txtIdBlocoPesquisa.Value = ""
    cbPedreiraBlocoPesquisa.Value = ""
    cbSerrariaBlocoPesquisa.Value = ""
    cbTemNota.Value = ""
    obPedreiraESerrada.Value = True
    obEmEstoque.Value = False
    obFechado.Value = False
    opPedreiraSerradaEmProcesso.Value = False
    opTodos.Value = False
    chbPedreida.Value = True
    chbSerraria.Value = True
    chbChapasBrutas.Value = False
    chbEmProcesso.Value = False
    chbEstoque.Value = False
    chbFechado.Value = False
End Sub
' Limpa os campos de pesquisa da tela estoque M²
Private Sub limparCamposPesquisaEstoqueM2()
    txtMaterialChapaPesquisa.Value = ""
    txtIdBlocoChapaPesquisa.Value = ""
    txtIdchapaEstoque.Value = ""
    cbPolideiraChapaPesquisa.Value = ""
    cbTipoPolimentoPesquisa.Value = ""
    obEstoqueZeroSim.Value = False
    obEstoqueZeroNao.Value = True
    txtNomeArquivoEstoqueChapas.Value = ""
End Sub
' Limpa os campos da tela cadastrao de blocos
Private Sub limparCamposCadastroBlocos()
    txtDataCadastro.Value = Date
    txtIdBlocoSistema.Value = ""
    cbPedreira.Value = ""
    cbSerrariaCB.Value = ""
    txtIdBloco.Value = ""
    txtNomeBloco.Value = ""
    cbTipoMaterial.Value = ""
    cbNotaC.Value = ""
    obPedreiraCB.Value = True
    obSerrariaCB.Value = False
    txtObsBlocoCB.Value = ""
    txtComprimentoBloco.Value = "0,0000"
    txtAlturaBloco.Value = "0,0000"
    txtLarguraBloco.Value = "0,0000"
    txtCompBrutoBloco.Value = "0,0000"
    txtAlturaBlocoBruto.Value = "0,0000"
    txtLarguraBlocoBruto.Value = "0,0000"
    txtAdicionais.Value = "0,00"
    txtValorFreteBloco.Value = "0,00"
    txtValorM3.Value = "0,00"
    lTotalDia.Caption = "0,00"
End Sub
' Limpa os campos de pesquisa da tela cadastro avulso
Private Sub limparCamposCadastroAvulso()
    txtDataCadastroChapaAvulsa.Value = Date
    txtIdBlocoAvulsoSistema.Value = ""
    txtIdBlocoAvulso.Value = ""
    txtMaterialAvulso.Value = ""
    cbTipoMaterialL.Value = ""
    cbTipoPolimentoL.Value = ""
    obAvulso.Value = True
    opImportado.Value = False
    cbTemNotaAvulso.Value = ""
    txtObsBlocoL.Value = ""
    txtComprimentoChapaAvulsa.Value = ""
    txtAlturaChapaAvulsa.Value = ""
    txtQuantidadeChapasAvulsas.Value = 0
    txtEspessuraAvulso.Value = "02"
    txtCompChapasBrutasAvulso.Value = "0,0000"
    txtAlturaChapasBrutasAvulso.Value = "0,0000"
    txtAdicionaisAvulso.Value = "0,00"
    txtValorFreteAvulso.Value = "0,00"
    txtValorMetroAvulso.Value = "0,00"
    txtTotalM2Avulso.Value = "0,00"
    txtCustoSimplesM2Avulso.Value = "0,00"
    txtTotalBlocoAvulso.Value = "0,00"
End Sub

'-----------------------------------------------------------------CARREAGMENTO DOS COMBOBBOX-----------------------------------
'                                                                 --------------------------
' Seleção do item da comboBox
Private Sub selecaoItem(nomeCmboBox As String, nomeSelecao As String)
    ' Variaveis do metodo
    Dim i As Integer
    
    ' Percorrer os itens da ComboBox
    With formControle.Controls(nomeCmboBox)
        For i = 0 To .ListCount - 1
            If .list(i) = nomeSelecao Then
                .ListIndex = i ' Seleciona o item desejado
                Exit For
            End If
        Next i
    End With
End Sub

' Carrega a combobox de pedreira
Private Sub carregarPedreiras(cbPedreiras As MSForms.comboBox)
    ' limpa a lista para carregamento
    cbPedreiras.Clear
    
    ' Carregamento para lista
    cbPedreiras.AddItem "PEDREIRA 01"
    cbPedreiras.AddItem "PEDREIRA 02"
    cbPedreiras.AddItem "MINERAÇÃO VISTA LINDA"
'    ' for para carregamento
'    For Each nomePedreira In pedreiras
'
'        If nomePedreira <> "IMPORTADO" And nomePedreira <> "AVULSO" Then
'            ComboBoxPedreira.AddItem nomePedreira ' Tela cadastro de blocos
'
'        End If
'
'        If nomePedreira <> "AVULSO" Then
'            ComboBoxPedreiraBlocoPesquisa.AddItem nomePedreira ' Tela pesquisa de blocos
'
'        End If
'    Next nomePedreira
End Sub
' Carrega a combobox de serraria
Private Sub carregarSerrarias(cbSerrarias As MSForms.comboBox)
    ' limpa a lista para carregamento
    cbSerrarias.Clear
    
    ' Carregamento para lista
    cbSerrarias.AddItem "SERRARIA 01"
    cbSerrarias.AddItem "SERRARIA 02"
    cbSerrarias.AddItem "ELSON BABISQUE"
    
'    ' for para carregamento
'    For Each nomePedreira In pedreiras
'
'        If nomePedreira <> "IMPORTADO" And nomePedreira <> "AVULSO" Then
'            ComboBoxPedreira.AddItem nomePedreira ' Tela cadastro de blocos
'
'        End If
'
'        If nomePedreira <> "AVULSO" Then
'            ComboBoxPedreiraBlocoPesquisa.AddItem nomePedreira ' Tela pesquisa de blocos
'
'        End If
'    Next nomePedreira
End Sub
' Carrega a combobox de tipo material
Private Sub carregarTiposMateriais(cbTiposMateriais As MSForms.comboBox)
    ' limpa a lista para carregamento
    cbTiposMateriais.Clear
    
    ' Carregamento para lista
    cbTiposMateriais.AddItem "TIPO 01"
    cbTiposMateriais.AddItem "TIPO 02"
    cbTiposMateriais.AddItem "COMERCIAL SATAND"
'    ' for para carregamento
'    For Each nomePedreira In pedreiras
'
'        If nomePedreira <> "IMPORTADO" And nomePedreira <> "AVULSO" Then
'            ComboBoxPedreira.AddItem nomePedreira ' Tela cadastro de blocos
'
'        End If
'
'        If nomePedreira <> "AVULSO" Then
'            ComboBoxPedreiraBlocoPesquisa.AddItem nomePedreira ' Tela pesquisa de blocos
'
'        End If
'    Next nomePedreira
End Sub
' Carrega a combobox tem nota
Private Sub carregarTemNota(cbTemNota As MSForms.comboBox)
    ' limpa a lista para carregamento
    cbTemNota.Clear
    
    ' Carregamento para lista
    cbTemNota.AddItem "SIM"
    cbTemNota.AddItem "NÃO"
End Sub
' Carrega a combobox de polideira
Private Sub carregarPolideiras(cbPolideiras As MSForms.comboBox)
    ' limpa a lista para carregamento
    cbPolideiras.Clear
    
    ' Carregamento para lista
    cbPolideiras.AddItem "POLIDEIRA 01"
    cbPolideiras.AddItem "POLIDEIRA 02"
    cbPolideiras.AddItem "SÃO ROQUE"
End Sub
' Carrega a combobox de tipo polimento
Private Sub carregarTiposPolimento(cbTiposPolimento As MSForms.comboBox)
    ' limpa a lista para carregamento
    cbTiposPolimento.Clear
    
    ' Carregamento para lista
    cbTiposPolimento.AddItem "TIPO 01"
    cbTiposPolimento.AddItem "TIPO 02"
    cbTiposPolimento.AddItem "POLIDO"
End Sub
' Carrega a combobox de estoque carregarCustoMedio
Private Sub carregarEstoque(cbTiposEstoque As MSForms.comboBox)
    ' limpa a lista para carregamento
    cbTiposEstoque.Clear
    
    ' Carregamento para lista
    cbTiposEstoque.AddItem "CASA DO GRANITO"
End Sub
' Carrega a combobox de custo medio
Private Sub carregarCustoMedio(cbCustoMedio As MSForms.comboBox)
    ' limpa a lista para carregamento
    cbCustoMedio.Clear
    
    ' Carregamento para lista
    cbCustoMedio.AddItem "SIM"
    cbCustoMedio.AddItem "NÃO"
End Sub
' Carrega a combobox de status
Private Sub carregarStatus(cbStatus As MSForms.comboBox)
    ' limpa a lista para carregamento
    cbStatus.Clear
    
    ' Carregamento para lista
    cbStatus.AddItem "STATUS 01"
    cbStatus.AddItem "STATUS 02"
    cbStatus.AddItem "EM PROCESSO"
End Sub

'-----------------------------------------------------------------CARREAGMENTO DAS LIST-----------------------------------
'                                                                 ---------------------
' Carrega a lista
Private Sub carregarList(lista As MSForms.ListBox)
'    'Variaveis do metodo
'    Dim Data As String
'    Dim listaBlocos As Variant
'    Dim totalDia As Double
'    Dim tamanhoLista As Integer
'    Dim i As Long
    
'    'Formata a data
'    Data = Util.ConverterFormatoData(txtDataCadastro.Value)
'
'    'Criando Coleção para manipulaçã
'    listaBlocos = BlocosDAO.listarBlocosEntreDatas(Data, Data)

    ' Limpar a ListBox
    lista.Clear
    
    ' NOME CABEÇALHO CHAPAS       | COD | DECRCIÇÃO | QTD  | COMP  | ALT  | M²    | TIPO     | ESP | VALOR | TOTAL |
    ' NOME CABEÇALHO BLOCOS       | COD | DECRCIÇÃO | COMP | ALT   | LARG | QTD   | VALOR M³ | ADD | FRETE | TOTAL |
    ' Tamanho do cabeçalho left   | 7   | 193       | 444  | 496,5 | 549  | 601,5 | 654      | 745 | 820,5 | 896   |
    ' Tamanho do cabeçalho width  | 185 | 250       | 52   | 52    | 52   | 52    | 90       | 75  | 75    | 74,5  |
    ' Tamanho das colunas da list
    lista.ColumnWidths = "185;250;52;52;52;52;90;75;75;74;"

'    'Captura o tamanho da matriz
'    tamanhoLista = TamanhoDaMatriz(listaBlocos)
    
'    'Sem não tiver dados
'    If listaBlocos(1, 1) = "SEM DADOS" Then
'
'        'Seta valor na label
'        lTotalDia = "0,00"
'        lTotalDia.Caption = Format(totalDia, "0.00")
'
'    Else
    
'        'Adiciona os dados na ListBox
'        For i = 1 To tamanhoLista
    
            'Adiciona uma linha
            lista.AddItem
            
            'Adiciona os dados do bloco
            lista.list(lista.ListCount - 1, 0) = "37766-50793-MOON-LIGHT-BL"
            lista.list(lista.ListCount - 1, 1) = "BLOCO BRANCO DALLAS MOON-LIGHT"
            lista.list(lista.ListCount - 1, 2) = "3,0000"
            lista.list(lista.ListCount - 1, 3) = "2,0000"
            lista.list(lista.ListCount - 1, 4) = "2,0000"
            lista.list(lista.ListCount - 1, 5) = "71"
            lista.list(lista.ListCount - 1, 6) = "1.500,00"
            lista.list(lista.ListCount - 1, 7) = "15.000,00"
            lista.list(lista.ListCount - 1, 8) = "5.000,00"
            lista.list(lista.ListCount - 1, 9) = "150.000,00"
                
'            'Soma o total do dia
'            lTotalDia = lTotalDia + Util.formatarComPontos(Format(Util.calcularValorBloco(CStr(listaBlocos(i, 11)), _
'                CStr(listaBlocos(i, 9))), "0.00"))
'        Next i
    
'        lTotalDia.Caption = listCadastradosHoje.List(listCadastradosHoje.ListCount - 1, 9)
'    End If

End Sub
' Carrega a lista ListTamanhosChapas tela edicao chapa
Private Sub carregarListTamanhosChapas(lista As MSForms.ListBox) ' Irá receber id chapa para carregamento
    ' Limpar a ListBox
    lista.Clear
    
    ' NOME CABEÇALHO BLOCOS       | TIPO  | ESP   | COMP  | ALT | M²  | QTD |
    ' Tamanho do cabeçalho left   | 192,5 | 331,5 | 362,5 | 411 | 460 | 511 |
    ' Tamanho do cabeçalho width  | 138,5 | 30    | 48    | 48  | 50  | 30  |
    ' Tamanho das colunas da list
    lista.ColumnWidths = "140,5;30;48;48;50;35;"

    'Adiciona uma linha
    lista.AddItem
    
    'Adiciona os dados do bloco
    lista.list(lista.ListCount - 1, 0) = "COMERCIAL SATAND"
    lista.list(lista.ListCount - 1, 1) = "02"
    lista.list(lista.ListCount - 1, 2) = "3,0000"
    lista.list(lista.ListCount - 1, 3) = "2,0000"
    lista.list(lista.ListCount - 1, 4) = "146,0000"
    lista.list(lista.ListCount - 1, 5) = "71"
End Sub
