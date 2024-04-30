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

' Variaveis para manipula��o com os bot�es e frames
Dim botoesMenu() As clsLabel
Dim botoesImg() As clsLabel
Dim botoesText() As clsLabel
Dim frameEfeito() As clsFrame
Dim errorStyle As clsErrorStyle

' Variaveis para manipula��o
Dim paginaAnterior As Integer
Dim status() As String
Dim listaObjeto As Collection

' Variaveis de objetos
Dim bloco As objBloco
Dim chapa As objChapa
Dim pedreira As objPedreira
Dim polideira As objPolideira
Dim serraria As objSerraria
Dim tipoMaterial As objTipoMaterial
Dim tipoPolimento As objTipoPolimento
Dim statusObj As objStatus
Dim estoque As objEstoque

' Inicializa��o do formControle
Private Sub UserForm_Initialize()
    ' Variaveis para o metodo
    Dim obj As Object
    Dim i As Long
    Dim j As Long
    Dim l As Long
    Dim m As Long
    Dim nameObj As String
    Dim nameObjInicio As String
    
    ' Carrega tradu��o do sistema
    Call M_TRADUCAO.carregarTraducaoErros
    
    ' Seta pagina
    paginaAnterior = 0
    
    ' Resevando espa�o em memoria para manipula��o das variaveis
    ReDim botoesMenu(1 To Me.Controls.Count)
    ReDim botoesImg(1 To Me.Controls.Count)
    ReDim botoesText(1 To Me.Controls.Count)
    ReDim frameEfeito(1 To Me.Controls.Count)
    ReDim status(1 To 6)
    Set errorStyle = New clsErrorStyle
    
    ' Atribui��es da variaveis
    status(1) = "PEDREIRA"
    status(2) = "SERRARIA"
    status(3) = "ESTOQUE"
    status(4) = "FECHADO"
    status(5) = "CHAPAS BRUTAS"
    status(6) = "EM PROCESSO"
    
    ' Separa os bot�es e frames
    For Each obj In Me.Controls
        
        ' Atribui��es das variaveis para manipula��es
        nameObj = obj.name
        nameObjInicio = Mid(nameObj, 1, 7)
        
        ' Captura os bot�es no menu
        If nameObjInicio = "btnLMen" Then
            i = i + 1
            Set botoesMenu(i) = New clsLabel
            Set botoesMenu(i).efeitoBotoesMenu = obj
        End If
        
        ' Captura os bot�es com imagens
        If nameObjInicio = "btnLImg" Then
            j = j + 1
            Set botoesImg(j) = New clsLabel
            Set botoesImg(j).efeitoBotoesImagem = obj
        End If
        
        ' Captura os bot�es com textos
        If nameObjInicio = "btnLTxt" Then
            l = l + 1
            Set botoesText(l) = New clsLabel
            Set botoesText(l).efeitoBotoesTexto = obj
        End If
        
        ' Captura os frames para efeitos com bot�es
        If nameObjInicio = "fTiraEf" Then
            m = m + 1
            Set frameEfeito(m) = New clsFrame
            Set frameEfeito(m).efeitoFrame = obj
        End If
    Next obj
    
    ' Limpando a variavel
    Set obj = Nothing
    
    ' Redefini��o dos espa�o em memoria das variaveis
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
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 1
    ' Seta pagina anterior
    paginaAnterior = 0
    ' Seta o foco
    txtMaterialBlocoPesquisa.SetFocus
    
    ' Carregar os comboBox da tela
    Call carregarPedreiras(Me.cbPedreiraBlocoPesquisa)
    Call carregarSerrarias(Me.cbSerrariaBlocoPesquisa)
    Call carregarTemNota(Me.cbTemNota)
End Sub
' Efeito para clique nas label btnLMenuChapa do menu
Private Sub btnLMenuChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Carregar os comboBox da tela
    Call carregarPolideiras(Me.cbPolideiraChapaPesquisa)
    Call carregarTiposPolimento(Me.cbTipoPolimentoPesquisa)
    
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 4
    ' Seta o foco
    txtMaterialChapaPesquisa.SetFocus
End Sub
' Efeito para clique nas label btnLMenuDespachar do menu
Private Sub btnLMenuDespachar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 8
    ' Seta o foco
    cbMotorista.SetFocus
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

'-----------------------------------------------------------------TELA ESTOQUE M�-----------------------------------
'                                                                 ---------------
' Efeito de label nome do pdf tela estoque m�
Private Sub lDigiteNomeArquivoM3Explemplo_Click()
    lDigiteNomeArquivoM3.Visible = True
    lDigiteNomeArquivoM3Explemplo.Visible = False
    txtNomeArquivoEstoqueBlocos.SetFocus
End Sub
' Efeito e coloca em caixa alta o texto em txtNomeArquivoEstoqueBlocos tela estoque m�
Private Sub txtNomeArquivoEstoqueBlocos_Change()
    lDigiteNomeArquivoM3.Visible = True
    lDigiteNomeArquivoM3Explemplo.Visible = False

    If txtNomeArquivoEstoqueBlocos.Value = "" Then
        lDigiteNomeArquivoM3.Visible = False
        lDigiteNomeArquivoM3Explemplo.Visible = True
    End If

    txtNomeArquivoEstoqueBlocos.Value = UCase(txtNomeArquivoEstoqueBlocos.Value)
End Sub
' Efeito ao sair da caixa txtNomeArquivoEstoqueBlocos de texto tela estoque m�
Private Sub txtNomeArquivoEstoqueBlocos_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtNomeArquivoEstoqueBlocos.Value = "" Then
        lDigiteNomeArquivoM3.Visible = False
        lDigiteNomeArquivoM3Explemplo.Visible = True
    End If
End Sub
' Efeito para quando sair do foco de txtNomeArquivoEstoqueBlocos de texto tela estoque m�
Private Sub fTiraEfeitoBotoesExportarBlocosM3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtNomeArquivoEstoqueBlocos.Value = "" Then
        lDigiteNomeArquivoM3.Visible = False
        lDigiteNomeArquivoM3Explemplo.Visible = True
    End If
End Sub
' txtDataInicioBlocoPesquisa tela estoque m�
Private Sub txtDataInicioBlocoPesquisa_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Deixa s� a digita��o de numero
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
    ' Coloca as barras para formata��o
    If Len(txtDataInicioBlocoPesquisa.Value) = 2 Or Len(txtDataInicioBlocoPesquisa.Value) = 5 Then
    
        txtDataInicioBlocoPesquisa.Value = txtDataInicioBlocoPesquisa.Value & "/"
    End If
End Sub
' txtDataFinalBlocoPesquisa tela estoque m�
Private Sub txtDataFinalBlocoPesquisa_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Deixa s� a digita��o de numero
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
    ' Coloca as barras para formata��o
    If Len(txtDataFinalBlocoPesquisa.Value) = 2 Or Len(txtDataFinalBlocoPesquisa.Value) = 5 Then
    
        txtDataFinalBlocoPesquisa.Value = txtDataFinalBlocoPesquisa.Value & "/"
    End If
End Sub
' txtMaterialBlocoPesquisa tela estoque m�
Private Sub txtMaterialBlocoPesquisa_Change()
    'Coloca tudo em caixa alta
    txtMaterialBlocoPesquisa.Value = UCase(txtMaterialBlocoPesquisa.Value)
End Sub
' txtIdBlocoPesquisa tela estoque m�
Private Sub txtIdBlocoPesquisa_Change()
    'Coloca tudo em caixa alta
    txtIdBlocoPesquisa.Value = UCase(txtIdBlocoPesquisa.Value)
End Sub
' Atelho para sele��o dos status, obPedreiraESerrada tela estoque m�
Private Sub obPedreiraESerrada_Click()
    chbPedreida.Value = True
    chbSerraria.Value = True
    chbChapasBrutas.Value = False
    chbEmProcesso.Value = False
    chbEstoque.Value = False
    chbFechado.Value = False
End Sub
' Atelho para sele��o dos status, obEmEstoque tela estoque m�
Private Sub obEmEstoque_Click()
    chbPedreida.Value = False
    chbSerraria.Value = False
    chbChapasBrutas.Value = True
    chbEmProcesso.Value = True
    chbEstoque.Value = True
    chbFechado.Value = False
End Sub
' Atelho para sele��o dos status, obFechado tela estoque m�
Private Sub obFechado_Click()
    chbPedreida.Value = False
    chbSerraria.Value = False
    chbChapasBrutas.Value = False
    chbEmProcesso.Value = False
    chbEstoque.Value = False
    chbFechado.Value = True
End Sub
' Atelho para sele��o dos status, opPedreiraSerradaEmProcesso tela estoque m�
Private Sub opPedreiraSerradaEmProcesso_Click()
    chbPedreida.Value = True
    chbSerraria.Value = True
    chbChapasBrutas.Value = True
    chbEmProcesso.Value = True
    chbEstoque.Value = True
    chbFechado.Value = False
End Sub
' Atelho para sele��o dos status, opTodos tela estoque m�
Private Sub opTodos_Click()
    chbPedreida.Value = True
    chbSerraria.Value = True
    chbChapasBrutas.Value = True
    chbEmProcesso.Value = True
    chbEstoque.Value = True
    chbFechado.Value = True
End Sub
' Bot�o btnLTxtPesquisarBlocos tela estoque m�
Private Sub btnLTxtPesquisarBlocos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Variaveis do metodo
    Dim listaBlocos As Collection
    Dim dataInicial As String
    Dim dataFinal As String
    Dim idBlocoPedreira As String
    Dim descricaoBloco As String
    Dim pedreiraBloco As String
    Dim serrariaBloco As String
    Dim temNota As String
    Dim statusPedreira As String
    Dim statusSerraria As String
    Dim statusChapasBrutas As String
    Dim statusEmProcesso As String
    Dim statusEstoque As String
    Dim statusFechado As String
    
    ' Formata a data inicial
    If txtDataInicioBlocoPesquisa.Value = "" Or Len(txtDataInicioBlocoPesquisa.Value) < 10 Then
        txtDataInicioBlocoPesquisa.Value = M_METODOS_GLOBAL.dataInicial
    End If

    ' Formata a data final
    If txtDataFinalBlocoPesquisa.Value = "" Or Len(txtDataFinalBlocoPesquisa.Value) < 10 Then
        txtDataFinalBlocoPesquisa.Value = M_METODOS_GLOBAL.dataFinal
    End If
    
    'Validando a data
    If IsDate(txtDataInicioBlocoPesquisa.Value) = False Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtDataInicioBlocoPesquisa, ADICIONE_DATA_MENSAGEM, ADICIONE_DATA_TITULO
        'Para o fluxo do sistema para a corre��o
        Exit Sub
    End If
    ' Deixa na cor patr�o
    errorStyle.sairErrorStyleTextBox txtDataInicioBlocoPesquisa

    'Validando a data
    If IsDate(txtDataFinalBlocoPesquisa.Value) = False Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtDataFinalBlocoPesquisa, ADICIONE_DATA_MENSAGEM, ADICIONE_DATA_TITULO
        ' Para o fluxo do sistema para a corre��o
        Exit Sub
    End If
    ' Deixa na cor patr�o
    errorStyle.sairErrorStyleTextBox txtDataFinalBlocoPesquisa
    
    'Atribui��o das variaveies
    dataInicial = txtDataInicioBlocoPesquisa.Value
    dataFinal = txtDataFinalBlocoPesquisa.Value
    idBlocoPedreira = txtIdBlocoPesquisa.Value
    descricaoBloco = txtMaterialBlocoPesquisa.Value
    pedreiraBloco = cbPedreiraBlocoPesquisa.Value
    serrariaBloco = cbSerrariaBlocoPesquisa.Value
    temNota = cbTemNota.Value
    
    'Status filter
    statusPedreira = ""
    statusSerraria = ""
    statusChapasBrutas = ""
    statusEmProcesso = ""
    statusEstoque = ""
    statusFechado = ""
    
    'Status para pesquisa e formata��o
    If chbPedreida.Value = True Then
        statusPedreira = chbPedreida.Caption
    End If
    
    If chbSerraria.Value = True Then
        statusSerraria = chbSerraria.Caption
    End If
    
    If chbChapasBrutas.Value = True Then
        statusChapasBrutas = chbChapasBrutas.Caption
    End If
    
    If chbEmProcesso.Value = True Then
        statusEmProcesso = chbEmProcesso.Caption
    End If
    
    If chbEstoque.Value = True Then
        statusEstoque = chbEstoque.Caption
    End If
    
    If chbFechado.Value = True Then
        statusFechado = chbFechado.Caption
    End If
            
    'Mensagem para o usuario escolher algum Status
    If chbPedreida.Value = False And chbSerraria.Value = False And chbChapasBrutas.Value = False _
            And chbEmProcesso.Value = False And chbEstoque.Value = False And chbFechado.Value = False Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleOptionButton obPedreiraESerrada, ADICIONE_STATUS_MENSAGEM, ADICIONE_STATUS_TITULO
        ' Para o fluxo do sistema para a corre��o
        Exit Sub
    End If
    ' Deixa na cor patr�o
    errorStyle.sairErrorStyleOptionButton obPedreiraESerrada
    
    ' Faz pesquisa com filtros no banco de dados e retoeno uma lista
    Set listaBlocos = daoBloco.listarBlocosFilter(dataInicial, dataFinal, idBlocoPedreira, _
            descricaoBloco, pedreiraBloco, serrariaBloco, temNota, statusPedreira, statusSerraria, _
            statusChapasBrutas, statusEmProcesso, statusEstoque, statusFechado)
            
    ' Carrega a lista
    Call carregarList(ListEstoqueM3, listaBlocos)
    
    ' Libera espe�o na memoria
    Set listaBlocos = Nothing
End Sub
' Bot�o btnLTxtLimparFiltrosBlocos tela estoque m�
Private Sub btnLTxtLimparFiltrosBlocos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    Call limparCamposPesquisaEstoqueM3
End Sub
' Bot�o btnLImgExportarEstoqueM3 tela estoque m�
Private Sub btnLImgExportarEstoqueM3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Variaveis do metodo
    Dim idsParaPesquisa As Collection
    Dim id As String
    Dim i As Integer
    
    ' Verifica se tem dados na lista
    If Me.ListEstoqueM3.ListCount > 0 Then
        ' Reatribui espa�o na memoria para variavel
        Set idsParaPesquisa = ObjectFactory.factoryLista(idsParaPesquisa)
    Else
        ' Mensagem de erro
        errorStyle.Informativo LIST_SEM_DADOS_MENSAGEM, LIST_SEM_DADOS_TITULO
        ' Para o fluxo do sistema para a corre��o
        Exit Sub
    End If
    
    ' Verifica se foi digitado nome para o arquivo
    If txtNomeArquivoEstoqueBlocos.Value = "" Or txtNomeArquivoEstoqueBlocos.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtNomeArquivoEstoqueBlocos, ARQUIVO_SEM_NOME_MENSAGEM, ARQUIVO_SEM_NOME_TITULO
        ' Para o fluxo do sistema para a corre��o
        Exit Sub
    End If
    ' Deixa na cor patr�o
     errorStyle.sairErrorStyleTextBox txtNomeArquivoEstoqueBlocos
    
    ' Captura ids da lista
    For i = 0 To Me.ListEstoqueM3.ListCount - 1
        idsParaPesquisa.Add Me.ListEstoqueM3.list(i, 0)
    Next i
    
    ' Pesquisa os ids
    Set listaObjeto = daoBloco.pesquisarPorIdsVariados(idsParaPesquisa)
    
    ' Exporta em pdf
    Call ExportarArquivos.exportarEstoqueBloco(listaObjeto, txtNomeArquivoEstoqueBlocos.Value)
    
    ' Libera espe�o na memoria
    Set idsParaPesquisa = Nothing
    Set listaObjeto = Nothing
End Sub
' Bot�o btnLTxtNovoBloco tela estoque m�
Private Sub btnLTxtNovoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 2
    ' Seta n�mero de pagina para poder voltar
    paginaAnterior = 1
    ' Seta o foco
    cbPedreira.SetFocus
    
    ' Coloca data atual na txtDataCadastro na tela cadastro de bloco
    txtDataCadastro.Value = Date
    
    ' Chama metodo para carregar comboBox
    Call carregarPedreiras(Me.cbPedreira)
    Call carregarSerrarias(Me.cbSerrariaCB)
    Call carregarTiposMateriais(Me.cbTipoMaterial)
    Call carregarTemNota(Me.cbNotaC)
    
    ' Pesquisa blocos cadastrado no dia atual
    Set listaObjeto = daoBloco.listarBlocosFilter(Date, Date, "", "", "", "", "", "", "", "", "", "", "")
    
    ' Chama metodo para carregar lista e blocos cadastros do dia atual
    Call carregarList(Me.listCadastradosHoje, listaObjeto)
    ' Libera espa�o em memoria
    Set listaObjeto = Nothing
End Sub
' Bot�o btnLTxtEditarBloco tela estoque m�
Private Sub btnLTxtEditarBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Verifica se tem algum item selecionado
    If Me.ListEstoqueM3.ListIndex = -1 Then
        ' Mensagem usu�rio
        errorStyle.Informativo SELECIONE_TEM_MENSAGEM, SELECIONE_TEM_TITULO
        Exit Sub
    End If
    
    ' Muda abra da multPage para tela editar bloco
    Me.MultiPageCEBC.Value = 3
    ' Seta o foco
    txtMaterialEditar.SetFocus
    
    ' Chama servi�o para pesquisa do bloco
    Set bloco = daoBloco.pesquisarPorId(Me.ListEstoqueM3.list(Me.ListEstoqueM3.ListIndex, 0)) ' Envia o id do bloco
    
    ' Carregar os comboBox da tela
    Call carregarTiposMateriais(Me.cbTipoMaterialEditar)
    Call carregarPedreiras(Me.cbPedreiraEditar)
    Call carregarSerrarias(Me.cbSerrariaEditar)
    Call carregarPolideiras(Me.cbPolideiraEditar)
    Call carregarEstoque(Me.cbEstoqueEditar)
    Call carregarTemNota(Me.cbNotaBlocoEditar)
    Call carregarStatus(Me.cbStatusBlocoEditar)
    Call carregarCustoMedio(Me.cbCustoMedioEditar)
    
    ' Carrega os dados na tela editar bloco
    Call carregarDadosBlocoTelaEdicaoBloco(bloco)
End Sub
' Bot�o btnLTxtADDEstoque tela estoque m�
Private Sub btnLTxtADDEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
        ' Verifica se tem algum item selecionado
    If Me.ListEstoqueM3.ListIndex = -1 Then
        ' Mensagem usu�rio
        errorStyle.Informativo SELECIONE_TEM_MENSAGEM, SELECIONE_TEM_TITULO
        Exit Sub
    End If
    
    ' Bot�o chapa
    formControle.Controls("btnLMenuChapa").BackColor = RGB(200, 230, 255)
    formControle.Controls("btnLMenuChapa").Font.Size = 32
    formControle.Controls("btnLMenuChapa").Font.Size = 20
    formControle.Controls("btnLMenuChapa").Left = 15
    formControle.Controls("btnLMenuChapa").Width = 172
    formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignCenter
                
    ' Bot�o Menu
    formControle.Controls("btnLMenuBloco").BackColor = RGB(0, 100, 200)
    formControle.Controls("btnLMenuBloco").Left = 2
    formControle.Controls("btnLMenuBloco").Width = 189
    formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignLeft
    
    ' Muda abra da multPage para tela editar bloco
    Me.MultiPageCEBC.Value = 3
    ' Seta n�mero de pagina para poder voltar
    paginaAnterior = 1
    ' Seta o foco
    cbPolideiraChapa.SetFocus
    
    ' Chama servi�o para pesquisa do bloco
    Set bloco = daoBloco.pesquisarPorId(Me.ListEstoqueM3.list(Me.ListEstoqueM3.ListIndex, 0)) ' Envia o id do bloco
    
    ' Carregar os comboBox da tela
    Call carregarTiposMateriais(Me.cbTipoMaterialEditar)
    Call carregarPedreiras(Me.cbPedreiraEditar)
    Call carregarSerrarias(Me.cbSerrariaEditar)
    Call carregarPolideiras(Me.cbPolideiraEditar)
    Call carregarEstoque(Me.cbEstoqueEditar)
    Call carregarTemNota(Me.cbNotaBlocoEditar)
    Call carregarCustoMedio(Me.cbCustoMedioEditar)
    
    ' Carrega os dados na tela editar bloco
    Call carregarDadosBlocoTelaEdicaoBloco(bloco)
    ' Libera espa�o em memoria
    Set bloco = Nothing
End Sub

'-----------------------------------------------------------------TELA CADASTRO DE BLOCOS-----------------------------------
'                                                                 -----------------------
' txtIdBloco tela cadastro de bloco
Private Sub txtIdBloco_Change()
    ' Coloca tudo em caixa alta
    txtIdBloco.Value = UCase(txtIdBloco.Value)
    
    ' Cria o c�digo para o sistema
    txtIdBlocoSistema.Value = txtIdBloco & "-" & M_METODOS_GLOBAL.ExtrairUltimaPalavra(txtNomeBloco.Value) & "-BL"
    
    ' Deixa em branco o codigo se as variaveis forem vazias
    If txtIdBloco.Value = "" And txtNomeBloco.Value = "" Then
        txtIdBlocoSistema.Value = ""
    End If
End Sub
' txtNomeBloco tela cadastro de bloco
Private Sub txtNomeBloco_Change()
    ' Coloca tudo em caixa alta
    txtNomeBloco.Value = UCase(txtNomeBloco.Value)
    
    ' Cria o c�digo para o sistema
    txtIdBlocoSistema.Value = txtIdBloco & "-" & M_METODOS_GLOBAL.ExtrairUltimaPalavra(txtNomeBloco.Value) & "-BL"
    
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
    txtComprimentoBloco.Value = M_METODOS_GLOBAL.formatarMetros(txtComprimentoBloco.Value)
    
    ' Seta o valor no comprimento bruto
    txtCompBrutoBloco.Value = txtComprimentoBloco.Value
    
    ' Move o cursor para o final do TextBox
    txtComprimentoBloco.SelStart = Len(txtComprimentoBloco.Value)
    
    ' Retorna valor calculado e formatado
    txtTotalM3.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM3(txtComprimentoBloco.Value, _
            txtAlturaBloco.Value, txtLarguraBloco.Value), "0.0000"))

    ' Retorna valor calculado e formatado
    txtValorTotalBloco.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValorBloco( _
            txtValorM3.Value, txtTotalM3.Value), "0.00"))
End Sub
' txtAlturaBloco tela cadastro de bloco
Private Sub txtAlturaBloco_Change()
    ' Define o resultado no TextBox
    txtAlturaBloco.Value = M_METODOS_GLOBAL.formatarMetros(txtAlturaBloco.Value)
    
    ' Seta o valor na altura bruto
    txtAlturaBlocoBruto.Value = txtAlturaBloco.Value
    
    ' Move o cursor para o final do TextBox
    txtAlturaBloco.SelStart = Len(txtAlturaBloco.Value)
    
    ' Retorna valor calculado e formatado
    txtTotalM3.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM3(txtComprimentoBloco.Value, _
            txtAlturaBloco.Value, txtLarguraBloco.Value), "0.0000"))

    ' Retorna valor calculado e formatado
    txtValorTotalBloco.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValorBloco(txtValorM3.Value, _
            txtTotalM3.Value), "0.00"))
End Sub
' txtLarguraBloco tela cadastro de bloco
Private Sub txtLarguraBloco_Change()
    ' Define o resultado no TextBox
    txtLarguraBloco.Value = M_METODOS_GLOBAL.formatarMetros(txtLarguraBloco.Value)
    
    ' Seta o valor na altura bruto
    txtLarguraBlocoBruto.Value = txtLarguraBloco.Value
    
    ' Move o cursor para o final do TextBox
    txtLarguraBloco.SelStart = Len(txtLarguraBloco.Value)
    
    ' Retorna valor calculado e formatado
    txtTotalM3.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM3(txtComprimentoBloco.Value, _
            txtAlturaBloco.Value, txtLarguraBloco.Value), "0.0000"))

    ' Retorna valor calculado e formatado
    txtValorTotalBloco.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValorBloco( _
            txtValorM3.Value, txtTotalM3.Value), "0.00"))
End Sub
' txtCompBrutoBloco tela cadastro de bloco
Private Sub txtCompBrutoBloco_Change()
    ' Define o resultado no TextBox
    txtCompBrutoBloco.Value = M_METODOS_GLOBAL.formatarMetros(txtCompBrutoBloco.Value)
    
    ' Move o cursor para o final do TextBox
    txtCompBrutoBloco.SelStart = Len(txtCompBrutoBloco.Value)
End Sub
' txtAlturaBlocoBruto tela cadastro de bloco
Private Sub txtAlturaBlocoBruto_Change()
    ' Define o resultado no TextBox
    txtAlturaBlocoBruto.Value = M_METODOS_GLOBAL.formatarMetros(txtAlturaBlocoBruto.Value)
    
    ' Move o cursor para o final do TextBox
    txtAlturaBlocoBruto.SelStart = Len(txtAlturaBlocoBruto.Value)
End Sub
' txtLarguraBlocoBruto tela cadastro de bloco
Private Sub txtLarguraBlocoBruto_Change()
    ' Define o resultado no TextBox
    txtLarguraBlocoBruto.Value = M_METODOS_GLOBAL.formatarMetros(txtLarguraBlocoBruto.Value)
    
    ' Move o cursor para o final do TextBox
    txtLarguraBlocoBruto.SelStart = Len(txtLarguraBlocoBruto.Value)
End Sub
' txtAdicionais tela cadastro de bloco
Private Sub txtAdicionais_Change()
    ' Define o resultado no TextBox
    txtAdicionais.Value = M_METODOS_GLOBAL.formatarValor(txtAdicionais.Value)
    
    ' Move o cursor para o final do TextBox
    txtAdicionais.SelStart = Len(txtAdicionais.Value)
End Sub
' txtValorFreteBloco tela cadastro de bloco
Private Sub txtValorFreteBloco_Change()
    ' Define o resultado no TextBox
    txtValorFreteBloco.Value = M_METODOS_GLOBAL.formatarValor(txtValorFreteBloco.Value)

    ' Move o cursor para o final do TextBox
    txtValorFreteBloco.SelStart = Len(txtValorFreteBloco.Value)
End Sub
' txtValorM3 tela cadastro de bloco
Private Sub txtValorM3_Change()
    ' Define o resultado no TextBox
    txtValorM3.Value = M_METODOS_GLOBAL.formatarValor(txtValorM3.Value)
    
    ' Move o cursor para o final do TextBox
    txtValorM3.SelStart = Len(txtValorM3.Value)

    ' Retorna valor calculado e formatado
    txtValorTotalBloco.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValorBloco( _
            txtValorM3.Value, txtTotalM3.Value), "0.00"))
End Sub
' Bot�o btnLImgCadastrarPedreira tela cadastrar bloco
Private Sub btnLImgCadastrarPedreira_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar pedreira, tela cadastrar bloco"
End Sub
' Bot�o btnLImgCadastrarSerrariaCB tela cadastrar bloco
Private Sub btnLImgCadastrarSerrariaCB_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar serraria, tela cadastrar bloco"
End Sub
'Bot�o btnLImgCadastroTipoMaterial tela cadastrar bloco
Private Sub btnLImgCadastroTipoMaterial_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar tipo material, tela cadastrar bloco"
End Sub
' Bot�o btnLTxtCadastrarBloco tela cadastrar bloco
Private Sub btnLTxtCadastrarBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Variaveis do medoto
    Dim blocoPesquisa As objBloco
    Dim resposta As VbMsgBoxResult ' Variavel para confirma��o na hora de cadastrar
    Dim nomeStatus As String
    Dim nomeMaterial As String
    Dim cadastro As Boolean
    
    ' Patr�o true
    cadastro = True
    
    ' Captura do status
    If obPedreiraCB.Value = True Then
        nomeStatus = status(1)
    Else
        nomeStatus = status(2)
    End If
    
    ' Valida��es
    nomeMaterial = "BLOCO " & txtNomeBloco.Value
    
    ' Verifica o Status
    If obSerrariaCB.Value = True Then
        If cbSerrariaCB.Value = "" Or cbSerrariaCB.Value = " " Then
            ' Deixa visivel o erro com mensagens
            errorStyle.EntrarErrorStyleComboBox cbSerrariaCB, STATUS_SERRARIA_MENSAGEM, STATUS_SERRARIA_TITULO
            ' Para o fluxo do sistema para a corre��o
            Exit Sub
        End If
    End If
    ' Deixa na cor patr�o
    errorStyle.sairErrorStyleComboBox cbSerrariaCB
    
    ' Verifica o Pedreira
    If cbPedreira.Value = "" Or cbPedreira.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleComboBox cbPedreira, NOME_PEDREIRA_MENSAGEM, NOME_PEDREIRA_TITULO
        ' Para o fluxo do sistema para a corre��o
        Exit Sub
    End If
    ' Deixa na cor patr�o
    errorStyle.sairErrorStyleComboBox cbPedreira
    
    ' Verifica o N�mero do bloco na pedreira
    If txtIdBloco.Value = "" Or txtIdBloco.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtIdBloco, NUMERO_BLOCO_PEDREIRA_MENSAGEM, NUMERO_BLOCO_PEDREIRA_TITULO
        ' Para o fluxo do sistema para a corre��o
        Exit Sub
    End If
    ' Deixa na cor patr�o
    errorStyle.sairErrorStyleTextBox txtIdBloco
    
    ' Verifica nome do bloco
    If txtNomeBloco.Value = "" Or txtNomeBloco.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtNomeBloco, NOME_BLOCO_MENSAGEM, NOME_BLOCO_PEDREIRA_TITULO
        ' Para o fluxo do sistema para a corre��o
        Exit Sub
    End If
    ' Deixa na cor patr�o
    errorStyle.sairErrorStyleTextBox txtNomeBloco
        
    ' Verifica se um cadastrado ou edi��o
    Set blocoPesquisa = daoBloco.pesquisarPorId(txtIdBlocoSistema.Value)
    If blocoPesquisa.idSistema = txtIdBlocoSistema.Value Then
        ' Troca patr�o para false
        cadastro = False
        resposta = vbYes
    Else
        ' Mensagem de confirma��o
        resposta = MsgBox(CONFIRMACAO_CADASTRO_MENSAGEM, vbQuestion + vbYesNo, CONFIRMACAO_CADASTRO_TITULO)
    End If
    
    ' Verifica a confirma��o do us�rio para poder cadastrar
    If resposta = vbYes Then
        ' Cria��o dos objetos
        Set pedreira = daoPedreira.pesquisarPorNome(cbPedreira.Value)
        Set serraria = daoSerrada.pesquisarPorNome(cbSerrariaCB.Value)
        Set tipoMaterial = daoTipoMaterial.pesquisarPorNome(cbTipoMaterial.Value)
        Set statusObj = daoStatus.pesquisarPorNome(nomeStatus)
        Set estoque = daoEstoqueM3.pesquisarPorNome("CASA DO GRANITO")
        Set bloco = ObjectFactory.factoryBloco(bloco)
        Set blocoPesquisa = ObjectFactory.factoryBloco(blocoPesquisa)
        ' Cria��o do objeto
        bloco.carregarBlocoCadastro txtDataCadastro.Value, txtIdBlocoSistema.Value, pedreira, serraria, txtIdBloco.Value, _
                                    nomeMaterial, tipoMaterial, cbNotaC.Value, statusObj, txtObsBlocoCB.Value, _
                                    txtCompBrutoBloco.Value, txtAlturaBlocoBruto.Value, txtLarguraBlocoBruto.Value, _
                                    txtComprimentoBloco.Value, txtAlturaBloco.Value, txtLarguraBloco.Value, estoque, _
                                    txtAdicionais.Value, txtValorFreteBloco.Value, txtValorM3.Value, txtTotalM3.Value, _
                                    txtValorTotalBloco.Value, "N�O"
        
        ' Chama servi�o para cadastrar do bloco
        Call daoBloco.cadastrarEEditar(bloco)
        
        ' verifica se foi um cadastro ou edi��o para personalisar as mensagens
        If cadastro = True Then
            ' Verifica se bloco foi cadastrado
            Set blocoPesquisa = daoBloco.pesquisarPorId(bloco.idSistema)
            If blocoPesquisa.idSistema = txtIdBlocoSistema.Value Then
                'Mensagem de cadastro realizado com sucesso. Mensagem de erro utilizada para sucesso na opera��o
                errorStyle.Informativo CADASTRO_CONFIRMADO_MENSAGEM, CADASTRO_CONFIRMADO_TITULO
                ' Limpa os campos
                Call limparCamposCadastroBlocos
                ' Recarregar a lista com blocos cadastrados hoje
                ' Pesquisa blocos cadastrado no dia atual
                Set listaObjeto = daoBloco.listarBlocosFilter(Date, Date, "", "", "", "", "", "", "", "", "", "", "")
                
                ' Chama metodo para carregar lista e blocos cadastros do dia atual
                Call carregarList(Me.listCadastradosHoje, listaObjeto)
            Else
                'Mensagem de cadastro realizado com sucesso. Mensagem de erro utilizada para sucesso na opera��o
                errorStyle.Informativo ERRO_DESCONHECIDO_MENSAGEM, ERRO_DESCONHECIDO_TITULO
            End If
        Else
            ' Mensagem de sucesso na edi��o
            errorStyle.Informativo SUCESSO_EDICAO_MENSAGEM, SUCESSO_EDICAO_TITULO
        End If

        ' Libera espa�o da memoria
        Set pedreira = Nothing
        Set serraria = Nothing
        Set tipoMaterial = Nothing
        Set statusObj = Nothing
        Set estoque = Nothing
        Set bloco = Nothing
        Set blocoPesquisa = Nothing
    Else
        ' Coloque o c�digo a ser executado se o usu�rio clicar em "N�o" aqui.
        errorStyle.Informativo ACAO_CANCELADA_MENSAGEM, ACAO_CANCELADA_TITULO
        ' Para o fluxo do sistema para a corre��o
        Exit Sub
    End If
    ' Deixa o cursor no cbPedreira para proximo cadastro
    cbPedreira.SetFocus
End Sub
' Bot�o btnLTxtVoltarCadastroBloco tela cadastrar bloco
Private Sub btnLTxtVoltarCadastroBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage - tela estoque m�
    Me.MultiPageCEBC.Value = paginaAnterior
    ' Seta o foco
    txtMaterialBlocoPesquisa.SetFocus
End Sub
' Bot�o btnLTextLimparCadastroBloco tela cadastrar bloco
Private Sub btnLTxtLimparCadastroBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    Call limparCamposCadastroBlocos
End Sub

'-----------------------------------------------------------------TELA EDITAR BLOCO-----------------------------------
'                                                                 -----------------
' txtNBlocoPedreiraEditar tela editar bloco
Private Sub txtNBlocoPedreiraEditar_Change()
    ' Coloca tudo em caixa alta
    txtNBlocoPedreiraEditar.Value = UCase(txtNBlocoPedreiraEditar.Value)
End Sub

' txtMaterialEditar tela editar bloco
Private Sub txtMaterialEditar_Change()
    ' Coloca tudo em caixa alta
    txtMaterialEditar.Value = UCase(txtMaterialEditar.Value)
End Sub

' txtObsEditar tela editar bloco
Private Sub txtObsEditar_Change()
' Coloca tudo em caixa alta
    txtObsEditar.Value = UCase(txtObsEditar.Value)
End Sub

' txtDataCadastroEditar tela editar bloco
Private Sub txtDataCadastroEditar_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Deixa s� a digita��o de numero
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
    ' Coloca as barras para formata��o
    If Len(txtDataCadastroEditar.Value) = 2 Or Len(txtDataCadastroEditar.Value) = 5 Then
    
        txtDataCadastroEditar.Value = txtDataCadastroEditar.Value & "/"
    End If
End Sub

' txtQtdM3blocoEditar tela editar bloco
Private Sub txtQtdM3blocoEditar_Change()
    ' Define o resultado no TextBox
    txtQtdM3blocoEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtQtdM3blocoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtQtdM3blocoEditar.SelStart = Len(txtQtdM3blocoEditar.Value)
End Sub

' txtQtdM2SerradaEditar tela editar bloco
Private Sub txtQtdM2SerradaEditar_Change()
    ' Define o resultado no TextBox
    txtQtdM2SerradaEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtQtdM2SerradaEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtQtdM2SerradaEditar.SelStart = Len(txtQtdM2SerradaEditar.Value)
End Sub

' txtQtdM2PolimentoEditar tela editar bloco
Private Sub txtQtdM2PolimentoEditar_Change()
    ' Define o resultado no TextBox
    txtQtdM2PolimentoEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtQtdM2PolimentoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtQtdM2PolimentoEditar.SelStart = Len(txtQtdM2PolimentoEditar.Value)
    
    ' Seta m� polimento para calculo do custo
    txtTotalM2PolimentoBlocoEditar.Value = txtQtdM2PolimentoEditar.Value
End Sub

' txtTotalChapaBlocoEditar tela editar bloco
Private Sub txtTotalChapaBlocoEditar_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Deixa s� a digita��o de numero
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

' txtTotalChapaBlocoEditar tela editar bloco
Private Sub txtTotalChapaBlocoEditar_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtTotalChapaBlocoEditar.Value = "" Or txtTotalChapaBlocoEditar.Value = " " Then
        txtTotalChapaBlocoEditar.Value = "0"
    End If
End Sub
' txtCompBrutaBlocoEditar tela editar bloco
Private Sub txtCompBrutaBlocoEditar_Change()
    ' Define o resultado no TextBox
    txtCompBrutaBlocoEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtCompBrutaBlocoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtCompBrutaBlocoEditar.SelStart = Len(txtCompBrutaBlocoEditar.Value)
End Sub

' txtAltBrutaBlocoEditar tela editar bloco
Private Sub txtAltBrutaBlocoEditar_Change()
    ' Define o resultado no TextBox
    txtAltBrutaBlocoEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtAltBrutaBlocoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtAltBrutaBlocoEditar.SelStart = Len(txtAltBrutaBlocoEditar.Value)
End Sub

' txtLArgBrutaBlocoEditar tela editar bloco
Private Sub txtLArgBrutaBlocoEditar_Change()
    ' Define o resultado no TextBox
    txtLArgBrutaBlocoEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtLArgBrutaBlocoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtLArgBrutaBlocoEditar.SelStart = Len(txtLArgBrutaBlocoEditar.Value)
End Sub

' txtCompLiquidoBlocoEditar tela editar bloco
Private Sub txtCompLiquidoBlocoEditar_Change()
    ' Define o resultado no TextBox
    txtCompLiquidoBlocoEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtCompLiquidoBlocoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtCompLiquidoBlocoEditar.SelStart = Len(txtCompLiquidoBlocoEditar.Value)
    
    ' Retorna valor calculado e formatado
    txtQtdM3blocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM3( _
            txtCompLiquidoBlocoEditar.Value, txtAltLiquidoBlocoEditar.Value, txtLArgLiquidoBlocoEditar.Value), "0.0000"))

    ' Retorna valor calculado e formatado
    txtValoBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValorServicos( _
            txtPrecoBlocoEditar.Value, txtQtdM3blocoEditar.Value), "0.00"))
End Sub

' txtAltLiquidoBlocoEditar tela editar bloco
Private Sub txtAltLiquidoBlocoEditar_Change()
    ' Define o resultado no TextBox
    txtAltLiquidoBlocoEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtAltLiquidoBlocoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtAltLiquidoBlocoEditar.SelStart = Len(txtAltLiquidoBlocoEditar.Value)
    
    ' Retorna valor calculado e formatado
    txtQtdM3blocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM3( _
            txtCompLiquidoBlocoEditar.Value, txtAltLiquidoBlocoEditar.Value, txtLArgLiquidoBlocoEditar.Value), "0.0000"))

    ' Retorna valor calculado e formatado
    txtValoBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValorServicos( _
            txtPrecoBlocoEditar.Value, txtQtdM3blocoEditar.Value), "0.00"))
End Sub

' txtLArgLiquidoBlocoEditar tela editar bloco
Private Sub txtLArgLiquidoBlocoEditar_Change()
    ' Define o resultado no TextBox
    txtLArgLiquidoBlocoEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtLArgLiquidoBlocoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtLArgLiquidoBlocoEditar.SelStart = Len(txtLArgLiquidoBlocoEditar.Value)
    
    ' Retorna valor calculado e formatado
    txtQtdM3blocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM3( _
            txtCompLiquidoBlocoEditar.Value, txtAltLiquidoBlocoEditar.Value, txtLArgLiquidoBlocoEditar.Value), "0.0000"))

    ' Retorna valor calculado e formatado
    txtValoBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValorServicos( _
            txtPrecoBlocoEditar.Value, txtQtdM3blocoEditar.Value), "0.00"))
End Sub

' txtCompBrutaBrutoChapaEditar tela editar bloco
Private Sub txtCompBrutaBrutoChapaEditar_Change()
    ' Define o resultado no TextBox
    txtCompBrutaBrutoChapaEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtCompBrutaBrutoChapaEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtCompBrutaBrutoChapaEditar.SelStart = Len(txtCompBrutaBrutoChapaEditar.Value)
End Sub

' txtAltBrutaBrutoChapaEditar tela editar bloco
Private Sub txtAltBrutaBrutoChapaEditar_Change()
    ' Define o resultado no TextBox
    txtAltBrutaBrutoChapaEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtAltBrutaBrutoChapaEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtAltBrutaBrutoChapaEditar.SelStart = Len(txtAltBrutaBrutoChapaEditar.Value)
End Sub

' txtCompBrutaliquidoChapaEditar tela editar bloco
Private Sub txtCompBrutaliquidoChapaEditar_Change()
    ' Define o resultado no TextBox
    txtCompBrutaliquidoChapaEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtCompBrutaliquidoChapaEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtCompBrutaliquidoChapaEditar.SelStart = Len(txtCompBrutaliquidoChapaEditar.Value)
    
    ' Retorna valor calculado e formatado
    txtQtdM2SerradaEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM2( _
                                txtCompBrutaliquidoChapaEditar.Value, txtAltBrutaLiquidoChapaEditar.Value, _
                                txtTotalChapaBlocoEditar.Value), "0.0000"))
End Sub

' txtAltBrutaLiquidoChapaEditar tela editar bloco
Private Sub txtAltBrutaLiquidoChapaEditar_Change()
    ' Define o resultado no TextBox
    txtAltBrutaLiquidoChapaEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtAltBrutaLiquidoChapaEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtAltBrutaLiquidoChapaEditar.SelStart = Len(txtAltBrutaLiquidoChapaEditar.Value)
    
    ' Retorna valor calculado e formatado
    txtQtdM2SerradaEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM2( _
                                txtCompBrutaliquidoChapaEditar.Value, txtAltBrutaLiquidoChapaEditar.Value, _
                                txtTotalChapaBlocoEditar.Value), "0.0000"))
End Sub

' txtCompPolidaBrutoChapaEditar tela editar bloco
Private Sub txtCompPolidaBrutoChapaEditar_Change()
    ' Define o resultado no TextBox
    txtCompPolidaBrutoChapaEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtCompPolidaBrutoChapaEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtCompPolidaBrutoChapaEditar.SelStart = Len(txtCompPolidaBrutoChapaEditar.Value)
End Sub

' txtAltPolidaBrutoChapaEditar tela editar bloco
Private Sub txtAltPolidaBrutoChapaEditar_Change()
    ' Define o resultado no TextBox
    txtAltPolidaBrutoChapaEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtAltPolidaBrutoChapaEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtAltPolidaBrutoChapaEditar.SelStart = Len(txtAltPolidaBrutoChapaEditar.Value)
End Sub

' txtCompPolidaLiquidoChapaEditar tela editar bloco
Private Sub txtCompPolidaLiquidoChapaEditar_Change()
    ' Define o resultado no TextBox
    txtCompPolidaLiquidoChapaEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtCompPolidaLiquidoChapaEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtCompPolidaLiquidoChapaEditar.SelStart = Len(txtCompPolidaLiquidoChapaEditar.Value)
    
    ' Retorna valor calculado e formatado
    txtQtdM2PolimentoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM2( _
                                txtCompPolidaLiquidoChapaEditar.Value, txtAltPolidaLiquidaChapaEditar.Value, _
                                txtTotalChapaBlocoEditar.Value), "0.0000"))
End Sub

' txtAltPolidaLiquidaChapaEditar tela editar bloco
Private Sub txtAltPolidaLiquidaChapaEditar_Change()
    ' Define o resultado no TextBox
    txtAltPolidaLiquidaChapaEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtAltPolidaLiquidaChapaEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtAltPolidaLiquidaChapaEditar.SelStart = Len(txtAltPolidaLiquidaChapaEditar.Value)
    
    ' Retorna valor calculado e formatado
    txtQtdM2PolimentoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM2( _
                                txtCompPolidaLiquidoChapaEditar.Value, txtAltPolidaLiquidaChapaEditar.Value, _
                                txtTotalChapaBlocoEditar.Value), "0.0000"))
End Sub

' txtPrecoBlocoEditar tela editar bloco
Private Sub txtPrecoBlocoEditar_Change()
    ' Define o resultado no TextBox
    txtPrecoBlocoEditar.Value = M_METODOS_GLOBAL.formatarValor(txtPrecoBlocoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtPrecoBlocoEditar.SelStart = Len(txtPrecoBlocoEditar.Value)
                                
    ' Retorna valor calculado e formatado
    txtValoBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValorServicos( _
            txtPrecoBlocoEditar.Value, txtQtdM3blocoEditar.Value), "0.00"))
End Sub

' txtValoBlocoEditar tela editar bloco
Private Sub txtValoBlocoEditar_Change()
    ' Define o resultado no TextBox
    txtValoBlocoEditar.Value = M_METODOS_GLOBAL.formatarValor(txtValoBlocoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtValoBlocoEditar.SelStart = Len(txtValoBlocoEditar.Value)
                                
    ' Valor total do bloco
    txtTotalBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.custoBloco( _
                    txtValoBlocoEditar.Value, txtFreteBlocoEditar.Value, txtTotalSerradaEditar.Value, _
                    txtTotalPolimentoEditar.Value, txtValorADDImpostosEditar.Value), "0.00"))
End Sub

' txtFreteBlocoEditar tela editar bloco
Private Sub txtFreteBlocoEditar_Change()
    ' Define o resultado no TextBox
    txtFreteBlocoEditar.Value = M_METODOS_GLOBAL.formatarValor(txtFreteBlocoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtFreteBlocoEditar.SelStart = Len(txtFreteBlocoEditar.Value)
                                
    ' Valor total do bloco
    txtTotalBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.custoBloco( _
                    txtValoBlocoEditar.Value, txtFreteBlocoEditar.Value, txtTotalSerradaEditar.Value, _
                    txtTotalPolimentoEditar.Value, txtValorADDImpostosEditar.Value), "0.00"))
End Sub

' txtValorSerradaEditar tela editar bloco
Private Sub txtValorSerradaEditar_Change()
    ' Define o resultado no TextBox
    txtValorSerradaEditar.Value = M_METODOS_GLOBAL.formatarValor(txtValorSerradaEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtValorSerradaEditar.SelStart = Len(txtValorSerradaEditar.Value)
                                
    ' Valor da serrada
    txtTotalSerradaEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValorServicos( _
                    txtQtdM2SerradaEditar.Value, txtValorSerradaEditar.Value), "0.00"))
End Sub

' txtValorPolimentoEditar tela editar bloco
Private Sub txtValorPolimentoEditar_Change()
    ' Define o resultado no TextBox
    txtValorPolimentoEditar.Value = M_METODOS_GLOBAL.formatarValor(txtValorPolimentoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtValorPolimentoEditar.SelStart = Len(txtValorPolimentoEditar.Value)
                                
    ' Valor da polimento
    txtTotalPolimentoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValorServicos( _
                    txtQtdM2PolimentoEditar.Value, txtValorPolimentoEditar.Value), "0.00"))
End Sub

' txtValorADDImpostosEditar tela editar bloco
Private Sub txtValorADDImpostosEditar_Change()
    ' Define o resultado no TextBox
    txtValorADDImpostosEditar.Value = M_METODOS_GLOBAL.formatarValor(txtValorADDImpostosEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtValorADDImpostosEditar.SelStart = Len(txtValorADDImpostosEditar.Value)
                                
    ' Valor total do bloco
    txtTotalBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.custoBloco( _
                    txtValoBlocoEditar.Value, txtFreteBlocoEditar.Value, txtTotalSerradaEditar.Value, _
                    txtTotalPolimentoEditar.Value, txtValorADDImpostosEditar.Value), "0.00"))
End Sub

' txtTotalSerradaEditar tela editar bloco
Private Sub txtTotalSerradaEditar_Change()
    ' Define o resultado no TextBox
    txtTotalSerradaEditar.Value = M_METODOS_GLOBAL.formatarValor(txtTotalSerradaEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtTotalSerradaEditar.SelStart = Len(txtTotalSerradaEditar.Value)
                                
    ' Valor total do bloco
    txtTotalBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.custoBloco( _
                    txtValoBlocoEditar.Value, txtFreteBlocoEditar.Value, txtTotalSerradaEditar.Value, _
                    txtTotalPolimentoEditar.Value, txtValorADDImpostosEditar.Value), "0.00"))
End Sub

' txtTotalPolimentoEditar tela editar bloco
Private Sub txtTotalPolimentoEditar_Change()
    ' Define o resultado no TextBox
    txtTotalPolimentoEditar.Value = M_METODOS_GLOBAL.formatarValor(txtTotalPolimentoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtTotalPolimentoEditar.SelStart = Len(txtTotalPolimentoEditar.Value)
                                
    ' Valor total do bloco
    txtTotalBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.custoBloco( _
                    txtValoBlocoEditar.Value, txtFreteBlocoEditar.Value, txtTotalSerradaEditar.Value, _
                    txtTotalPolimentoEditar.Value, txtValorADDImpostosEditar.Value), "0.00"))
End Sub

' txtCustoMaterialBlocoEditar tela editar bloco
Private Sub txtCustoMaterialBlocoEditar_Change()
    ' Define o resultado no TextBox
    txtCustoMaterialBlocoEditar.Value = M_METODOS_GLOBAL.formatarValor(txtCustoMaterialBlocoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtCustoMaterialBlocoEditar.SelStart = Len(txtCustoMaterialBlocoEditar.Value)
End Sub

' txtTotalM2PolimentoBlocoEditar tela editar bloco
Private Sub txtTotalM2PolimentoBlocoEditar_Change()
    ' Define o resultado no TextBox
    txtTotalM2PolimentoBlocoEditar.Value = M_METODOS_GLOBAL.formatarMetros(txtTotalM2PolimentoBlocoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtTotalM2PolimentoBlocoEditar.SelStart = Len(txtTotalM2PolimentoBlocoEditar.Value)
                                
    ' Custo por metro
    txtCustoMaterialBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularCustoMaterial( _
                        txtTotalM2PolimentoBlocoEditar.Value, txtTotalBlocoEditar.Value), "0.00"))
End Sub

' txtTotalBlocoEditar tela editar bloco
Private Sub txtTotalBlocoEditar_Change()
    ' Define o resultado no TextBox
    txtTotalBlocoEditar.Value = M_METODOS_GLOBAL.formatarValor(txtTotalBlocoEditar.Value)
    
    ' Move o cursor para o final do TextBox
    txtTotalBlocoEditar.SelStart = Len(txtTotalBlocoEditar.Value)
                                
    ' Custo por metro
    txtCustoMaterialBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularCustoMaterial( _
                        txtTotalM2PolimentoBlocoEditar.Value, txtTotalBlocoEditar.Value), "0.00"))
End Sub

' Carrega os campos com os dados do bloco tela editar bloco
Private Sub carregarDadosBlocoTelaEdicaoBloco(bloco As objBloco)
    ' Exibir o resultado da pesquisa
    ' Descri��o e dimens�es finais
    txtIdBlocoEditar.Value = bloco.idSistema
    txtMaterialEditar.Value = bloco.nomeMaterial
    cbTipoMaterialEditar.Value = bloco.tipoMaterial.nome
    txtObsEditar.Value = bloco.observacao
    cbPedreiraEditar.Value = bloco.pedreira.nome
    cbSerrariaEditar.Value = bloco.serraria.nome
    cbPolideiraEditar.Value = bloco.polideira.nome
    txtNBlocoPedreiraEditar.Value = bloco.numeroBlocoPedreira
    cbEstoqueEditar.Value = bloco.estoque.nome
    txtDataCadastroEditar.Value = bloco.dataCadastro
    txtQtdM3blocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.qtdM3, "0.0000"))
    txtQtdM2SerradaEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.qtdM2Serrada, "0.0000"))
    txtQtdM2PolimentoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.qtdM2Polimento, "0.0000"))
    txtTotalChapaBlocoEditar.Value = bloco.qtdChapas
    cbStatusBlocoEditar.Value = bloco.status.nome
    cbNotaBlocoEditar.Value = bloco.nota
    cbCustoMedioEditar.Value = bloco.consultarCustoMedio
    ' Dimens�es bloco e m�dias chapas
    txtCompBrutaBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.compBrutoBloco, "0.0000"))
    txtAltBrutaBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.altBrutoBloco, "0.0000"))
    txtLArgBrutaBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.largBrutoBloco, "0.0000"))
    txtCompLiquidoBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.compLiquidoBloco, "0.0000"))
    txtAltLiquidoBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.altLiquidoBloco, "0.0000"))
    txtLArgLiquidoBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.largLiquidoBloco, "0.0000"))
    txtCompBrutaBrutoChapaEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.compBrutoChapaBruta, "0.0000"))
    txtAltBrutaBrutoChapaEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.altBrutoChapaBruta, "0.0000"))
    txtCompBrutaliquidoChapaEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.compLiquidoChapaBruta, "0.0000"))
    txtAltBrutaLiquidoChapaEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.altBrutoChapaBruta, "0.0000"))
    txtCompPolidaBrutoChapaEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.compBrutoChapaPolida, "0.0000"))
    txtAltPolidaBrutoChapaEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.altBrutoChapaPolida, "0.0000"))
    txtCompPolidaLiquidoChapaEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.compLiquidoChapaPolida, "0.0000"))
    txtAltPolidaLiquidaChapaEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.altBrutoChapaPolida, "0.0000"))
    ' Valores
    txtValoBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.valorBloco, "0.00"))
    txtPrecoBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.precoM3Bloco, "0.00"))
    txtFreteBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.freteBloco, "0.00"))
    txtValorSerradaEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.valorMetroSerrada, "0.00"))
    txtValorPolimentoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.valorMetroPolimento, "0.00"))
    txtValorADDImpostosEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.valoresAdicionais, "0.00"))
    txtTotalSerradaEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.valorTotalSerrada, "0.00"))
    txtTotalPolimentoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.valorTotalPolimento, "0.00"))
    ' Custos
    txtCustoMaterialBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.custoMaterial, "0.00"))
    txtTotalM2PolimentoBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.qtdM2Polimento, "0.00"))
    txtTotalBlocoEditar.Value = M_METODOS_GLOBAL.formatarComPontos(Format(bloco.valorTotalBloco, "0.00"))
    
    ' Se Status do bloco for finalizado deixar visivel lBlocoFinalizado e cbAbrirBlocoEditar e desabilitar todos os campos
    If bloco.status.nome = "FECHADO" Then
        cbAbrirBlocoEditar.Visible = True
        lBlocoFinalizado.Visible = True
    End If
End Sub
' Habilita e desabilita campos para edi��o tela editar bloco
Private Sub cbAbrirBlocoEditar_Click()
    If cbAbrirBlocoEditar.Value = True Then
        Call habilitaCamposBlocoEditar
    Else
        Call desabilitaCamposBlocoEditar
    End If
End Sub
' Bot�o btnLTxtSalvarEdicaoBloco tela editar bloco
Private Sub btnLTxtSalvarEdicaoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Verifica o N�mero do bloco na pedreira
    If txtNBlocoPedreiraEditar.Value = "" Or txtNBlocoPedreiraEditar.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtNBlocoPedreiraEditar, NUMERO_BLOCO_PEDREIRA_MENSAGEM, NUMERO_BLOCO_PEDREIRA_TITULO
        ' Para o fluxo do sistema para a corre��o
        Exit Sub
    End If
    ' Deixa na cor patr�o
    errorStyle.sairErrorStyleTextBox txtNBlocoPedreiraEditar
    
    ' Verifica nome do bloco
    If txtMaterialEditar.Value = "" Or txtMaterialEditar.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtMaterialEditar, NOME_BLOCO_MENSAGEM, NOME_BLOCO_PEDREIRA_TITULO
        ' Para o fluxo do sistema para a corre��o
        Exit Sub
    End If
    ' Deixa na cor patr�o
    errorStyle.sairErrorStyleTextBox txtMaterialEditar
    
    ' Cria��o dos objetos
    Set pedreira = daoPedreira.pesquisarPorNome(cbPedreiraEditar.Value)
    Set serraria = daoSerrada.pesquisarPorNome(cbSerrariaEditar.Value)
    Set polideira = daoPolideira.pesquisarPorNome(cbPolideiraEditar)
    Set tipoMaterial = daoTipoMaterial.pesquisarPorNome(cbTipoMaterialEditar.Value)
    Set statusObj = daoStatus.pesquisarPorNome(cbStatusBlocoEditar)
    Set estoque = daoEstoqueM3.pesquisarPorNome(cbEstoqueEditar.Value)
    Set bloco = ObjectFactory.factoryBloco(bloco)
    
    ' Cria��o do objeto
    bloco.carregarBlocoEdicao txtIdBlocoEditar.Value, txtMaterialEditar.Value, txtObsEditar.Value, txtNBlocoPedreiraEditar.Value, estoque, _
                    txtDataCadastroEditar.Value, txtQtdM3blocoEditar.Value, txtQtdM2SerradaEditar.Value, txtQtdM2PolimentoEditar.Value, txtTotalChapaBlocoEditar.Value, cbNotaBlocoEditar.Value, _
                    cbCustoMedioEditar.Value, txtCompBrutaBlocoEditar.Value, txtAltBrutaBlocoEditar.Value, txtLArgBrutaBlocoEditar.Value, txtCompLiquidoBlocoEditar.Value, _
                    txtAltLiquidoBlocoEditar.Value, txtLArgLiquidoBlocoEditar.Value, txtCompBrutaBrutoChapaEditar.Value, txtAltBrutaBrutoChapaEditar.Value, _
                    txtCompBrutaliquidoChapaEditar.Value, txtAltBrutaLiquidoChapaEditar.Value, txtCompPolidaBrutoChapaEditar.Value, txtAltPolidaBrutoChapaEditar.Value, _
                    txtCompPolidaLiquidoChapaEditar.Value, txtAltPolidaLiquidaChapaEditar.Value, txtValoBlocoEditar.Value, txtPrecoBlocoEditar.Value, _
                    txtFreteBlocoEditar.Value, txtValorSerradaEditar.Value, txtValorPolimentoEditar.Value, txtValorADDImpostosEditar.Value, _
                    txtTotalSerradaEditar.Value, txtTotalPolimentoEditar.Value, txtCustoMaterialBlocoEditar.Value, txtTotalBlocoEditar.Value, _
                    statusObj, tipoMaterial, pedreira, serraria, polideira
    
    ' Chama servi�o para cadastrar do bloco
    Call daoBloco.cadastrarEEditar(bloco)
    
    ' Chama servi�o para pesquisa do bloco
    Set bloco = daoBloco.pesquisarPorId(bloco.idSistema) ' Envia o id do bloco
    
    ' Recarrega os dados na tela editar bloco
    Call carregarDadosBlocoTelaEdicaoBloco(bloco)
    
    ' Libera espa�o da memorio
    Set pedreira = Nothing
    Set serraria = Nothing
    Set polideira = Nothing
    Set tipoMaterial = Nothing
    Set statusObj = Nothing
    Set estoque = Nothing
    Set bloco = Nothing
    
    ' Mensagem de edi��o realizada com sucesso. Mensagem de erro utilizada para sucesso na opera��o
    errorStyle.Informativo SUCESSO_EDICAO_MENSAGEM, SUCESSO_EDICAO_TITULO
    ' Seta o foco
    txtMaterialEditar.SetFocus
End Sub
' Bot�o btnLTxtVoltarEdicaoBloco tela editar bloco
Private Sub btnLTxtVoltarEdicaoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 1
    ' Seta o foco
    txtMaterialBlocoPesquisa.SetFocus
End Sub

'-----------------------------------------------------------------TELA ESTOQUE M�-----------------------------------
'                                                                 ---------------
' Efeito de label nome do pdf tela estoque m�
Private Sub lDigiteNomeArquivoM2Explemplo_Click()
    lDigiteNomeArquivoM2.Visible = True
    lDigiteNomeArquivoM2Explemplo.Visible = False
    txtNomeArquivoEstoqueChapas.SetFocus
End Sub
' Efeito e coloca em caixa alta o texto em txtNomeArquivoEstoqueChapas tela estoque m�
Private Sub txtNomeArquivoEstoqueChapas_Change()
    lDigiteNomeArquivoM2.Visible = True
    lDigiteNomeArquivoM2Explemplo.Visible = False

    If txtNomeArquivoEstoqueChapas.Value = "" Then
        lDigiteNomeArquivoM2.Visible = False
        lDigiteNomeArquivoM2Explemplo.Visible = True
    End If

    txtNomeArquivoEstoqueChapas.Value = UCase(txtNomeArquivoEstoqueChapas.Value)
End Sub
' Efeito ao sair da caixa txtNomeArquivoEstoqueChapas de texto tela estoque m�
Private Sub txtNomeArquivoEstoqueChapas_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtNomeArquivoEstoqueChapas.Value = "" Then
        lDigiteNomeArquivoM2.Visible = False
        lDigiteNomeArquivoM2Explemplo.Visible = True
    End If
End Sub
' Efeito para quando sair do foco de txtNomeArquivoEstoqueChapas de texto tela estoque m�
Private Sub fTiraEfeitoBotoesExportarChapasM2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtNomeArquivoEstoqueChapas.Value = "" Then
        lDigiteNomeArquivoM2.Visible = False
        lDigiteNomeArquivoM2Explemplo.Visible = True
    End If
End Sub
' txtMaterialChapaPesquisa tela estoque m�
Private Sub txtMaterialChapaPesquisa_Change()
    ' Coloca tudo em caixa alta
    txtMaterialChapaPesquisa.Value = UCase(txtMaterialChapaPesquisa.Value)
End Sub
' txtIdBlocoChapaPesquisa tela estoque m�
Private Sub txtIdBlocoChapaPesquisa_Change()
    ' Coloca tudo em caixa alta
    txtIdBlocoChapaPesquisa.Value = UCase(txtIdBlocoChapaPesquisa.Value)
End Sub
' txtIdchapaEstoque tela estoque m�
Private Sub txtIdchapaEstoque_Change()
    ' Coloca tudo em caixa alta
    txtIdchapaEstoque.Value = UCase(txtIdchapaEstoque.Value)
End Sub
' Bot�o btnLTxtPesquisarChapas tela estoque m�
Private Sub btnLTxtPesquisarChapas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o pesquisar chapa, tela estoque m�"
    ' Seta o foco
    txtMaterialChapaPesquisa.SetFocus
End Sub
' Bot�o btnLTxtLimparFiltrosChapas tela estoque m�
Private Sub btnLTxtLimparFiltrosChapas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    Call limparCamposPesquisaEstoqueM2
End Sub
' Bot�o btnLImgExportarEstoqueM2 tela estoque m�
Private Sub btnLImgExportarEstoqueM2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o esportar estoque m�, tela estoque m�"
End Sub
'Bot�o btnLTxtNovoAvulso tela estoque m�
Private Sub btnLTxtNovoAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Coloca data atual na txtDataCadastroChapaAvulsa na tela cadastro chapa avulso
    txtDataCadastroChapaAvulsa.Value = Date
    
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 5
    ' Seta o foco
    txtIdBlocoAvulso.SetFocus
    
    ' Chama metodo para carregar comboBox
    Call carregarTiposMateriais(Me.cbTipoMaterialL)
    Call carregarTiposPolimento(Me.cbTipoPolimentoL)
    Call carregarTemNota(Me.cbTemNotaAvulso)
    
'    ' Pesquisa blocos cadastrado no dia atual
'    Set listaObjeto = daoBloco.listarBlocosFilter(Date, Date, "", "", "", "", "", "", "", "", "", "", "")
'
'    ' Chama metodo para carregar lista e blocos cadastros do dia atual
'    Call carregarList(Me.listCadastradosHoje, listaObjeto)
'
'    ' Chama metodo para carregar lista
'    Call carregarList(Me.ListMaterias)
End Sub
' Bot�o btnLTxtEditarChapa tela estoque m�
Private Sub btnLTxtEditarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 6
    ' Seta n�mero de pagina para poder voltar
    paginaAnterior = 4
    ' Seta o foco
    cbPolideiraChapa.SetFocus
    
    ' Chama servi�o para pesquisa da chapa
    
    
    ' Carrega os ComboBox da tela
    Call carregarPolideiras(cbPolideiraChapa)
    Call carregarTiposPolimento(cbTipoPolimentoChapa)
    Call carregarTiposMateriais(cbTipoMaterialChapaC)
    Call carregarEstoque(cbEstoqueChapaC)
    Call carregarTiposMateriais(cbTiposMateriaisChapas)
    
    ' Carrega os dados na tela editar chapa
    Call carregarDadosChapaTelaEdicaoChapa ' Ir� enviar o objeto chapa para poder carregar os campos
End Sub
' Bot�o btnLTxtTrocaEstoque tela estoque m�
Private Sub btnLTxtTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 7
    ' Seta o foco
    txtQtdMaterialParaTroca01.SetFocus
End Sub

'-----------------------------------------------------------------TELA CADASTRO AVULSO-----------------------------------
'                                                                 --------------------
' txtIdBlocoAvulso tela cadastro avulso
Private Sub txtIdBlocoAvulso_Change()
    ' Coloca tudo em caixa alta
    txtIdBlocoAvulso.Value = UCase(txtIdBlocoAvulso.Value)
    
    ' Cria o c�digo para o sistema
    txtIdBlocoAvulsoSistema.Value = txtIdBlocoAvulso & "-" & M_METODOS_GLOBAL.ExtrairUltimaPalavra( _
                txtMaterialAvulso.Value) & "-BL"
    
    ' Deixa em branco o codigo se as variaveis forem vazias
    If txtIdBlocoAvulso.Value = "" And txtMaterialAvulso.Value = "" Then
        txtIdBlocoAvulsoSistema.Value = ""
    End If
End Sub
' txtMaterialAvulso tela cadastro avulso
Private Sub txtMaterialAvulso_Change()
    ' Coloca tudo em caixa alta
    txtMaterialAvulso.Value = UCase(txtMaterialAvulso.Value)
    
    ' Cria o c�digo para o sistema
    txtIdBlocoAvulsoSistema.Value = txtIdBlocoAvulso & "-" & M_METODOS_GLOBAL.ExtrairUltimaPalavra( _
                txtMaterialAvulso.Value) & "-BL"
    
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
    txtComprimentoChapaAvulsa.Value = M_METODOS_GLOBAL.formatarMetros(txtComprimentoChapaAvulsa)
    
    ' Seta o valor no comprimento bruto
    txtCompChapasBrutasAvulso.Value = txtComprimentoChapaAvulsa.Value

    'Move o cursor para o final do TextBox
    txtComprimentoChapaAvulsa.SelStart = Len(txtComprimentoChapaAvulsa.Value)
    
    'Retorna valor calculado e formatado
    txtTotalM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM2(txtComprimentoChapaAvulsa.Value, _
        txtAlturaChapaAvulsa.Value, txtQuantidadeChapasAvulsas.Value), "0.0000"))
        
    'Seta o custo do material m�
    txtCustoSimplesM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.custoMaterialM2( _
            txtTotalBlocoAvulso.Value, txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", _
            txtTotalM2Avulso.Value), "0.00"))
End Sub
' txtAlturaChapaAvulsa tela cadastro avulso
Private Sub txtAlturaChapaAvulsa_Change()
   'Define o resultado no TextBox
    txtAlturaChapaAvulsa.Value = M_METODOS_GLOBAL.formatarMetros(txtAlturaChapaAvulsa)
    
    ' Seta o valor na altura bruto
    txtAlturaChapasBrutasAvulso.Value = txtAlturaChapaAvulsa.Value

    'Move o cursor para o final do TextBox
    txtAlturaChapaAvulsa.SelStart = Len(txtAlturaChapaAvulsa.Value)
    
    'Retorna valor calculado e formatado
    txtTotalM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM2( _
        txtComprimentoChapaAvulsa.Value, txtAlturaChapaAvulsa.Value, txtQuantidadeChapasAvulsas.Value), "0.0000"))
        
    'Seta o custo do material m�
    txtCustoSimplesM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.custoMaterialM2( _
            txtTotalBlocoAvulso.Value, txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", _
            txtTotalM2Avulso.Value), "0.00"))
End Sub
' txtQuantidadeChapasAvulsas tela cadastro avulso
Private Sub txtQuantidadeChapasAvulsas_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Deixa s� a digita��o de numero
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub
' txtQuantidadeChapasAvulsas tela cadastro avulso
Private Sub txtQuantidadeChapasAvulsas_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Se n�o tiver valor deixa com 0
    If txtQuantidadeChapasAvulsas.Value = "" Then
        txtQuantidadeChapasAvulsas.Value = 0
    End If
    
    'Retorna valor calculado e formatado
    txtTotalM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM2( _
        txtComprimentoChapaAvulsa.Value, txtAlturaChapaAvulsa.Value, txtQuantidadeChapasAvulsas.Value), "0.0000"))
    'Seta o custo do material m�
    txtCustoSimplesM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.custoMaterialM2( _
            txtTotalBlocoAvulso.Value, txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", _
            txtTotalM2Avulso.Value), "0.00"))
End Sub
' txtCompChapasBrutasAvulso tela cadastro avulso
Private Sub txtCompChapasBrutasAvulso_Change()
   'Define o resultado no TextBox
    txtCompChapasBrutasAvulso.Value = M_METODOS_GLOBAL.formatarMetros(txtCompChapasBrutasAvulso)

    'Move o cursor para o final do TextBox
    txtCompChapasBrutasAvulso.SelStart = Len(txtCompChapasBrutasAvulso.Value)
End Sub
' txtAlturaChapasBrutasAvulso tela cadastro avulso
Private Sub txtAlturaChapasBrutasAvulso_Change()
   'Define o resultado no TextBox
    txtAlturaChapasBrutasAvulso.Value = M_METODOS_GLOBAL.formatarMetros(txtAlturaChapasBrutasAvulso)

    'Move o cursor para o final do TextBox
    txtAlturaChapasBrutasAvulso.SelStart = Len(txtAlturaChapasBrutasAvulso.Value)
End Sub
' txtAdicionaisAvulso tela cadastro avulso
Private Sub txtAdicionaisAvulso_Change()
    ' Define o resultado no TextBox
    txtAdicionaisAvulso.Value = M_METODOS_GLOBAL.formatarValor(txtAdicionaisAvulso.Value)
    
    ' Move o cursor para o final do TextBox
    txtAdicionaisAvulso.SelStart = Len(txtAdicionaisAvulso.Value)
    
    'Seta o custo do material m�
    txtCustoSimplesM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.custoMaterialM2( _
            txtTotalBlocoAvulso.Value, txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", _
            txtTotalM2Avulso.Value), "0.00"))
End Sub
' txtValorFreteAvulso tela cadastro avulso
Private Sub txtValorFreteAvulso_Change()
    ' Define o resultado no TextBox
    txtValorFreteAvulso.Value = M_METODOS_GLOBAL.formatarValor(txtValorFreteAvulso.Value)

    ' Move o cursor para o final do TextBox
    txtValorFreteAvulso.SelStart = Len(txtValorFreteAvulso.Value)
    
    'Seta o custo do material m�
    txtCustoSimplesM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.custoMaterialM2( _
            txtTotalBlocoAvulso.Value, txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", _
            txtTotalM2Avulso.Value), "0.00"))
End Sub
' txtValorBlocoAvulso tela cadastro avulso
Private Sub txtValorMetroAvulso_Change()
    ' Define o resultado no TextBox
    txtValorMetroAvulso.Value = M_METODOS_GLOBAL.formatarValor(txtValorMetroAvulso.Value)
    
    ' Move o cursor para o final do TextBox
    txtValorMetroAvulso.SelStart = Len(txtValorMetroAvulso.Value)

    ' Retorna valor calculado e formatado
    txtTotalBlocoAvulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValorBloco( _
            txtTotalM2Avulso.Value, txtValorMetroAvulso.Value), "0.00"))
            
    'Se m� for diferente de 0 calcula o custo do material
    'Seta o custo do material m�
    txtCustoSimplesM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(Util.custoMaterialM2(txtTotalBlocoAvulso.Value, _
            txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", txtTotalM2Avulso.Value), "0.00"))
End Sub
' txtTotalM2Avulso tela cadastro avulso
Private Sub txtTotalM2Avulso_Change()
    'Se m� for diferente de 0 calcula o custo do material
    'Seta o custo do material m�
    txtCustoSimplesM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.custoMaterialM2( _
            txtTotalBlocoAvulso.Value, txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", _
            txtTotalM2Avulso.Value), "0.00"))
End Sub
' Bot�o btnLImgCadastrarMaterialAvulso tela cadastro avulso
Private Sub btnLImgCadastrarMaterialAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar tipo material, tela cadastro avulso"
End Sub
' Bot�o btnLImgCadastrarPolimentoAvulso tela cadastro avulso
Private Sub btnLImgCadastrarPolimentoAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar tipo polimento, tela cadastro avulso"
End Sub
' Bot�o btnLTxtCadastrarChapaAvulso tela cadastro avulso
Private Sub btnLTxtCadastrarChapaAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar chapas avulsos, tela cadastro avulso"
    ' Seta o foco
    txtIdBlocoAvulso.SetFocus
End Sub
' Bot�o btnLTxtVoltarCadatradoChapasAvulso tela cadastro avulso
Private Sub btnLTxtVoltarCadatradoChapasAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage - tela estoque m�
    Me.MultiPageCEBC.Value = 4
    ' Seta o foco
    txtMaterialChapaPesquisa.SetFocus
End Sub
' Bot�o btnLTxtLimparCadastroChapaAvulso tela cadastro avulso
Private Sub btnLTxtLimparCadastroChapaAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    Call limparCamposCadastroAvulso
End Sub

'-----------------------------------------------------------------TELA LAN�AMENTO E EDI��O CHAPA-----------------------------------
'                                                                 ------------------------------
' Carrega os campos com os dados da chapa tela lan�amento e edi��o chapa
Private Sub carregarDadosChapaTelaEdicaoChapa() ' Ir� receber o objeto Chapa para poder carregar os campos e algum campos do bloco
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
    
    ' Dimens�es e custos
    Call selecaoItem("cbPolideiraChapa", "S�O ROQUE")
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
    Call carregarListTamanhosChapas(ListTamanhosChapas) ' Ir� enviar id chapa para carregamento
   
    ' Se Status do bloco for finalizado deixar visivel lBlocoFinalizadoChapa e cbAbrirParaEdicao e desabilitar todos os campos
    
End Sub
' Bot�o btnLImgCadastrarPolideiraChapa tela lan�amento e edi��o chapa
Private Sub btnLImgCadastrarPolideiraChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar Polideira, tela lan�amento e edi��o chapa"
End Sub
' Bot�o btnLImgCadastrarTipoPolideiraChapa tela lan�amento e edi��o chapa
Private Sub btnLImgCadastrarTipoPolideiraChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar tipo polimento, tela lan�amento e edi��o chapa"
End Sub
' Bot�o btnLImgCadastrarTipoMaterialChapa tela lan�amento e edi��o chapa
Private Sub btnLImgCadastrarTipoMaterialChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o tipo material, tela lan�amento e edi��o chapa"
End Sub
' Bot�o btnLImgCadastrarTipoMaterialChapaTamanhos tela lan�amento e edi��o chapa
Private Sub btnLImgCadastrarTipoMaterialChapaTamanhos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o tipo material, tela lan�amento e edi��o chapa"
End Sub
' Bot�o btnLTxtAdicionarTamanhoChapa tela lan�amento e edi��o chapa
Private Sub btnLTxtAdicionarTamanhoChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o adicionar tamanhos, tela lan�amento e edi��o chapa"
    ' Seta o foco
    cbTiposMateriaisChapas.SetFocus
End Sub
' Bot�o btnLTxtEditarTamanhoChapa tela lan�amento e edi��o chapa
Private Sub btnLTxtEditarTamanhoChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o editar tamanho chapa, tela lan�amento e edi��o chapa"
    ' Seta o foco
    cbTiposMateriaisChapas.SetFocus
End Sub
' Bot�o btnLTxtTirarDaLista tela lan�amento e edi��o chapa
Private Sub btnLTxtTirarDaLista_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o tira tamanho da lista, tela lan�amento e edi��o chapa"
    ' Seta o foco
    cbTiposMateriaisChapas.SetFocus
End Sub
' Bot�o btnLTxtSalvarChapa tela lan�amento e edi��o chapa
Private Sub btnLTxtSalvarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o salva altera��o da chapa, tela lan�amento e edi��o chapa"
    ' Seta o foco
    cbPolideiraChapa.SetFocus
End Sub
' Bot�o btnLTxtVoltarChapa tela lan�amento e edi��o chapa
Private Sub btnLTxtVoltarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Volta para tela que chamou
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = paginaAnterior
        
    If paginaAnterior = 1 Then
        ' Bot�o chapa
        formControle.Controls("btnLMenuBloco").BackColor = RGB(200, 230, 255)
        formControle.Controls("btnLMenuBloco").Font.Size = 32
        formControle.Controls("btnLMenuBloco").Font.Size = 20
        formControle.Controls("btnLMenuBloco").Left = 15
        formControle.Controls("btnLMenuBloco").Width = 172
        formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignCenter
                    
        ' Bot�o Menu
        formControle.Controls("btnLMenuChapa").BackColor = RGB(0, 100, 200)
        formControle.Controls("btnLMenuChapa").Left = 2
        formControle.Controls("btnLMenuChapa").Width = 189
        formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignLeft
        
            ' Seta o foco
        txtMaterialBlocoPesquisa.SetFocus
    Else
        ' Seta o foco
        txtMaterialChapaPesquisa.SetFocus
    End If

End Sub

'-----------------------------------------------------------------TELA TROCA ESTOQUE-----------------------------------
'                                                                 ------------------
' Bot�o btnLTxtAdicionarTrocaEstoque tela troca estoque
Private Sub btnLTxtAdicionarTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o adicionar material para troca, tela troca estoque"
    ' Seta o foco
    txtMaterialParaTroca02.SetFocus
End Sub
' Bot�o btnLTxtTrocarEstoque tela troca estoque
Private Sub btnLTxtTrocarEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o troca de estoque, tela troca estoque"
End Sub
' Bot�o btnLTxtVoltarTrocaEstoque tela troca estoque
Private Sub btnLTxtVoltarTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o voltar, tela troca estoque"
    ' Seta o foco
    txtMaterialChapaPesquisa.SetFocus
End Sub

'-----------------------------------------------------------------TELA DESPACHE-----------------------------------
'                                                                 -------------
' Bot�o btnLImgCadastrarMotoristaDespache tela despache
Private Sub btnLImgCadastrarMotoristaDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastro motorista, tela despache"
End Sub
' Bot�o btnLImgCadastrarDestinoDespache tela despache
Private Sub btnLImgCadastrarDestinoDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastro destino, tela despache"
End Sub
' Bot�o btnLTxtAdicionar tela despache
Private Sub btnLTxtAdicionar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o adicionar chapa, tela despache"
    ' Seta o foco
    txtMaterial.SetFocus
End Sub
' Bot�o btnLTxtDespachar tela despache
Private Sub btnLTxtDespachar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o despachar, tela despache"
End Sub
' Bot�o btnLTxtLimparDespache tela despache
Private Sub btnLTxtLimparDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o limpar dados, tela despache"
End Sub

'-----------------------------------------------------------------TELA CARREGOS-----------------------------------
'                                                                 -------------
' Bot�o btnLTxtPesquisarCarregos tela carregos
Private Sub btnLTxtPesquisarCarregos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o pesquisar por carregos, tela carregos"
    ' Seta o foco
    cbMotoristaL.SetFocus
End Sub
' Bot�o btnLTxtLimparListas tela carregos
Private Sub btnLTxtLimparListas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o limpar dados filtro, tela carregos"
'    ' Seta o foco
'    cbMotoristaL.SetFocus
End Sub
' Bot�o btnLImgExportarCarregoPDF tela carregos
Private Sub btnLImgExportarCarregoPDF_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o exportar carregos em pdf, tela carregos"
    ' Seta o foco
'    cbMotoristaL.SetFocus
End Sub
' Bot�o btnLTxtEditarCarrego tela carregos
Private Sub btnLTxtEditarCarrego_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o editar carrego, tela carregos"
    ' Seta o foco
    cbMotoristaL.SetFocus
End Sub
' Bot�o btnLTxtVoltarCArrego tela carregos
Private Sub btnLTxtVoltarCArrego_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o voltar, tela carregos"
    ' Seta o foco
    cbMotoristaL.SetFocus
End Sub

'-----------------------------------------------------------------TELA CADASTROS DIVERSOS-----------------------------------
'                                                                 -----------------------
' Bot�o btnLTxtSalvarPedreira tela cadastros diversos
Private Sub btnLTxtSalvarPedreira_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar pedreira, tela cadastros diversos"
End Sub
' Bot�o btnLTxtSalvarSerraria tela cadastros diversos
Private Sub btnLTxtSalvarSerraria_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar serraria, tela cadastros diversos"
End Sub
' Bot�o btnLTxtSalvarPolideira tela cadastros diversos
Private Sub btnLTxtSalvarPolideira_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar polideira, tela cadastros diversos"
End Sub
' Bot�o btnLTxtSalvarTipoMaterial tela cadastros diversos
Private Sub btnLTxtSalvarTipoMaterial_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar tipo material, tela cadastros diversos"
End Sub
' Bot�o btnLTxtSalvarTipoPolimento tela cadastros diversos
Private Sub btnLTxtSalvarTipoPolimento_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar tipo polimento, tela cadastros diversos"
End Sub
' Bot�o btnLTxtSalvarMotorista tela cadastros diversos
Private Sub btnLTxtSalvarMotorista_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar motorista, tela cadastros diversos"
End Sub
' Bot�o btnLTxtSalvarMotorista tela cadastros diversos
Private Sub btnLTxtSalvarDestino_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar destino, tela cadastros diversos"
End Sub

'-----------------------------------------------------------------TELA USUARIO-----------------------------------
'                                                                 ------------
' Bot�o btnLTxtSalvarUsuario tela usuarios
Private Sub btnLTxtSalvarUsuario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar usu�rio, tela usuarios"
End Sub
' Bot�o btnLTxtListUsuario tela usuarios
Private Sub btnLTxtListUsuario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o carrega lista com usu�rios, tela usuarios"
End Sub
' Bot�o btnLTxtListUsuarioLog tela usuarios
Private Sub btnLTxtListUsuarioLog_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Servi�o
    MsgBox "Chama Servi�o carrega lista com log dos usu�rios, tela usuarios"
End Sub

'-----------------------------------------------------------------DESABILITA  E HABILITA CAMPOS-----------------------------------
'                                                                 -----------------------------
' Desabilita campos da tela editar bloco
Private Sub desabilitaCamposBlocoEditar()
    txtIdBlocoEditar.Enabled = False
    txtMaterialEditar.Enabled = False
    cbTipoMaterialEditar.Enabled = False
    txtObsEditar.Enabled = False
    cbPedreiraEditar.Enabled = False
    cbSerrariaEditar.Enabled = False
    cbPolideiraEditar.Enabled = False
    txtNBlocoPedreiraEditar.Enabled = False
    cbEstoqueEditar.Enabled = False
    txtDataCadastroEditar.Enabled = False
    txtQtdM3blocoEditar.Enabled = False
    txtQtdM2SerradaEditar.Enabled = False
    txtQtdM2PolimentoEditar.Enabled = False
    txtTotalChapaBlocoEditar.Enabled = False
    cbStatusBlocoEditar.Enabled = False
    cbNotaBlocoEditar.Enabled = False
    cbCustoMedioEditar.Enabled = False
    
    ' Dimens�es bloco e m�dias chapas
    txtCompBrutaBlocoEditar.Enabled = False
    txtAltBrutaBlocoEditar.Enabled = False
    txtLArgBrutaBlocoEditar.Enabled = False
    txtCompLiquidoBlocoEditar.Enabled = False
    txtAltLiquidoBlocoEditar.Enabled = False
    txtLArgLiquidoBlocoEditar.Enabled = False
    txtCompBrutaBrutoChapaEditar.Enabled = False
    txtAltBrutaBrutoChapaEditar.Enabled = False
    txtCompBrutaliquidoChapaEditar.Enabled = False
    txtAltBrutaLiquidoChapaEditar.Enabled = False
    txtCompPolidaBrutoChapaEditar.Enabled = False
    txtAltPolidaBrutoChapaEditar.Enabled = False
    txtCompPolidaLiquidoChapaEditar.Enabled = False
    txtAltPolidaLiquidaChapaEditar.Enabled = False
    
    ' Valores
    txtValoBlocoEditar.Enabled = False
    txtPrecoBlocoEditar.Enabled = False
    txtFreteBlocoEditar.Enabled = False
    txtValorSerradaEditar.Enabled = False
    txtValorPolimentoEditar.Enabled = False
    txtValorADDImpostosEditar.Enabled = False
    txtTotalSerradaEditar.Enabled = False
    txtTotalPolimentoEditar.Enabled = False
    
    ' Custos
    txtCustoMaterialBlocoEditar.Enabled = False
    txtTotalM2PolimentoBlocoEditar.Enabled = False
    txtTotalBlocoEditar.Enabled = False
End Sub

' Habilita campos da tela editar bloco
Private Sub habilitaCamposBlocoEditar()
    txtIdBlocoEditar.Enabled = True
    txtMaterialEditar.Enabled = True
    cbTipoMaterialEditar.Enabled = True
    txtObsEditar.Enabled = True
    cbPedreiraEditar.Enabled = True
    cbSerrariaEditar.Enabled = True
    cbPolideiraEditar.Enabled = True
    txtNBlocoPedreiraEditar.Enabled = True
    cbEstoqueEditar.Enabled = True
    txtDataCadastroEditar.Enabled = True
    txtQtdM3blocoEditar.Enabled = True
    txtQtdM2SerradaEditar.Enabled = True
    txtQtdM2PolimentoEditar.Enabled = True
    txtTotalChapaBlocoEditar.Enabled = True
    cbStatusBlocoEditar.Enabled = True
    cbNotaBlocoEditar.Enabled = True
    cbCustoMedioEditar.Enabled = True
    
    ' Dimens�es bloco e m�dias chapas
    txtCompBrutaBlocoEditar.Enabled = True
    txtAltBrutaBlocoEditar.Enabled = True
    txtLArgBrutaBlocoEditar.Enabled = True
    txtCompLiquidoBlocoEditar.Enabled = True
    txtAltLiquidoBlocoEditar.Enabled = True
    txtLArgLiquidoBlocoEditar.Enabled = True
    txtCompBrutaBrutoChapaEditar.Enabled = True
    txtAltBrutaBrutoChapaEditar.Enabled = True
    txtCompBrutaliquidoChapaEditar.Enabled = True
    txtAltBrutaLiquidoChapaEditar.Enabled = True
    txtCompPolidaBrutoChapaEditar.Enabled = True
    txtAltPolidaBrutoChapaEditar.Enabled = True
    txtCompPolidaLiquidoChapaEditar.Enabled = True
    txtAltPolidaLiquidaChapaEditar.Enabled = True
    
    ' Valores
    txtValoBlocoEditar.Enabled = True
    txtPrecoBlocoEditar.Enabled = True
    txtFreteBlocoEditar.Enabled = True
    txtValorSerradaEditar.Enabled = True
    txtValorPolimentoEditar.Enabled = True
    txtValorADDImpostosEditar.Enabled = True
    txtTotalSerradaEditar.Enabled = True
    txtTotalPolimentoEditar.Enabled = True
    
    ' Custos
    txtCustoMaterialBlocoEditar.Enabled = False
    txtTotalM2PolimentoBlocoEditar.Enabled = False
    txtTotalBlocoEditar.Enabled = False
End Sub

'-----------------------------------------------------------------LIMPAR CAMPOS-----------------------------------
'                                                                 -------------
' Limpa os campos de pesquisa da tela estoque M�
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

' Limpa os campos de pesquisa da tela estoque M�
Private Sub limparCamposPesquisaEstoqueM2()
    txtMaterialChapaPesquisa.Value = ""
    txtIdBlocoChapaPesquisa.Value = ""
    txtIdchapaEstoque.Value = ""
    cbPolideiraChapaPesquisa.Value = ""
    cbTipoPolimentoPesquisa.Value = ""
    obEstoqueZeroSim.Value = False
    obEstoqueZeroNao.Value = True
    txtNomeArquivoEstoqueChapas.Value = ""
    ' Seta o foco
    txtMaterialChapaPesquisa.SetFocus
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
    
    ' Seta o foco
    cbPedreira.SetFocus
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
    ' Seta o foco
    txtIdBlocoAvulso.SetFocus
End Sub

'-----------------------------------------------------------------CARREAGMENTO DOS COMBOBBOX-----------------------------------
'                                                                 --------------------------
' Sele��o do item da comboBox
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
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim i As Integer
    
    ' Criando a lista
    Set listaObjetos = daoPedreira.listarPedreiras

    ' limpa a lista para carregamento
    cbPedreiras.Clear
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se n�o tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Carregamento com primeiro item vazio
        cbPedreiras.AddItem ""
        
        ' Loop atrav�s dos itens da cole��o
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set pedreira = listaObjetos(i)
            ' Carregamento para lista
            cbPedreiras.AddItem pedreira.nome
            ' Libera espa�o memoria
            Set pedreira = Nothing
        Next i
    End If
    ' Libera espa�o da memoria
    Set listaObjetos = Nothing
End Sub

' Carrega a combobox de serraria
Private Sub carregarSerrarias(cbSerrarias As MSForms.comboBox)
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim i As Integer
    
    ' Criando a lista
    Set listaObjetos = daoSerrada.listarSerrarias

    ' limpa a lista para carregamento
    cbSerrarias.Clear
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se n�o tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
       ' Loop atrav�s dos itens da cole��o
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set serraria = listaObjetos(i)
            ' Carregamento para lista
            cbSerrarias.AddItem serraria.nome
            ' Libera espa�o memoria
            Set serraria = Nothing
        Next i
    End If
    ' Libera espa�o da memoria
    Set listaObjetos = Nothing
End Sub

' Carrega a combobox de tipo material
Private Sub carregarTiposMateriais(cbTiposMateriais As MSForms.comboBox)
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim i As Integer
    
    ' Criando a lista
    Set listaObjetos = daoTipoMaterial.listarTiposMateriais

    ' limpa a lista para carregamento
    cbTiposMateriais.Clear
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se n�o tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop atrav�s dos itens da cole��o
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set tipoMaterial = listaObjetos(i)
            ' Carregamento para lista
            cbTiposMateriais.AddItem tipoMaterial.nome
            ' Libera espa�o memoria
            Set tipoMaterial = Nothing
        Next i
        ' Deixar um item j� selecionado
        Call selecaoItem("cbTipoMaterial", "EXTRA")
    End If
    ' Libera espa�o da memoria
    Set listaObjetos = Nothing
End Sub

' Carrega a combobox tem nota
Private Sub carregarTemNota(cbTemNota As MSForms.comboBox)
    ' limpa a lista para carregamento
    cbTemNota.Clear
    
    ' Deixar um item j� selecionado
    If Me.MultiPageCEBC.Value = 1 Then
        cbTemNota.AddItem ""
    End If
    
    ' Carregamento para lista
    cbTemNota.AddItem "SIM"
    cbTemNota.AddItem "N�O"
    
    ' Deixar um item j� selecionado
    If Me.MultiPageCEBC.Value = 2 Then
        Call selecaoItem("cbNotaC", "N�O")
    End If
End Sub

' Carrega a combobox de polideira
Private Sub carregarPolideiras(cbPolideiras As MSForms.comboBox)
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim i As Integer
    
    ' Criando a lista
    Set listaObjetos = daoPolideira.listarPolideiras

    ' limpa a lista para carregamento
    cbPolideiras.Clear
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se n�o tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop atrav�s dos itens da cole��o
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set polideira = listaObjetos(i)
            ' Carregamento para lista
            cbPolideiras.AddItem polideira.nome
            ' Libera espa�o memoria
            Set polideira = Nothing
        Next i
    End If
    ' Libera espa�o da memoria
    Set listaObjetos = Nothing
End Sub

' Carrega a combobox de tipo polimento
Private Sub carregarTiposPolimento(cbTiposPolimento As MSForms.comboBox)
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim i As Integer
    
    ' Criando a lista
    Set listaObjetos = daoTipoPolimento.listarTipoPolideiras

    ' limpa a lista para carregamento
    cbTiposPolimento.Clear
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se n�o tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop atrav�s dos itens da cole��o
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set tipoPolimento = listaObjetos(i)
            ' Carregamento para lista
            cbTiposPolimento.AddItem tipoPolimento.nome
            ' Libera espa�o memoria
            Set tipoPolimento = Nothing
        Next i
    End If
    ' Libera espa�o da memoria
    Set listaObjetos = Nothing
End Sub

' Carrega a combobox de estoque carregarCustoMedio
Private Sub carregarEstoque(cbTiposEstoque As MSForms.comboBox)
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim i As Integer
    
    ' Criando a lista
    Set listaObjetos = daoEstoqueM3.listarEstoqueM3

    ' limpa a lista para carregamento
    cbTiposEstoque.Clear
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se n�o tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop atrav�s dos itens da cole��o
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set estoque = listaObjetos(i)
            ' Carregamento para lista
            cbTiposEstoque.AddItem estoque.nome
            ' Libera espa�o memoria
            Set estoque = Nothing
        Next i
    End If
    ' Libera espa�o da memoria
    Set listaObjetos = Nothing
End Sub

' Carrega a combobox de custo medio
Private Sub carregarCustoMedio(cbCustoMedio As MSForms.comboBox)
    ' limpa a lista para carregamento
    cbCustoMedio.Clear
    
    ' Carregamento para lista
    cbCustoMedio.AddItem "SIM"
    cbCustoMedio.AddItem "N�O"
End Sub

' Carrega a combobox de status
Private Sub carregarStatus(cbStatus As MSForms.comboBox)
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim i As Integer
    
    ' Criando a lista
    Set listaObjetos = daoStatus.listarStatus

    ' limpa a lista para carregamento
    cbStatus.Clear
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se n�o tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop atrav�s dos itens da cole��o
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set statusObj = listaObjetos(i)
            ' Carregamento para lista
            cbStatus.AddItem statusObj.nome
            ' Libera espa�o memoria
            Set statusObj = Nothing
        Next i
    End If
    ' Libera espa�o da memoria
    Set listaObjetos = Nothing
End Sub

'-----------------------------------------------------------------CARREAGMENTO DAS LIST-----------------------------------
'                                                                 ---------------------
' Carrega a lista
Private Sub carregarList(ListBox As MSForms.ListBox, listaCollection As Collection)
   'Variaveis do metodo
    Dim objeto As objBloco
    Dim i As Integer
    Dim qtdChapas As Integer
    
    ' Limpar a ListBox
    ListBox.Clear
    
    ' NOME CABE�ALHO CHAPAS       | COD | DECRCI��O | QTD  | COMP  | ALT  | M�    | TIPO     | ESP | VALOR | TOTAL |
    ' NOME CABE�ALHO BLOCOS       | COD | DECRCI��O | COMP | ALT   | LARG | QTD   | VALOR M� | ADD | FRETE | TOTAL |
    ' Tamanho do cabe�alho left   | 7   | 193       | 444  | 496,5 | 549  | 601,5 | 654      | 745 | 820,5 | 896   |
    ' Tamanho do cabe�alho width  | 185 | 250       | 52   | 52    | 52   | 52    | 90       | 75  | 75    | 74,5  |
    ' Tamanho das colunas da list
    ListBox.ColumnWidths = "185;250;52;52;52;52;90;75;75;74;"
    
    ' Verifica se tem algum dado a pesquisa
    If listaCollection.Count = -1 Or listaCollection.Count = 0 Then ' Se n�o tiver dados
        If paginaAnterior <> 1 Then ' Ativa mensagem se a pagina anterior n�o for a do menu
            ' Mensagem de retorno
            errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        End If
    Else
        ' Loop atrav�s dos itens da cole��o
        For i = 1 To listaCollection.Count
            ' Seta o ojeto
            Set objeto = listaCollection(i)
            
            ' Adiciona uma linha
            ListBox.AddItem
            
            ' Adiciona os dados do bloco
            ListBox.list(ListBox.ListCount - 1, 0) = objeto.idSistema
            ListBox.list(ListBox.ListCount - 1, 1) = objeto.nomeMaterial
            ListBox.list(ListBox.ListCount - 1, 2) = _
                                    M_METODOS_GLOBAL.formatarComPontos(Format(objeto.compLiquidoBloco, "0.0000"))
            ListBox.list(ListBox.ListCount - 1, 3) = _
                                    M_METODOS_GLOBAL.formatarComPontos(Format(objeto.altLiquidoBloco, "0.0000"))
            ListBox.list(ListBox.ListCount - 1, 4) = _
                                    M_METODOS_GLOBAL.formatarComPontos(Format(objeto.largLiquidoBloco, "0.0000"))
            ListBox.list(ListBox.ListCount - 1, 5) = objeto.qtdChapas
            ListBox.list(ListBox.ListCount - 1, 6) = _
                                    M_METODOS_GLOBAL.formatarComPontos(Format(objeto.precoM3Bloco, "0.00"))
            ListBox.list(ListBox.ListCount - 1, 7) = _
                                    M_METODOS_GLOBAL.formatarComPontos(Format(objeto.valoresAdicionais, "0.00"))
            ListBox.list(ListBox.ListCount - 1, 8) = _
                                    M_METODOS_GLOBAL.formatarComPontos(Format(objeto.freteBloco, "0.00"))
            ListBox.list(ListBox.ListCount - 1, 9) = _
                                    M_METODOS_GLOBAL.formatarComPontos(Format(objeto.valorBloco, "0.00"))
            
            ' Total de blocos pesquisados
            lQtdBlocos.Caption = i
            ' Soma a qtd de chapas
            qtdChapas = qtdChapas + CInt(objeto.qtdChapas)
            ' Libera espa�o da memoria
            Set objeto = Nothing
        Next i
        ' Total de chapas
        lQtdChapas.Caption = qtdChapas
    End If
    ' Libera espa�o da memoria
    Set listaObjeto = Nothing
End Sub

' Carrega a lista ListTamanhosChapas tela edicao chapa
Private Sub carregarListTamanhosChapas(lista As MSForms.ListBox) ' Ir� receber id chapa para carregamento
    ' Limpar a ListBox
    lista.Clear
    
    ' NOME CABE�ALHO BLOCOS       | TIPO  | ESP   | COMP  | ALT | M�  | QTD |
    ' Tamanho do cabe�alho left   | 192,5 | 331,5 | 362,5 | 411 | 460 | 511 |
    ' Tamanho do cabe�alho width  | 138,5 | 30    | 48    | 48  | 50  | 30  |
    ' Tamanho das colunas da list
    lista.ColumnWidths = "140,5;30;48;48;50;35;"

    'Adiciona uma linha
    lista.AddItem
    
    'Adiciona os dados do bloco
    lista.list(lista.ListCount - 1, 0) = "COMERCIAL SATAND"
    lista.list(lista.ListCount - 1, 1) = "02"
    lista.list(lista.ListCount - 1, 2) = M_METODOS_GLOBAL.formatarComPontos(Format("3,0000", "0.0000"))
    lista.list(lista.ListCount - 1, 3) = M_METODOS_GLOBAL.formatarComPontos(Format("2,0000", "0.0000"))
    lista.list(lista.ListCount - 1, 4) = M_METODOS_GLOBAL.formatarComPontos(Format("146,0000", "0.0000"))
    lista.list(lista.ListCount - 1, 5) = "71"
End Sub
