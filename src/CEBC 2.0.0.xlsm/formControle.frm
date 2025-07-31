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
Dim errorStyle As clsErrorStyle

' Variaveis para manipulação
Public paginaAnterior As Integer
Dim status() As String
Dim listaObjeto As Collection
Dim frm As formTamanhos

' Variaveis de objetos
Dim bloco As objBloco
Public chapa As objChapa
Dim pedreira As objPedreira
Dim polideira As objPolideira
Dim serraria As objSerraria
Public tamanho As objTamanho
Dim tipoMaterial As objTipoMaterial
Dim tipoPolimento As objTipoPolimento
Dim statusObj As objStatus
Dim estoque As objEstoque
Dim estoqueChapa As objEstoqueChapa
Dim motorista As objMotorista
Dim destino As objDestino
Dim despache As objDespache

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
    
    ' Carrega tradução do sistema
    Call M_TRADUCAO.carregarTraducaoErros
    
    ' Seta pagina
    paginaAnterior = 0
    
    ' Resevando espaço em memoria para manipulação das variaveis
    ReDim botoesMenu(1 To Me.Controls.Count)
    ReDim botoesImg(1 To Me.Controls.Count)
    ReDim botoesText(1 To Me.Controls.Count)
    ReDim frameEfeito(1 To Me.Controls.Count)
    ReDim status(1 To 6)
    Set errorStyle = New clsErrorStyle
    
    ' Atribuições da variaveis
    status(1) = "PEDREIRA"
    status(2) = "SERRARIA"
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
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 4
    ' Seta pagina anterior
    paginaAnterior = 0
    ' Seta o foco
    txtMaterialChapaPesquisa.SetFocus
    
    ' Carregar os comboBox da tela
    Call carregarPolideiras(Me.cbPolideiraChapaPesquisa)
    Call carregarTiposPolimento(Me.cbTipoPolimentoPesquisa)
End Sub
' Efeito para clique nas label btnLMenuDespachar do menu
Private Sub btnLMenuDespachar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 8
    ' Seta pagina anterior
    paginaAnterior = 0
    ' Seta o foco
    cbMotorista.SetFocus
    ' Seta da atual
    txtDataDespacho.Value = Date
    
    ' Carregar os comboBox da tela
    Call carregarMotoristas(Me.cbMotorista)
    Call carregarDestinos(Me.cbDestino)
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
' Efeito para quando sair do foco de txtNomeArquivoEstoqueBlocos de texto tela estoque m³
Private Sub fTiraEfeitoBotoesExportarBlocosM3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
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
    ' Chama serviço para pesquisa
    Call pesquisarBlocosFilter
End Sub
' Botão btnLTxtLimparFiltrosBlocos tela estoque m³
Private Sub btnLTxtLimparFiltrosBlocos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    Call limparCamposPesquisaEstoqueM3
End Sub
' Botão btnLImgExportarEstoqueM3 tela estoque m³
Private Sub btnLImgExportarEstoqueM3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Variaveis do metodo
    Dim idsParaPesquisa As Collection
    Dim id As String
    Dim i As Integer
    
    ' Verifica se tem dados na lista
    If Me.ListEstoqueM3.ListCount > 0 Then
        ' Reatribui espaço na memoria para variavel
        Set idsParaPesquisa = ObjectFactory.factoryLista(idsParaPesquisa)
    Else
        ' Mensagem de erro
        errorStyle.Informativo LIST_SEM_DADOS_MENSAGEM, LIST_SEM_DADOS_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    
    ' Verifica se foi digitado nome para o arquivo
    If txtNomeArquivoEstoqueBlocos.Value = "" Or txtNomeArquivoEstoqueBlocos.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtNomeArquivoEstoqueBlocos, ARQUIVO_SEM_NOME_MENSAGEM, ARQUIVO_SEM_NOME_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
     errorStyle.sairErrorStyleTextBox txtNomeArquivoEstoqueBlocos
    
    ' Captura ids da lista
    For i = 0 To Me.ListEstoqueM3.ListCount - 1
        idsParaPesquisa.Add Me.ListEstoqueM3.list(i, 0)
    Next i
    
    ' Pesquisa os ids
    Set listaObjeto = daoBloco.pesquisarPorIdsVariados(idsParaPesquisa)
    
    ' Exporta em pdf
    Call ExportarArquivos.exportarEstoqueBloco(listaObjeto, txtNomeArquivoEstoqueBlocos.Value)
    
    ' Libera espeço na memoria
    Set idsParaPesquisa = Nothing
    Set listaObjeto = Nothing
End Sub
' Botão btnLTxtNovoBloco tela estoque m³
Private Sub btnLTxtNovoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 2
    ' Seta número de pagina para poder voltar
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
    ' Libera espaço em memoria
    Set listaObjeto = Nothing
End Sub
' Botão btnLTxtEditarBloco tela estoque m³
Private Sub btnLTxtEditarBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Verifica se tem algum item selecionado
    If Me.ListEstoqueM3.ListIndex = -1 Then
        ' Mensagem usuário
        errorStyle.Informativo SELECIONE_TEM_MENSAGEM, SELECIONE_TEM_TITULO
        Exit Sub
    End If
    
    ' Muda abra da multPage para tela editar bloco
    Me.MultiPageCEBC.Value = 3
    ' Seta paniga anterior para futuras condições
    paginaAnterior = 1
    
    ' Chama serviço para pesquisa do bloco
    Set bloco = daoBloco.pesquisarPorId(Me.ListEstoqueM3.list(Me.ListEstoqueM3.ListIndex, 0), True) ' Envia o id do bloco e true para fechar conexão ao final da pesquisar
    
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
' Botão btnLTxtADDEstoque tela estoque m³
Private Sub btnLTxtADDEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ' Variaveis do metodo
    Dim temCadastro As Boolean
    Dim chapaCadastro As objChapa
    Dim tamanhos As Collection
    Dim idChapa As String
    Dim descricaoChapa As String
    Dim valorTotalSerrada As String
    
    ' Verifica se tem algum item selecionado
    If Me.ListEstoqueM3.ListIndex = -1 Then
        ' Mensagem usuário
        errorStyle.Informativo SELECIONE_TEM_MENSAGEM, SELECIONE_TEM_TITULO
        Exit Sub
    End If
    
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
    
    ' Seta número de pagina para poder voltar
    paginaAnterior = 1
    
    ' Chama serviço para pesquisa do bloco
    Set bloco = daoBloco.pesquisarPorId(Me.ListEstoqueM3.list(Me.ListEstoqueM3.ListIndex, 0), True) ' Envia o id do bloco e true para fechar conexão ao final da pesquisar
    
    ' Pesquisa se já existe cadastro de chapas com id do bloco
    temCadastro = daoChapa.pesquisarPorIdPedreira(bloco.numeroBlocoPedreira)
    
    If temCadastro = True Then
        ' Direciona para tela de estoque de chapas para que possa ser escolhido qual chapa será adicionada no estoque
        Me.MultiPageCEBC.Value = 4
        ' Seta numero do bloco para pesquisa
        txtIdBlocoChapaPesquisa.Value = bloco.numeroBlocoPedreira
        ' Chama Serviço
        Call pesquisarChapasFilter

    Else
        ' Direciona para tela lançamento e edição de chapa
        Me.MultiPageCEBC.Value = 6
        ' Carrega combox da tela lançamento e edição de chapa
        Call carregarTiposMateriais(Me.cbTipoMaterialChapaC)
        Call carregarEstoqueChapas(Me.cbEstoqueChapaC)
        Call carregarPolideiras(Me.cbPolideiraChapa)
        
        ' limpa a lista para carregamento com tipo de polimento só com 'bruto'
        cbTipoPolimentoChapa.Clear
        cbTipoPolimentoChapa.AddItem "BRUTO"
        
        ' Cria chapa e direciona para tela de lançamento e edição de chapa para colocar demais informações
        Set chapaCadastro = ObjectFactory.factoryChapa(chapaCadastro)
        Set tipoPolimento = daoTipoPolimento.pesquisarPorNome("BRUTO")
        Set tamanhos = ObjectFactory.factoryLista(tamanhos)
        
        ' Formatar id, descrição da chapa e valor total serrada
        idChapa = M_METODOS_GLOBAL.formatarIdChapa(bloco.idSistema, "BT")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(bloco.nomeMaterial, "BRUTO")
        valorTotalSerrada = M_METODOS_GLOBAL.calcularValor(bloco.qtdM2Serrada, bloco.valorMetroSerrada)
        
        chapaCadastro.carregarChapa idChapa, descricaoChapa, valorTotalSerrada, bloco.numeroBlocoPedreira, _
                        tipoPolimento, bloco, tamanhos
                        
        ' Carrega os dados na tela lançamento e edição de chapa
        Call carregarDadosChapaTelaEdicaoChapa(chapaCadastro, bloco)
        
        ' Libera espaço em memoria
        Set chapaCadastro = Nothing
        Set tipoPolimento = Nothing
        Set tamanhos = Nothing
    End If
    
    ' Libera espaço em memoria
    Set bloco = Nothing
End Sub
' Botão btnLTxtExcluirBloco tela estoque m³
Private Sub btnLTxtExcluirBloco_Click()
    
    Dim resposta As VbMsgBoxResult ' Variavel para confirmação do bloco
    
    ' Verifica se tem algum item selecionado
    If Me.ListEstoqueM3.ListIndex = -1 Then
        ' Mensagem usuário
        errorStyle.Informativo SELECIONE_TEM_MENSAGEM, SELECIONE_TEM_TITULO
        Exit Sub
    End If
    
    ' Mensagem de confirmação
    resposta = MsgBox(EXCLUIR_BLOCO_MENSAGEM, vbQuestion + vbYesNo, EXCLUIR_BLOCO_TITULO)
    
    If resposta = vbYes Then
        ' Chama serviço para excluir o bloco
        daoBloco.excluir (Me.ListEstoqueM3.list(Me.ListEstoqueM3.ListIndex, 0))
    End If
    
    ' Chama serviço para pesquisa
    Call pesquisarBlocosFilter
    
    ' Mensagem usuário
    errorStyle.Informativo SUCESSO_EXCLUIR_BLOCO_MENSAGEM, SUCESSO_EXCLUIR_BLOCO_TITULO
End Sub

'-----------------------------------------------------------------TELA CADASTRO DE BLOCOS-----------------------------------
'                                                                 -----------------------
' txtIdBloco tela cadastro de bloco
Private Sub txtIdBloco_Change()
    ' Coloca tudo em caixa alta
    txtIdBloco.Value = UCase(txtIdBloco.Value)
    
    ' Cria o código para o sistema
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
    
    ' Cria o código para o sistema
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
    txtValorBloco.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValor( _
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
    txtValorBloco.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValor(txtValorM3.Value, _
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
    txtValorBloco.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValor( _
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
    txtValorBloco.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValor( _
            txtValorM3.Value, txtTotalM3.Value), "0.00"))
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
    Dim blocoPesquisa As objBloco
    Dim resposta As VbMsgBoxResult ' Variavel para confirmação na hora de cadastrar
    Dim nomeStatus As String
    Dim nomeMaterial As String
    Dim valorTotalBloco As String
    Dim cadastro As Boolean
    
    ' Patrão true
    cadastro = True
    
    ' Captura do status
    If obPedreiraCB.Value = True Then
        nomeStatus = status(1)
    Else
        nomeStatus = status(2)
    End If
    
    ' Validações
    nomeMaterial = "BLOCO " & txtNomeBloco.Value
    
    ' Verifica o Status
    If obSerrariaCB.Value = True Then
        If cbSerrariaCB.Value = "" Or cbSerrariaCB.Value = " " Then
            ' Deixa visivel o erro com mensagens
            errorStyle.EntrarErrorStyleComboBox cbSerrariaCB, STATUS_SERRARIA_MENSAGEM, STATUS_SERRARIA_TITULO
            ' Para o fluxo do sistema para a correção
            Exit Sub
        End If
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleComboBox cbSerrariaCB
    
    ' Verifica o Pedreira
    If cbPedreira.Value = "" Or cbPedreira.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleComboBox cbPedreira, NOME_PEDREIRA_MENSAGEM, NOME_PEDREIRA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleComboBox cbPedreira
    
    ' Verifica o Número do bloco na pedreira
    If txtIdBloco.Value = "" Or txtIdBloco.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtIdBloco, NUMERO_BLOCO_PEDREIRA_MENSAGEM, NUMERO_BLOCO_PEDREIRA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtIdBloco
    
    ' Verifica nome do bloco
    If txtNomeBloco.Value = "" Or txtNomeBloco.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtNomeBloco, NOME_BLOCO_MENSAGEM, NOME_BLOCO_PEDREIRA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtNomeBloco
        
    ' Verifica se um cadastrado ou edição
    Set blocoPesquisa = daoBloco.pesquisarPorId(txtIdBlocoSistema.Value)
    If blocoPesquisa.idSistema = txtIdBlocoSistema.Value Then
        ' Mensagem de id já cadastrado no sistema
        errorStyle.Informativo BLOCO_JA_CADASTRADO_MENSAGEM, BLOCO_JA_CADASTRADO_TITULO
        Exit Sub
    Else
        ' Mensagem de confirmação
        resposta = MsgBox(CONFIRMACAO_CADASTRO_MENSAGEM, vbQuestion + vbYesNo, CONFIRMACAO_CADASTRO_TITULO)
    End If
    
    ' Verifica a confirmação do usário para poder cadastrar
    If resposta = vbYes Then
        ' Criação dos objetos
        Set pedreira = daoPedreira.pesquisarPorNome(cbPedreira.Value)
        Set serraria = daoSerrada.pesquisarPorNome(cbSerrariaCB.Value)
        Set tipoMaterial = daoTipoMaterial.pesquisarPorNome(cbTipoMaterial.Value)
        Set statusObj = daoStatus.pesquisarPorNome(nomeStatus)
        Set estoque = daoEstoqueM3.pesquisarPorNome("CASA DO GRANITO")
        Set bloco = ObjectFactory.factoryBloco(bloco)
        Set blocoPesquisa = ObjectFactory.factoryBloco(blocoPesquisa)
        
        ' Calcula valor total bloco
        valorTotalBloco = M_METODOS_GLOBAL.formatarComPontos(Format(custoBloco( _
                    txtValorBloco.Value, txtValorFreteBloco.Value, "0", "0", txtAdicionais.Value), "0.00"))
        
        ' Criação do objeto
        bloco.carregarBlocoCadastro txtDataCadastro.Value, txtIdBlocoSistema.Value, pedreira, serraria, txtIdBloco.Value, _
                                    nomeMaterial, tipoMaterial, cbNotaC.Value, statusObj, txtObsBlocoCB.Value, _
                                    txtCompBrutoBloco.Value, txtAlturaBlocoBruto.Value, txtLarguraBlocoBruto.Value, _
                                    txtComprimentoBloco.Value, txtAlturaBloco.Value, txtLarguraBloco.Value, estoque, _
                                    txtAdicionais.Value, txtValorFreteBloco.Value, txtValorM3.Value, txtTotalM3.Value, _
                                    txtValorBloco.Value, valorTotalBloco, "NÃO"
        
        ' Chama serviço para cadastrar do bloco
        Call daoBloco.cadastrarEEditar(bloco)
        
        ' Verifica se foi um cadastro ou edição para personalisar as mensagens
        If cadastro = True Then
            ' Verifica se bloco foi cadastrado
            Set blocoPesquisa = daoBloco.pesquisarPorId(bloco.idSistema)
            If blocoPesquisa.idSistema = txtIdBlocoSistema.Value Then
                ' Limpa os campos
                Call limparCamposCadastroBlocos
                ' Recarregar a lista com blocos cadastrados hoje
                ' Pesquisa blocos cadastrado no dia atual
                Set listaObjeto = daoBloco.listarBlocosFilter(Date, Date, "", "", "", "", "", "", "", "", "", "", "")
                
                ' Chama metodo para carregar lista e blocos cadastros do dia atual
                Call carregarList(Me.listCadastradosHoje, listaObjeto)
                ' Mensagem de cadastro realizado com sucesso.
                errorStyle.Informativo CADASTRO_CONFIRMADO_MENSAGEM, CADASTRO_CONFIRMADO_TITULO
            Else
                ' Mensagem de erro desconhecido
                errorStyle.Informativo ERRO_DESCONHECIDO_MENSAGEM, ERRO_DESCONHECIDO_TITULO
            End If
        End If

        ' Libera espaço da memoria
        Set pedreira = Nothing
        Set serraria = Nothing
        Set tipoMaterial = Nothing
        Set statusObj = Nothing
        Set estoque = Nothing
        Set bloco = Nothing
        Set blocoPesquisa = Nothing
    Else
        ' Coloque o código a ser executado se o usuário clicar em "Não" aqui.
        errorStyle.Informativo ACAO_CANCELADA_MENSAGEM, ACAO_CANCELADA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa o cursor no cbPedreira para proximo cadastro
    cbPedreira.SetFocus
End Sub
' Botão btnLTxtVoltarCadastroBloco tela cadastrar bloco
Private Sub btnLTxtVoltarCadastroBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage - tela estoque m³
    Me.MultiPageCEBC.Value = paginaAnterior
    ' Seta o foco
    txtMaterialBlocoPesquisa.SetFocus
    ' Chama serviço para pesquisa
    Call pesquisarBlocosFilter
End Sub
' Botão btnLTextLimparCadastroBloco tela cadastrar bloco
Private Sub btnLTxtLimparCadastroBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
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
    ' Deixa só a digitação de numero
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
    ' Coloca as barras para formatação
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
    
    ' Seta m² polimento para calculo do custo
    txtTotalM2PolimentoBlocoEditar.Value = txtQtdM2PolimentoEditar.Value
End Sub

' txtTotalChapaBlocoEditar tela editar bloco
Private Sub txtTotalChapaBlocoEditar_Change()
    ' Variaveis do metodo
    Dim textoDigitado As String
    Dim textoFormatado As String
    Dim i As Integer

    'Recebi o texto digitadado pelo usúario
    textoDigitado = txtTotalChapaBlocoEditar.Value
 
    'Remove todos os caracteres não numéricos
    For i = 1 To Len(textoDigitado)
        If IsNumeric(Mid(textoDigitado, i, 1)) Then
            textoFormatado = textoFormatado & Mid(textoDigitado, i, 1)
        End If
        
        'Remove o zero na esquerda do texto
        If Len(textoFormatado) = 2 Then
            If Left(textoFormatado, 1) = 0 Then
                textoFormatado = Mid(textoFormatado, 2, 1)
            End If
        End If
    Next i
    
    If textoFormatado = "" Or textoFormatado = " " Then
        textoFormatado = "0"
    End If
    
    txtTotalChapaBlocoEditar.Value = textoFormatado
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
    ' Descrição e dimensões finais
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
    ' Dimensões bloco e médias chapas
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
        Call desabilitaCamposBlocoEditar
    Else
        cbAbrirBlocoEditar.Visible = False
        lBlocoFinalizado.Visible = False
        Call habilitaCamposBlocoEditar
    End If
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
    ' Variaveis do medoto
    Dim blocoPesquisa As objBloco
    
    ' Verifica se esta habilitado para edição bloco finalizado
    Set blocoPesquisa = daoBloco.pesquisarPorId(txtIdBlocoEditar.Value, True)
    
    If blocoPesquisa.status.nome = "FECHADO" Then
        If cbAbrirBlocoEditar.Value = False Then
            ' Mensagem de habilite para edição
            errorStyle.Informativo HABILITE_EDICAO_MENSAGEM, HABILITE_EDICAO_TITULO
            Exit Sub
        End If
    End If
    ' Desabilita edição
    cbAbrirBlocoEditar.Value = False
    
    ' Verifica o Número do bloco na pedreira
    If txtNBlocoPedreiraEditar.Value = "" Or txtNBlocoPedreiraEditar.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtNBlocoPedreiraEditar, NUMERO_BLOCO_PEDREIRA_MENSAGEM, NUMERO_BLOCO_PEDREIRA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtNBlocoPedreiraEditar
    
    ' Verifica nome do bloco
    If txtMaterialEditar.Value = "" Or txtMaterialEditar.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtMaterialEditar, NOME_BLOCO_MENSAGEM, NOME_BLOCO_PEDREIRA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtMaterialEditar
    
    ' Criação dos objetos
    Set pedreira = daoPedreira.pesquisarPorNome(cbPedreiraEditar.Value)
    Set serraria = daoSerrada.pesquisarPorNome(cbSerrariaEditar.Value)
    Set polideira = daoPolideira.pesquisarPorNome(cbPolideiraEditar)
    Set tipoMaterial = daoTipoMaterial.pesquisarPorNome(cbTipoMaterialEditar.Value)
    Set statusObj = daoStatus.pesquisarPorNome(cbStatusBlocoEditar)
    Set estoque = daoEstoqueM3.pesquisarPorNome(cbEstoqueEditar.Value)
    Set bloco = ObjectFactory.factoryBloco(bloco)
    
    ' Criação do objeto
    bloco.carregarBlocoEdicao txtIdBlocoEditar.Value, txtMaterialEditar.Value, txtObsEditar.Value, txtNBlocoPedreiraEditar.Value, estoque, _
                    txtDataCadastroEditar.Value, txtQtdM3blocoEditar.Value, txtQtdM2SerradaEditar.Value, txtQtdM2PolimentoEditar.Value, txtTotalChapaBlocoEditar.Value, cbNotaBlocoEditar.Value, _
                    cbCustoMedioEditar.Value, txtCompBrutaBlocoEditar.Value, txtAltBrutaBlocoEditar.Value, txtLArgBrutaBlocoEditar.Value, txtCompLiquidoBlocoEditar.Value, _
                    txtAltLiquidoBlocoEditar.Value, txtLArgLiquidoBlocoEditar.Value, txtCompBrutaBrutoChapaEditar.Value, txtAltBrutaBrutoChapaEditar.Value, _
                    txtCompBrutaliquidoChapaEditar.Value, txtAltBrutaLiquidoChapaEditar.Value, txtCompPolidaBrutoChapaEditar.Value, txtAltPolidaBrutoChapaEditar.Value, _
                    txtCompPolidaLiquidoChapaEditar.Value, txtAltPolidaLiquidaChapaEditar.Value, txtValoBlocoEditar.Value, txtPrecoBlocoEditar.Value, _
                    txtFreteBlocoEditar.Value, txtValorSerradaEditar.Value, txtValorPolimentoEditar.Value, txtValorADDImpostosEditar.Value, _
                    txtTotalSerradaEditar.Value, txtTotalPolimentoEditar.Value, txtCustoMaterialBlocoEditar.Value, txtTotalBlocoEditar.Value, _
                    statusObj, tipoMaterial, pedreira, serraria, polideira
    
    ' Chama serviço para cadastrar do bloco
    Call daoBloco.cadastrarEEditar(bloco)
    
    ' Chama serviço para pesquisa do bloco
    Set bloco = daoBloco.pesquisarPorId(bloco.idSistema, True) ' Envia o id do bloco
    
    ' Recarrega os dados na tela editar bloco
    Call carregarDadosBlocoTelaEdicaoBloco(bloco)
    
    ' Libera espaço da memorio
    Set pedreira = Nothing
    Set serraria = Nothing
    Set polideira = Nothing
    Set tipoMaterial = Nothing
    Set statusObj = Nothing
    Set estoque = Nothing
    Set bloco = Nothing
    
    ' Mensagem de edição realizada com sucesso.
    errorStyle.Informativo HABILITE_EDICAO_MENSAGEM, HABILITE_EDICAO_TITULO
End Sub
' Botão btnLTxtVoltarEdicaoBloco tela editar bloco
Private Sub btnLTxtVoltarEdicaoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 1
    ' Desabilita
    cbAbrirBlocoEditar.Visible = False
    lBlocoFinalizado.Visible = False
    ' Chama serviço para pesquisa
    Call pesquisarBlocosFilter
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

' Efeito para quando sair do foco de txtNomeArquivoEstoqueChapas de texto tela estoque m²
Private Sub fTiraEfeitoBotoesExportarChapasM2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
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
    Call pesquisarChapasFilter
End Sub

' Botão btnLTxtLimparFiltrosChapas tela estoque m²
Private Sub btnLTxtLimparFiltrosChapas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    Call limparCamposPesquisaEstoqueM2
End Sub

' Botão btnLImgExportarEstoqueM2 tela estoque m²
Private Sub btnLImgExportarEstoqueM2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Variaveis do metodo
    Dim idsParaPesquisa As Collection
    Dim id As String
    Dim i As Integer
    
    ' Verifica se tem dados na lista
    If Me.ListEstoqueChapas.ListCount > 0 Then
        ' Reatribui espaço na memoria para variavel
        Set idsParaPesquisa = ObjectFactory.factoryLista(idsParaPesquisa)
    Else
        ' Mensagem de erro
        errorStyle.Informativo LIST_SEM_DADOS_MENSAGEM, LIST_SEM_DADOS_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    
    ' Verifica se foi digitado nome para o arquivo
    If txtNomeArquivoEstoqueChapas.Value = "" Or txtNomeArquivoEstoqueChapas.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtNomeArquivoEstoqueChapas, ARQUIVO_SEM_NOME_MENSAGEM, ARQUIVO_SEM_NOME_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
     errorStyle.sairErrorStyleTextBox txtNomeArquivoEstoqueChapas
    
    ' Captura ids da lista
    For i = 0 To Me.ListEstoqueChapas.ListCount - 1
        idsParaPesquisa.Add Me.ListEstoqueChapas.list(i, 0)
    Next i
    
    ' Pesquisa os ids
    Set listaObjeto = daoChapa.pesquisarPorListaIdsChapas(idsParaPesquisa)
    
    ' Exporta em pdf
    Call ExportarArquivos.exportarEstoqueChapa(listaObjeto, txtNomeArquivoEstoqueChapas.Value)
    
    ' Libera espeço na memoria
    Set idsParaPesquisa = Nothing
    Set listaObjeto = Nothing
End Sub

'Botão btnLTxtNovoAvulso tela estoque m²
Private Sub btnLTxtNovoAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Variaveis do metodo
    Dim listaAvulsosCadastradosHoje As Collection
    Dim idsChapaAvulso As Collection
    Dim blocoLista As objBloco
    Dim primeiroNome As String
    Dim i As Integer
    
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 5
    ' Seta paniga anterior para futuras condições
    paginaAnterior = 4
    ' Coloca data atual na txtDataCadastroChapaAvulsa na tela cadastro chapa avulso
    txtDataCadastroChapaAvulsa.Value = Date
    ' Seta o foco
    txtIdBlocoAvulso.SetFocus
    
    ' Chama metodo para carregar comboBox
    Call carregarTiposMateriais(Me.cbTipoMaterialL)
    Call carregarTiposPolimento(Me.cbTipoPolimentoL)
    Call carregarTemNota(Me.cbTemNotaAvulso)
    
    ' Pesquisa blocos cadastrado no dia atual
    Set listaObjeto = daoBloco.listarBlocosFilter(Date, Date, "", "", "", "", "", "", "", "", "", "", "")
    Set idsChapaAvulso = ObjectFactory.factoryLista(idsChapaAvulso)
    
    ' Seleciona só os avulsos
    For i = 1 To listaObjeto.Count
        ' Seta bloco da lista
        Set blocoLista = listaObjeto.Item(i)
        ' Captura o primeiro nome da descrição
        primeiroNome = Mid(blocoLista.nomeMaterial, 1, 5)
        ' Confere se é um avulso ou importado
        If primeiroNome <> "BLOCO" Then
            ' Captura as chapa avulso/importado para pesquisa
            idsChapaAvulso.Add blocoLista.numeroBlocoPedreira
        End If
    Next i
    
    If idsChapaAvulso.Count = 0 Or idsChapaAvulso.Count = -1 Then
        ' Apanas cria o objeto
        Set listaAvulsosCadastradosHoje = ObjectFactory.factoryLista(listaAvulsosCadastradosHoje)
    Else
        ' Pesquisa pelas chapas avulsas e importadas
        Set listaAvulsosCadastradosHoje = daoChapa.pesquisarPorListaIdsPedreira(idsChapaAvulso)
    End If
    
    ' Chama metodo para carregar lista e blocos cadastros do dia atual
    Call carregarList(ListMateriais, listaAvulsosCadastradosHoje)
    
    ' Libera espaço na memoria
    Set listaObjeto = Nothing
    Set listaAvulsosCadastradosHoje = Nothing
    Set idsChapaAvulso = Nothing
End Sub

' Botão btnLTxtNovoChapa tela estoque m²
Private Sub btnLTxtNovoChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Variaveis do metodo
    Dim listaPolimentosJaCadastras As Collection
    Dim listaChapasPesquisa As Collection
    Dim tamanhos As Collection
    Dim chapaPesquisa As objChapa
    Dim chapaCadastro As objChapa
    Dim idChapa As String
    Dim descricaoChapa As String
    Dim valorTotalSerrada As String
    Dim i As Integer
    
    ' Verifica se tem algum item selecionado
    If Me.ListEstoqueChapas.ListCount = 0 Then
        ' Mensagem usuário
        errorStyle.Informativo PESQUISA_SEM_DADOS_MENSAGEM, PESQUISA_SEM_DADOS_TITULO
        Exit Sub
    End If
    
    ' Muda abra da multPage para tela editar bloco
    Me.MultiPageCEBC.Value = 6
    ' Seta paniga anterior para futuras condições
    paginaAnterior = 4
    
    ' Analisar quais tipos de polimentos vão ser carregador
    Set chapaPesquisa = daoChapa.pesquisarPorId(Me.ListEstoqueChapas.list(0, 0))
    Set listaChapasPesquisa = daoChapa.pesquisarPorFKBloco(chapaPesquisa.bloco.idSistema)
    Set listaPolimentosJaCadastras = ObjectFactory.factoryLista(listaPolimentosJaCadastras)
    
    ' Carrega combox da tela lançamento e edição de chapa
    Call carregarPolideiras(Me.cbPolideiraChapa)
    Call carregarTiposMateriais(Me.cbTipoMaterialChapaC)
    Call carregarEstoqueChapas(Me.cbEstoqueChapaC)
    
    ' Loop através dos itens da coleção para obter os polimentos já cadastrados
    For i = 1 To listaChapasPesquisa.Count
        ' Seta o ojeto
        Set chapaPesquisa = listaChapasPesquisa(i)
        ' Seta os polimentos já cadastrados
        listaPolimentosJaCadastras.Add chapaPesquisa.tipoPolimento.nome
    Next i
    
    ' Carrega só os tipos deferentes
    Call carregarTiposPolimentoAlgum(cbTipoPolimentoChapa, listaPolimentosJaCadastras)
    
    ' Cria chapa e direciona para tela de lançamento e edição de chapa para colocar demais informações
    ' Chama serviço para pesquisa do bloco
    Set bloco = daoBloco.pesquisarPorId(chapaPesquisa.bloco.idSistema, True) ' Envia o id do bloco e true para fechar conexão ao final da pesquisar
    Set chapaCadastro = ObjectFactory.factoryChapa(chapaCadastro)
    Set tipoPolimento = ObjectFactory.factoryTipoPolimento(tipoPolimento)
    Set tamanhos = ObjectFactory.factoryLista(tamanhos)
    
    ' Formatar id, descrição da chapa e valor total serrada
    idChapa = bloco.numeroBlocoPedreira
    descricaoChapa = Mid(bloco.nomeMaterial, 7, Len(bloco.nomeMaterial))
    valorTotalSerrada = "0,00"
    
    ' Cria o objeto
    chapaCadastro.carregarChapa idChapa, descricaoChapa, valorTotalSerrada, bloco.numeroBlocoPedreira, _
                        tipoPolimento, bloco, tamanhos
                        
    ' Carrega os dados na tela lançamento e edição de chapa
    Call carregarDadosChapaTelaEdicaoChapa(chapaCadastro, bloco)
    
    ' Libera espaço em memoria
    Set bloco = Nothing
    Set chapaPesquisa = Nothing
    Set chapaCadastro = Nothing
    Set polideira = Nothing
    Set tipoPolimento = Nothing
    Set estoqueChapa = Nothing
    Set tamanhos = Nothing
End Sub

' Botão btnLTxtEditarChapa tela estoque m²
Private Sub btnLTxtEditarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ' Variaveis do metodos
    Dim chapaPesquisa As objChapa
    
    ' Verifica se tem algum item selecionado
    If Me.ListEstoqueChapas.ListIndex = -1 Then
        ' Mensagem usuário
        errorStyle.Informativo ESCOLHA_CHAPA_MENSAGEM, ESCOLHA_CHAPA_TITULO
        Exit Sub
    End If
    
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 6
    ' Seta número de pagina para poder voltar
    paginaAnterior = 4
    
    ' Carrega os ComboBox da tela
    Call carregarPolideiras(cbPolideiraChapa)
    Call carregarTiposPolimento(cbTipoPolimentoChapa)
    Call carregarTiposMateriais(cbTipoMaterialChapaC)
    Call carregarEstoque(cbEstoqueChapaC)
    
    ' Pesquisa pela chapa e o bloco do mesmo
    Set chapaPesquisa = daoChapa.pesquisarPorId(Me.ListEstoqueChapas.list(Me.ListEstoqueChapas.ListIndex, 0))
    Set bloco = daoBloco.pesquisarPorId(chapaPesquisa.bloco.idSistema, True) ' Envia o id do bloco e true para fechar conexão ao final da pesquisar
    
    ' Carrega os dados na tela lançamento e edição de chapa
    Call carregarDadosChapaTelaEdicaoChapa(chapaPesquisa, bloco)
    
    ' Libera espaço em memoria
    Set bloco = Nothing
    Set chapaPesquisa = Nothing
End Sub
    
' Botão btnLTxtTrocaEstoque tela estoque m²
Private Sub btnLTxtTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ' Verifica se tem algum item selecionado
    If Me.ListEstoqueChapas.ListIndex = -1 Then
        ' Mensagem usuário
        errorStyle.Informativo ESCOLHA_CHAPA_MENSAGEM, ESCOLHA_CHAPA_TITULO
        Exit Sub
    End If
    
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 7
    ' Seta pagina anterior
    paginaAnterior = 4
    ' Seta o foco
    cbTipoPolimentoTroca.SetFocus
    
    ' Pesquisa pela chapa e o bloco do mesmo
    Set chapa = daoChapa.pesquisarPorId(Me.ListEstoqueChapas.list(Me.ListEstoqueChapas.ListIndex, 0))
    
    ' Abre tela com tamanhos da chapa
    formTamanhos.Show
End Sub

' Botão btnLTxtTamanhos tela estoque m²
Private Sub btnLTxtTamanhos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ' Verifica se tem algum item selecionado
    If Me.ListEstoqueChapas.ListIndex = -1 Then
        ' Mensagem usuário
        errorStyle.Informativo ESCOLHA_CHAPA_MENSAGEM, ESCOLHA_CHAPA_TITULO
        Exit Sub
    End If
    
    ' Seta pagina anterior
    paginaAnterior = 4
    
    ' Pesquisa pela chapa
    Set chapa = daoChapa.pesquisarPorId(Me.ListEstoqueChapas.list(Me.ListEstoqueChapas.ListIndex, 0))
    
    ' Abre tela com tamanhos da chapa
    formTamanhos.Show
End Sub

'-----------------------------------------------------------------TELA CADASTRO AVULSO-----------------------------------
'                                                                 --------------------
' txtIdBlocoAvulso tela cadastro avulso
Private Sub txtIdBlocoAvulso_Change()
    ' Coloca tudo em caixa alta
    txtIdBlocoAvulso.Value = UCase(txtIdBlocoAvulso.Value)
    
    ' Cria o código para o sistema
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
    
    ' Cria o código para o sistema
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
        
    'Seta o custo do material m²
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
        
    'Seta o custo do material m²
    txtCustoSimplesM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.custoMaterialM2( _
            txtTotalBlocoAvulso.Value, txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", _
            txtTotalM2Avulso.Value), "0.00"))
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
    txtTotalM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM2( _
        txtComprimentoChapaAvulsa.Value, txtAlturaChapaAvulsa.Value, txtQuantidadeChapasAvulsas.Value), "0.0000"))
    'Seta o custo do material m²
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
    
    'Seta o custo do material m²
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
    
    'Seta o custo do material m²
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
    txtTotalBlocoAvulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValor( _
            txtTotalM2Avulso.Value, txtValorMetroAvulso.Value), "0.00"))
            
    'Se m² for diferente de 0 calcula o custo do material
    'Seta o custo do material m²
    txtCustoSimplesM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.custoMaterialM2(txtTotalBlocoAvulso.Value, _
            txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", txtTotalM2Avulso.Value), "0.00"))
End Sub
' txtTotalM2Avulso tela cadastro avulso
Private Sub txtTotalM2Avulso_Change()
    'Se m² for diferente de 0 calcula o custo do material
    'Seta o custo do material m²
    txtCustoSimplesM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.custoMaterialM2( _
            txtTotalBlocoAvulso.Value, txtValorFreteAvulso.Value, txtAdicionaisAvulso.Value, "0,00", "0,00", _
            txtTotalM2Avulso.Value), "0.00"))
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
    ' Variaveis do medoto
    Dim resposta As VbMsgBoxResult ' Variavel para confirmação na hora de cadastrar
    Dim polideiraAvulso As objPolideira
    Dim serrariaAvulso As objSerraria
    Dim pedreiraAvulso As objPedreira
    Dim listaTamanhoChapaAvulso As Collection
    Dim polimento As objTipoPolimento
    Dim tamanhoChapaAvulso As objTamanho
    Dim chapaAvulsa As objChapa
    Dim chapaAvulsaPesquisa As objChapa
    Dim blocoPesquisa As objBloco
    Dim blocoLista As objBloco
    Dim listaObjeto As Collection
    Dim listaAvulsosCadastradosHoje As Collection
    Dim idsChapaAvulso As Collection
    Dim idChapa As String
    Dim descricaoChapa As String
    Dim tipoPolimento As String
    Dim primeiroNome As String
    Dim nomeStatus As String
    Dim nomeMaterial As String
    Dim valorTotalBloco As String
    Dim cadastro As Boolean
    Dim i As Integer
    
    ' Patrão true
    cadastro = True
    
    ' Captura do tipo
    If obAvulso.Value = True Then
        nomeMaterial = "AVULSO " & txtNomeBloco.Value
    Else
        nomeMaterial = "IMPORTADO " & txtNomeBloco.Value
    End If
    
    ' Validações
    ' Verifica o Número do bloco na pedreira
    If txtIdBlocoAvulso.Value = "" Or txtIdBlocoAvulso.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtIdBlocoAvulso, NUMERO_BLOCO_PEDREIRA_MENSAGEM, NUMERO_BLOCO_PEDREIRA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtIdBlocoAvulso
    
    ' Verifica nome do bloco
    If txtMaterialAvulso.Value = "" Or txtMaterialAvulso.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtMaterialAvulso, NOME_AVULSO_MENSAGEM, NOME_AVULSO_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtMaterialAvulso
        
    ' Verifica se um cadastrado ou edição
    Set blocoPesquisa = daoBloco.pesquisarPorId(txtIdBlocoAvulso.Value, True)
    If blocoPesquisa.idSistema = txtIdBlocoAvulso.Value Then
        ' Mensagem de id já cadastrado no sistema
        errorStyle.Informativo AVULSO_JA_CADASTRADO_MENSAGEM, AVULSO_JA_CADASTRADO_TITULO
        Exit Sub
    Else
        ' Mensagem de confirmação
        resposta = MsgBox(CONFIRMACAO_CADASTRO_MENSAGEM, vbQuestion + vbYesNo, CONFIRMACAO_CADASTRO_TITULO)
    End If
    
    ' Verifica a confirmação do usário para poder cadastrar
    If resposta = vbYes Then
        ' Criação dos objetos
        Set polideiraAvulso = daoPolideira.pesquisarPorNome("AVULSO")
        Set serrariaAvulso = daoSerrada.pesquisarPorNome("AVULSO")
        Set pedreiraAvulso = daoPedreira.pesquisarPorNome("AVULSO")
        Set polimento = daoTipoPolimento.pesquisarPorNome(cbTipoPolimentoL.Value)
        Set tipoMaterial = daoTipoMaterial.pesquisarPorNome(cbTipoMaterialL.Value)
        Set statusObj = daoStatus.pesquisarPorNome("ESTOQUE")
        Set estoque = daoEstoqueM3.pesquisarPorNome("CASA DO GRANITO")
        Set estoqueChapa = daoEstoqueChapa.pesquisarPorNome("CASA DO GRANITO")
        Set bloco = ObjectFactory.factoryBloco(bloco)
        Set blocoPesquisa = ObjectFactory.factoryBloco(blocoPesquisa)
        Set chapaAvulsa = ObjectFactory.factoryChapa(chapaAvulsa)
        Set tamanhoChapaAvulso = ObjectFactory.factoryTamanho(tamanhoChapaAvulso)
        Set listaTamanhoChapaAvulso = ObjectFactory.factoryLista(listaTamanhoChapaAvulso)
        
        ' Calcula valor total bloco
        valorTotalBloco = M_METODOS_GLOBAL.formatarComPontos(Format(custoBloco( _
                    txtValorBloco.Value, txtValorFreteBloco.Value, "0", "0", txtAdicionais.Value), "0.00"))
        
        ' Criação do objeto bloco
        bloco.dataCadastro = txtDataCadastroChapaAvulsa.Value
        bloco.idSistema = txtIdBlocoAvulsoSistema.Value
        bloco.numeroBlocoPedreira = txtIdBlocoAvulso.Value
        bloco.nomeMaterial = nomeMaterial & txtMaterialAvulso.Value
        bloco.nota = cbTemNotaAvulso.Value
        bloco.observacao = txtObsBlocoL.Value
        bloco.compBrutoChapaPolida = txtComprimentoChapaAvulsa.Value
        bloco.altBrutoChapaPolida = txtAlturaChapaAvulsa.Value
        bloco.qtdChapas = txtQuantidadeChapasAvulsas.Value
        bloco.compBrutoChapaBruta = txtCompChapasBrutasAvulso.Value
        bloco.altBrutoChapaBruta = txtAlturaChapasBrutasAvulso.Value
        bloco.qtdM2Polimento = txtTotalM2Avulso.Value
        bloco.valoresAdicionais = txtAdicionaisAvulso.Value
        bloco.freteBloco = txtValorFreteAvulso.Value
        bloco.valorMetroPolimento = txtCustoSimplesM2Avulso.Value
        bloco.valorMetroPolimento = valorTotalBloco
        bloco.valorBloco = valorTotalBloco
        bloco.consultarCustoMedio = "NÃO"
        
        Set bloco.estoque = estoque
        Set bloco.pedreira = pedreiraAvulso
        Set bloco.polideira = polideiraAvulso
        Set bloco.serraria = serrariaAvulso
        Set bloco.status = statusObj
        Set bloco.tipoMaterial = tipoMaterial
        
        ' Criação do objeto chapa
        If cbTipoPolimentoL.Value = "POLIDO" Then
            idChapa = M_METODOS_GLOBAL.formatarIdChapa(txtIdBlocoAvulsoSistema.Value, "PO")
            tipoPolimento = "POLIDO"
            If obAvulso.Value = True Then
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaAvulso(bloco.nomeMaterial, "POLIDO")
            Else
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaImportado(bloco.nomeMaterial, "POLIDO")
            End If
    
        ElseIf cbTipoPolimentoL.Value = "BI POLIDO" Then
            idChapa = M_METODOS_GLOBAL.formatarIdChapa(txtIdBlocoAvulsoSistema.Value, "BPO")
            tipoPolimento = "BI POLIDO"
            If obAvulso.Value = True Then
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaAvulso(bloco.nomeMaterial, "BI POLIDO")
            Else
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaImportado(bloco.nomeMaterial, "BI POLIDO")
            End If
    
        ElseIf cbTipoPolimentoL.Value = "ESCOVADO" Then
            idChapa = M_METODOS_GLOBAL.formatarIdChapa(txtIdBlocoAvulsoSistema.Value, "ES")
            tipoPolimento = "ESCOVADO"
            If obAvulso.Value = True Then
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaAvulso(bloco.nomeMaterial, "ESCOVADO")
            Else
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaImportado(bloco.nomeMaterial, "ESCOVADO")
            End If
    
        ElseIf cbTipoPolimentoL.Value = "BI ESCOVADO" Then
            idChapa = M_METODOS_GLOBAL.formatarIdChapa(txtIdBlocoAvulsoSistema.Value, "BES")
            tipoPolimento = "BI ESCOVADO"
            If obAvulso.Value = True Then
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaAvulso(bloco.nomeMaterial, "BI ESCOVADO")
            Else
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaImportado(bloco.nomeMaterial, "BI ESCOVADO")
            End If
    
        ElseIf cbTipoPolimentoL.Value = "LEVIGADO" Then
            idChapa = M_METODOS_GLOBAL.formatarIdChapa(txtIdBlocoAvulsoSistema.Value, "LE")
            tipoPolimento = "LEVIGADO"
            If obAvulso.Value = True Then
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaAvulso(bloco.nomeMaterial, "LEVIGADO")
            Else
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaImportado(bloco.nomeMaterial, "LEVIGADO")
            End If
    
        ElseIf cbTipoPolimentoL.Value = "FLAMIADO" Then
            idChapa = M_METODOS_GLOBAL.formatarIdChapa(txtIdBlocoAvulsoSistema.Value, "FL")
            tipoPolimento = "FLAMIADO"
            If obAvulso.Value = True Then
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaAvulso(bloco.nomeMaterial, "FLAMIADO")
            Else
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaImportado(bloco.nomeMaterial, "FLAMIADO")
            End If
    
        ElseIf cbTipoPolimentoL.Value = "RIPADO" Then
            idChapa = M_METODOS_GLOBAL.formatarIdChapa(txtIdBlocoAvulsoSistema.Value, "RI")
            tipoPolimento = "RIPADO"
            If obAvulso.Value = True Then
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaAvulso(bloco.nomeMaterial, "RIPADO")
            Else
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaImportado(bloco.nomeMaterial, "RIPADO")
            End If
    
        ElseIf cbTipoPolimentoL.Value = "MATTE" Then
            idChapa = M_METODOS_GLOBAL.formatarIdChapa(txtIdBlocoAvulsoSistema.Value, "MA")
            tipoPolimento = "MATTE"
            If obAvulso.Value = True Then
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaAvulso(bloco.nomeMaterial, "MATTE")
            Else
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaImportado(bloco.nomeMaterial, "MATTE")
            End If
            
        ElseIf cbTipoPolimentoL.Value = "RESIN PINTADO" Then
            idChapa = M_METODOS_GLOBAL.formatarIdChapa(txtIdBlocoAvulsoSistema.Value, "RP")
            tipoPolimento = "RESIN PINTADO"
            If obAvulso.Value = True Then
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaAvulso(bloco.nomeMaterial, "RESIN PINTADO")
            Else
                descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapaImportado(bloco.nomeMaterial, "RESIN PINTADO")
            End If
        End If
        
        chapaAvulsa.idSistema = idChapa
        chapaAvulsa.nomeMaterial = descricaoChapa
        chapaAvulsa.valorTotal = valorTotalBloco
        chapaAvulsa.numeroBlocoPedreira = txtIdBlocoAvulso.Value
        Set chapaAvulsa.tipoPolimento = polimento
        Set chapaAvulsa.bloco = bloco
        chapaAvulsa.estoqueZero = "NÃO"
        
        ' Criação do objeto tamanho
        tamanhoChapaAvulso.compremento = txtComprimentoChapaAvulsa.Value
        tamanhoChapaAvulso.altura = txtAlturaChapaAvulsa.Value
        tamanhoChapaAvulso.qtdEstoque = txtQuantidadeChapasAvulsas.Value
        tamanhoChapaAvulso.qtdM2 = txtTotalM2Avulso.Value
        tamanhoChapaAvulso.valorPolimento = txtCustoSimplesM2Avulso.Value
        tamanhoChapaAvulso.espessura = txtEspessuraAvulso.Value
        
        Set tamanhoChapaAvulso.chapa = chapaAvulsa
        Set tamanhoChapaAvulso.tipoMaterial = tipoMaterial
        Set tamanhoChapaAvulso.polideira = polideira
        Set tamanhoChapaAvulso.estoque = estoqueChapa
        
        ' Adiciona na lista
        listaTamanhoChapaAvulso.Add tamanhoChapaAvulso
        
        ' Atribuições
        Set chapaAvulsa.tamanhos = listaTamanhoChapaAvulso
        Set tamanhoChapaAvulso.chapa = chapaAvulsa
        
        ' Chama serviço para cadastrar bloco
        Call daoBloco.cadastrarEEditar(bloco)
        ' Chama serviço para cadastrar chapa
        Call daoChapa.cadastrarEEditar(chapa)
        
        ' Verifica se foi um cadastro ou edição para personalisar as mensagens
        If cadastro = True Then
            ' Verifica se bloco foi cadastrado
            Set blocoPesquisa = daoBloco.pesquisarPorId(bloco.idSistema, False)
            Set chapaAvulsaPesquisa = daoChapa.pesquisarPorId(chapaAvulsa.idSistema)
            
            If blocoPesquisa.idSistema = txtIdBlocoSistema.Value And chapaAvulsaPesquisa.idSistema = idChapa Then
                ' Limpa os campos
                Call limparCamposCadastroBlocos
                
                ' Mensagem de cadastro realizado com sucesso.
                errorStyle.Informativo CADASTRO_CONFIRMADO_MENSAGEM, CADASTRO_CONFIRMADO_TITULO
            Else
                ' Mensagem de erro desconhecido
                errorStyle.Informativo ERRO_DESCONHECIDO_MENSAGEM, ERRO_DESCONHECIDO_TITULO
                Exit Sub
            End If
        End If
        
        ' Pesquisa avulsos cadastrado no dia atual
        Set listaObjeto = daoBloco.listarBlocosFilter(Date, Date, "", "", "", "", "", "", "", "", "", "", "")
        Set idsChapaAvulso = ObjectFactory.factoryLista(idsChapaAvulso)
        
        ' Seleciona só os avulsos
        For i = 1 To listaObjeto.Count
            ' Seta bloco da lista
            Set blocoLista = listaObjeto.Item(i)
            ' Captura o primeiro nome da descrição
            primeiroNome = Mid(blocoLista.nomeMaterial, 1, 5)
            ' Confere se é um avulso ou importado
            If primeiroNome <> "BLOCO" Then
                ' Captura as chapa avulso/importado para pesquisa
                idsChapaAvulso.Add blocoLista.numeroBlocoPedreira
            End If
        Next i
        
        If idsChapaAvulso.Count = 0 Or idsChapaAvulso.Count = -1 Then
            ' Apanas cria o objeto
            Set listaAvulsosCadastradosHoje = ObjectFactory.factoryLista(listaAvulsosCadastradosHoje)
        Else
            ' Pesquisa pelas chapas avulsas e importadas
            Set listaAvulsosCadastradosHoje = daoChapa.pesquisarPorListaIdsPedreira(idsChapaAvulso)
        End If
        
        ' Chama metodo para carregar lista com avulsos cadastros do dia atual
        Call carregarList(ListMateriais, listaAvulsosCadastradosHoje)
        
        ' Libera espaço na memoria
        Set pedreira = Nothing
        Set serraria = Nothing
        Set tipoMaterial = Nothing
        Set statusObj = Nothing
        Set estoque = Nothing
        Set bloco = Nothing
        Set blocoPesquisa = Nothing
        Set polimento = Nothing
        Set statusObj = Nothing
        Set chapaAvulsa = Nothing
        Set chapaAvulsaPesquisa = Nothing
        Set tamanhoChapaAvulso = Nothing
        Set listaTamanhoChapaAvulso = Nothing
        Set listaObjeto = Nothing
        Set listaAvulsosCadastradosHoje = Nothing
        Set idsChapaAvulso = Nothing
        
    Else
        ' Coloque o código a ser executado se o usuário clicar em "Não" aqui.
        errorStyle.Informativo ACAO_CANCELADA_MENSAGEM, ACAO_CANCELADA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If

    ' Deixa o cursor no txtIdBlocoAvulso para proximo cadastro
    txtIdBlocoAvulso.SetFocus
End Sub
' Botão btnLTxtVoltarCadatradoChapasAvulso tela cadastro avulso
Private Sub btnLTxtVoltarCadatradoChapasAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage - tela estoque m²
    Me.MultiPageCEBC.Value = 4
    ' Seta o foco
    txtMaterialChapaPesquisa.SetFocus
End Sub
' Botão btnLTxtLimparCadastroChapaAvulso tela cadastro avulso
Private Sub btnLTxtLimparCadastroChapaAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    Call limparCamposCadastroAvulso
End Sub

'-----------------------------------------------------------------TELA LANÇAMENTO E EDIÇÃO CHAPA-----------------------------------
'                                                                 ------------------------------
' Carrega os campos com os dados da chapa tela lançamento e edição chapa
Private Sub carregarDadosChapaTelaEdicaoChapa(chapa As objChapa, bloco As objBloco)
    ' Variaveis do medoto
    Dim i As Integer
    
    ' Dados bloco
    txtIdBlocoPedreiraChapa.Value = bloco.idSistema
    txtDecricaoBlocoChapa.Value = bloco.nomeMaterial
    txtQtdDisponivelChapaBloco.Value = bloco.qtdChapas
    txtNBlocoPedreiraChapa.Value = bloco.numeroBlocoPedreira
    txtTipoMaterialChapa.Value = bloco.tipoMaterial.nome
    txtCompBrutoChapa.Value = bloco.compBrutoChapaBruta
    txtAlturaBrutaChapa.Value = bloco.altBrutoChapaBruta
    txtQtsM2Chapa.Value = bloco.qtdM2Serrada
    
    ' Dados chapa
    txtIdChapaSistema.Value = chapa.idSistema
    txtDescricaoChapa.Value = chapa.nomeMaterial
    
    ' Dimensões e custos
    'Call selecaoItem("cbPolideiraChapa", chapa.polideira.nome)
    Call selecaoItem("cbTipoPolimentoChapa", chapa.tipoPolimento.nome)
    
    ' Carrega lista com tamanhos das chapas
    Call carregarListTamanhosChapas(ListTamanhosChapas, chapa.tamanhos) ' Irá enviar id chapa para carregamento
        
    ' Verifica se a lista só tem um tamanho
    If chapa.tamanhos.Count = 1 Then
        For i = 1 To chapa.tamanhos.Count
            ' Seta o ojeto
            Set tamanho = chapa.tamanhos(i)

            ' Tamanho único
            Call selecaoItem("cbTipoMaterialChapaC", tamanho.tipoMaterial.nome)
            txtCompBrutoChapa.Value = M_METODOS_GLOBAL.formatarComPontos(Format(tamanho.compremento, "0.0000"))
            txtAlturaBrutaChapa.Value = M_METODOS_GLOBAL.formatarComPontos(Format(tamanho.altura, "0.0000"))
            txtEspTiposMateriaisChapa.Value = tamanho.espessura

            ' Libera espaço na memoria
            Set tamanho = Nothing
        Next i
    End If
    
    ' Se Status do bloco for finalizado deixar visivel lBlocoFinalizadoChapa e cbAbrirParaEdicao e desabilitar todos os campos
    If bloco.status.nome = "FECHADO" Then
        lBlocoFinalizadoChapa.Visible = True
        cbAbrirParaEdicao.Visible = True
        Call desabilitaCamposChapas
    Else
        lBlocoFinalizadoChapa.Visible = False
        cbAbrirParaEdicao.Visible = False
        Call habilitaCamposChapas
    End If
End Sub

' Habilita e desabilita campos para edição tela lançamento e edição de chapa
Private Sub cbAbrirParaEdicao_Click()
    If cbAbrirParaEdicao.Value = True Then
        Call habilitaCamposChapas
    Else
        Call desabilitaCamposChapas
    End If
End Sub

' Logica para criar id da chapa no sistema tela lançamento e edição de chapa
Private Sub cbTipoPolimentoChapa_Change()
    
    ' Varuaveis do metodo
    Dim idChapas As String
    Dim descricaoChapa As String
    Dim codFinal As String
    Dim posicao As Integer
    Dim idBloco As String
    Dim descricao As String
    
    ' Id do bloco
    idBloco = txtIdBlocoPedreiraChapa.Value
    descricao = txtDecricaoBlocoChapa.Value
     
    'Captura o tipo de polimento, cria o id e descrição da chapa
    If cbTipoPolimentoChapa.Value = "BRUTO" Then
        idChapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "BT")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "BRUTO")
        
    ElseIf cbTipoPolimentoChapa.Value = "BI POLIDO" Then
        idChapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "BPO")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "BI POLIDO")
        
    ElseIf cbTipoPolimentoChapa.Value = "ESCOVADO" Then
        idChapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "ES")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "ESCOVADO")
        
    ElseIf cbTipoPolimentoChapa.Value = "BI ESCOVADO" Then
        idChapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "BES")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "BI ESCOVADO")
        
    ElseIf cbTipoPolimentoChapa.Value = "LEVIGADO" Then
        idChapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "LE")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "LEVIGADO")
        
    ElseIf cbTipoPolimentoChapa.Value = "FLAMIADO" Then
        idChapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "FL")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "FLAMIADO")
        
    ElseIf cbTipoPolimentoChapa.Value = "RIPADO" Then
        idChapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "RI")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "RIPADO")
        
    ElseIf cbTipoPolimentoChapa.Value = "POLIDO" Then
        idChapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "PO")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "POLIDO")
        
    ElseIf cbTipoPolimentoChapa.Value = "MATTE" Then
        idChapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "MA")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "MATTE")
        
    ElseIf cbTipoPolimentoChapa.Value = "RESIN PINTADO" Then
        idChapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "RP")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "RESIN PINTADO")
    End If
    
    ' Seta id e descrição
    txtIdChapaSistema.Value = idChapas
    txtDescricaoChapa.Value = descricaoChapa
End Sub

' Formata retorno txtCompBrutoChapa tela lançamento e edição chapa
Private Sub txtCompBrutoChapa_Change()
    ' Define o resultado no TextBox
    txtCompBrutoChapa_Change.Value = M_METODOS_GLOBAL.formatarMetros(txtCompBrutoChapa.Value)
    
'    'Retorna valor calculado e formatado
'    txtQtsM2Chapa.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM2( _
'        txtCompBrutoChapa.Value, txtAlturaBrutaChapa.Value, TextBoxQuantidadeChapasAvulsas.Value), "0.0000"))
End Sub

' Formata retorno txtAlturaBrutaChapa tela lançamento e edição chapa
Private Sub txtAlturaBrutaChapa_Change()
    ' Define o resultado no TextBox
    txtAlturaBrutaChapa.Value = M_METODOS_GLOBAL.formatarMetros(txtAlturaBrutaChapa.Value)
End Sub

' Formata retorno txtCompTipoMateriaisChapa tela lançamento e edição chapa
Private Sub txtCompTipoMateriaisChapa_Change()
    ' Define o resultado no TextBox
    txtCompTipoMateriaisChapa.Value = M_METODOS_GLOBAL.formatarMetros(txtCompTipoMateriaisChapa.Value)
    
    ' Retorna valor calculado e formatado
    txtQtdM2TipoMateriaisChapas.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM2( _
        txtCompTipoMateriaisChapa.Value, txtAltTipoMateriaisChapa.Value, txtQtdTipoMateriaisChapas.Value), "0.0000"))
End Sub

' Formata retorno txtAltTipoMateriaisChapa tela lançamento e edição chapa
Private Sub txtAltTipoMateriaisChapa_Change()
    ' Define o resultado no TextBox
    txtAltTipoMateriaisChapa.Value = M_METODOS_GLOBAL.formatarMetros(txtAltTipoMateriaisChapa.Value)
    
    ' Retorna valor calculado e formatado
    txtQtdM2TipoMateriaisChapas.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM2( _
        txtCompTipoMateriaisChapa.Value, txtAltTipoMateriaisChapa.Value, txtQtdTipoMateriaisChapas.Value), "0.0000"))
End Sub

' Formata retorno txtQtdTipoMateriaisChapas tela lançamento e edição chapa
Private Sub txtQtdTipoMateriaisChapas_Change()
    ' Variaveis do metodo
    Dim textoDigitado As String
    Dim textoFormatado As String
    Dim i As Integer

    'Recebi o texto digitadado pelo usúario
    textoDigitado = txtQtdTipoMateriaisChapas.Value
 
    'Remove todos os caracteres não numéricos
    For i = 1 To Len(textoDigitado)
        If IsNumeric(Mid(textoDigitado, i, 1)) Then
            textoFormatado = textoFormatado & Mid(textoDigitado, i, 1)
        End If
        
        'Remove o zero na esquerda do texto
        If Len(textoFormatado) = 2 Then
            If Left(textoFormatado, 1) = 0 Then
                textoFormatado = Mid(textoFormatado, 2, 1)
            End If
        End If
    Next i
    
    If textoFormatado = "" Or textoFormatado = " " Then
        textoFormatado = "0"
    End If
    
    txtQtdTipoMateriaisChapas.Value = textoFormatado
    
    ' Retorna valor calculado e formatado
    txtQtdM2TipoMateriaisChapas.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM2( _
        txtCompTipoMateriaisChapa.Value, txtAltTipoMateriaisChapa.Value, txtQtdTipoMateriaisChapas.Value), "0.0000"))
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
        
    ' Validações
    ' Verifica o Polideira
    If cbTipoPolimentoChapa.Value = "" Or cbTipoPolimentoChapa.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleComboBox cbTipoPolimentoChapa, SELECIONE_POLIMENTO_MENSAGEM, SELECIONE_POLIMENTO_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleComboBox cbTipoPolimentoChapa
    
    ' Verifica o complimento
    If txtCompTipoMateriaisChapa.Value = "0,0000" Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtCompTipoMateriaisChapa, INFORMACAO_INVALIDA_MENSAGEM, INFORMACAO_INVALIDA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtCompTipoMateriaisChapa
    
    ' Verifica o altura
    If txtAltTipoMateriaisChapa.Value = "0,0000" Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtAltTipoMateriaisChapa, INFORMACAO_INVALIDA_MENSAGEM, INFORMACAO_INVALIDA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtAltTipoMateriaisChapa
    
    ' Verifica o espessura
    If txtEspTiposMateriaisChapa.Value = "" Or txtEspTiposMateriaisChapa.Value = " " Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtEspTiposMateriaisChapa, INFORMACAO_INVALIDA_MENSAGEM, INFORMACAO_INVALIDA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtEspTiposMateriaisChapa
    
    ' Verifica o quantidade
    If txtQtdTipoMateriaisChapas.Value = "0" Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtQtdTipoMateriaisChapas, INFORMACAO_INVALIDA_MENSAGEM, INFORMACAO_INVALIDA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtQtdTipoMateriaisChapas
    
    ' Verifica o custo
    If txtCustoChapa.Value = "0,00" Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtCustoChapa, INFORMACAO_INVALIDA_MENSAGEM, INFORMACAO_INVALIDA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtCustoChapa
    
    ' Comparação para fazer se é um novo ou edição
    If lTamanhoCadastroEdicao.Caption = "-1" Then
        ' Adiciona uma linha
        ListTamanhosChapas.AddItem
        
        ' Adiciona os dados do bloco
        ListTamanhosChapas.list(ListTamanhosChapas.ListCount - 1, 0) = cbTipoMaterialChapaC.Value
        ListTamanhosChapas.list(ListTamanhosChapas.ListCount - 1, 1) = M_METODOS_GLOBAL.formatarComPontos( _
                                    Format(txtCompTipoMateriaisChapa.Value, "0.0000"))
        ListTamanhosChapas.list(ListTamanhosChapas.ListCount - 1, 2) = M_METODOS_GLOBAL.formatarComPontos( _
                                    Format(txtAltTipoMateriaisChapa.Value, "0.0000"))
        ListTamanhosChapas.list(ListTamanhosChapas.ListCount - 1, 3) = M_METODOS_GLOBAL.formatarComPontos( _
                                    Format(txtQtdM2TipoMateriaisChapas.Value, "0.0000"))
        ListTamanhosChapas.list(ListTamanhosChapas.ListCount - 1, 4) = txtQtdTipoMateriaisChapas.Value
        ListTamanhosChapas.list(ListTamanhosChapas.ListCount - 1, 5) = txtEspTiposMateriaisChapa.Value
        ListTamanhosChapas.list(ListTamanhosChapas.ListCount - 1, 6) = M_METODOS_GLOBAL.formatarComPontos( _
                                    Format(txtCustoChapa.Value, "currency"))
        ListTamanhosChapas.list(ListTamanhosChapas.ListCount - 1, 7) = cbPolideiraChapa.Value
        ListTamanhosChapas.list(ListTamanhosChapas.ListCount - 1, 8) = cbEstoqueChapaC.Value
    Else
        ' Faz atualização na lista
        ListTamanhosChapas.list(CInt(lTamanhoCadastroEdicao.Caption), 0) = cbTipoMaterialChapaC.Value
        ListTamanhosChapas.list(CInt(lTamanhoCadastroEdicao.Caption), 1) = M_METODOS_GLOBAL.formatarComPontos( _
                                    Format(txtCompTipoMateriaisChapa.Value, "0.0000"))
        ListTamanhosChapas.list(CInt(lTamanhoCadastroEdicao.Caption), 2) = M_METODOS_GLOBAL.formatarComPontos( _
                                    Format(txtAltTipoMateriaisChapa.Value, "0.0000"))
        ListTamanhosChapas.list(CInt(lTamanhoCadastroEdicao.Caption), 3) = M_METODOS_GLOBAL.formatarComPontos( _
                                    Format(txtQtdM2TipoMateriaisChapas.Value, "0.0000"))
        ListTamanhosChapas.list(CInt(lTamanhoCadastroEdicao.Caption), 4) = txtQtdTipoMateriaisChapas.Value
        ListTamanhosChapas.list(CInt(lTamanhoCadastroEdicao.Caption), 5) = txtEspTiposMateriaisChapa.Value
        ListTamanhosChapas.list(CInt(lTamanhoCadastroEdicao.Caption), 6) = M_METODOS_GLOBAL.formatarComPontos( _
                                    Format(txtCustoChapa.Value, "currency"))
        ListTamanhosChapas.list(CInt(lTamanhoCadastroEdicao.Caption), 7) = cbPolideiraChapa.Value
        ListTamanhosChapas.list(CInt(lTamanhoCadastroEdicao.Caption), 8) = cbEstoqueChapaC.Value
                
        ' Volta valor original
        lTamanhoCadastroEdicao.Caption = "-1"
    End If
    ' Limpa os campos de tamanho
    Call limparCamposTamanhoChapa
End Sub

' Botão btnLTxtEditarTamanhoChapa tela lançamento e edição chapa
Private Sub btnLTxtEditarTamanhoChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Variaveis do metodo
    Dim linha As Double
    
    ' Verifica se tem algum item selecionado
    If Me.ListTamanhosChapas.ListIndex = -1 Then
        ' Mensagem usuário
        errorStyle.Informativo SELECIONE_TAMANHO_MENSAGEM, SELECIONE_TAMANHO_TITULO
        Exit Sub
    End If
    
    ' Captura a linha selecionada
    linha = ListTamanhosChapas.ListIndex
    ' Seta linha para edição
    lTamanhoCadastroEdicao.Caption = linha
    
    ' Adiciona nos campos para edição
    Call selecaoItem("cbTipoMaterialChapaC", ListTamanhosChapas.list(linha, 0))
    txtCompTipoMateriaisChapa.Value = ListTamanhosChapas.list(linha, 1)
    txtAltTipoMateriaisChapa.Value = ListTamanhosChapas.list(linha, 2)
    txtQtdM2TipoMateriaisChapas.Value = ListTamanhosChapas.list(linha, 3)
    txtQtdTipoMateriaisChapas.Value = ListTamanhosChapas.list(linha, 4)
    txtEspTiposMateriaisChapa.Value = ListTamanhosChapas.list(linha, 5)
    txtCustoChapa.Value = ListTamanhosChapas.list(linha, 6)
    Call selecaoItem("cbPolideiraChapa", ListTamanhosChapas.list(linha, 7))
    Call selecaoItem("cbEstoqueChapaC", ListTamanhosChapas.list(linha, 8))
End Sub

' Botão btnLTxtTirarDaLista tela lançamento e edição chapa
Private Sub btnLTxtTirarDaLista_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Variaveis do metodo
    Dim linha As Double
    
    ' Verifica se tem algum item selecionado
    If Me.ListTamanhosChapas.ListIndex = -1 Then
        ' Mensagem usuário
        errorStyle.Informativo SELECIONE_TAMANHO_MENSAGEM, SELECIONE_TAMANHO_TITULO
        Exit Sub
    End If
    
    ' Captura a linha clicada
    linha = ListTamanhosChapas.ListIndex
    
    ' Verifica se tem algum item selecionado
    If Me.ListTamanhosChapas.ListIndex <> -1 Then
        ' Remove o item selecionado
        ListTamanhosChapas.RemoveItem linha
    End If
End Sub

' Botão btnLTxtSalvarChapa tela lançamento e edição chapa
Private Sub btnLTxtSalvarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Variaveis do medoto
    Dim listaTamanhosAtualizados As Collection
    Dim custoTotal As Double
    Dim i As Integer
    
    ' Verificando se tem algum tamanho adicionado
    If ListTamanhosChapas.ListCount = 0 Then
        ' Deixa visivel o erro com mensagens
        errorStyle.Informativo TAMANHO_CHAPA_MENSAGEM, TAMANHO_CHAPA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    
    ' Cria as variaveis
    Set listaTamanhosAtualizados = ObjectFactory.factoryLista(listaTamanhosAtualizados)
    Set bloco = daoBloco.pesquisarPorId(txtIdBlocoPedreiraChapa.Value, True)
    Set tipoPolimento = daoTipoPolimento.pesquisarPorNome(cbTipoPolimentoChapa.Value)
    Set chapa = daoChapa.pesquisarPorId(txtIdChapaSistema.Value)
    
    ' Adiciona campos em tamanho
    For i = 0 To ListTamanhosChapas.ListCount - 1
        ' Cria objeto
        Set tamanho = ObjectFactory.factoryTamanho(tamanho)
        ' Seta os dados
        Set tipoMaterial = daoTipoMaterial.pesquisarPorNome(ListTamanhosChapas.list(i, 0))
        tamanho.compremento = ListTamanhosChapas.list(i, 1)
        tamanho.altura = ListTamanhosChapas.list(i, 2)
        tamanho.qtdM2 = ListTamanhosChapas.list(i, 3)
        tamanho.qtdEstoque = ListTamanhosChapas.list(i, 4)
        tamanho.espessura = ListTamanhosChapas.list(i, 5)
        tamanho.espessura = ListTamanhosChapas.list(i, 5)
        tamanho.valorPolimento = ListTamanhosChapas.list(i, 6)
        Set polideira = daoPolideira.pesquisarPorNome(ListTamanhosChapas.list(i, 7))
        Set estoqueChapa = daoEstoqueChapa.pesquisarPorNome(ListTamanhosChapas.list(i, 8))
        ' Se tiver codigo adiciona
        If IsNull(ListTamanhosChapas.list(i, 9)) Then
        Else
            custoTotal = custoTotal + CDbl(ListTamanhosChapas.list(i, 6))
        End If
        ' Seta tamanho
        listaTamanhosAtualizados.Add tamanho
        
        ' Libera memoria
        Set tamanho = Nothing
        Set tipoMaterial = Nothing
        Set polideira = Nothing
        Set estoqueChapa = Nothing
    Next i
    
    ' Seta os dados em chapa
    chapa.idSistema = txtIdChapaSistema.Value
    chapa.nomeMaterial = txtDescricaoChapa.Value
    chapa.valorTotal = txtTotalChapas.Value
    chapa.nomeMaterial = bloco.numeroBlocoPedreira

    If txtEstoqueChapa.Value = "0" Then
        chapa.estoqueZero = "SIM"
    Else
        chapa.estoqueZero = "NÃO"
    End If
    
    ' Atribuições
    Set chapa.bloco = bloco
    Set tipoPolimento = tipoPolimento
    Set chapa.tamanhos = listaTamanhosAtualizados
    
    ' Cadastra ou atualiza chapa
    Call daoChapa.cadastrarEEditar(chapa)
    
    Set chapa = daoChapa.pesquisarPorId(chapa.idSistema)
    ' Recarrega a tela com dados atualizados
    Call carregarDadosChapaTelaEdicaoChapa(chapa, chapa.bloco)
    
    ' Deixa visivel o erro com mensagens
    errorStyle.Informativo SUCESSO_CADASTRO_EDICAO_MENSAGEM, SUCESSO_CADASTRO_EDICAO_TITULO
End Sub

' Botão btnLTxtVoltarChapa tela lançamento e edição chapa
Private Sub btnLTxtVoltarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Volta para tela que chamou
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = paginaAnterior
        
    ' Limpa os campos
    Call limparCamposChapa
    Call limparCamposTamanhoChapa
    
    ' Volta para pagina que chamou
    If paginaAnterior = 1 Then
        ' Botão chapa
        formControle.Controls("btnLMenuBloco").BackColor = RGB(200, 230, 255)
        formControle.Controls("btnLMenuBloco").Font.Size = 32
        formControle.Controls("btnLMenuBloco").Font.Size = 20
        formControle.Controls("btnLMenuBloco").Left = 15
        formControle.Controls("btnLMenuBloco").Width = 172
        formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignCenter
                    
        ' Botão Menu
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
' Botão txtQtdMaterialParaTroca02 tela troca estoque
Private Sub txtQtdMaterialParaTroca02_Change()
    ' Variaveis do metodo
    Dim textoDigitado As String
    Dim textoFormatado As String
    Dim lancamento As Integer
    Dim estoque As Integer
    Dim i As Integer

    ' Recebi o texto digitadado pelo usúario
    textoDigitado = txtQtdMaterialParaTroca02.Value
 
    ' Remove todos os caracteres não numéricos
    For i = 1 To Len(textoDigitado)
        If IsNumeric(Mid(textoDigitado, i, 1)) Then
            textoFormatado = textoFormatado & Mid(textoDigitado, i, 1)
        End If
        
        'Remove o zero na esquerda do texto
        If Len(textoFormatado) = 2 Then
            If Left(textoFormatado, 1) = 0 Then
                textoFormatado = Mid(textoFormatado, 2, 1)
            End If
        End If
    Next i
    
    If textoDigitado = "" Then
        textoFormatado = "0"
    End If
    
    txtQtdMaterialParaTroca02.Value = textoFormatado
    
    'Conversão para comparação
    lancamento = CInt(txtQtdMaterialParaTroca02.Value)
    estoque = CInt(txtQtdDispovelMaterialParaTroca01.Value)
    
    'Confere se tem no estoque
    If lancamento > estoque Then
        ' Mensagem de retorno
        errorStyle.Informativo VALOR_SUPERIOR_MENSAGEM, VALOR_SUPERIOR_TITULO
        Exit Sub
    End If
End Sub

' Botão btnLTxtAdicionarTrocaEstoque tela troca estoque
Private Sub btnLTxtAdicionarTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Variaveis do metodo
    Dim chapaTroca As objChapa
    Dim idChapa As String
    Dim descricaoChapa As String
    Dim codFinal As String
    Dim posicao As Integer
    Dim idBloco As String
    Dim descricao As String
    Dim temChapa As Boolean
    
    If txtQtdMaterialParaTroca02.Value = "0" Then
        Exit Sub
    End If
    
    ' Id e descricao do bloco
    idBloco = chapa.bloco.idSistema
    descricao = chapa.bloco.nomeMaterial
     
    'Captura o tipo de polimento, cria o id e descrição da chapa
    If cbTipoPolimentoTroca.Value = "BI POLIDO" Then
        idChapa = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "BPO")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "BI POLIDO")
        
    ElseIf cbTipoPolimentoTroca.Value = "ESCOVADO" Then
        idChapa = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "ES")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "ESCOVADO")
        
    ElseIf cbTipoPolimentoTroca.Value = "BI ESCOVADO" Then
        idChapa = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "BES")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "BI ESCOVADO")
        
    ElseIf cbTipoPolimentoTroca.Value = "LEVIGADO" Then
        idChapa = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "LE")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "LEVIGADO")
        
    ElseIf cbTipoPolimentoTroca.Value = "FLAMIADO" Then
        idChapa = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "FL")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "FLAMIADO")
        
    ElseIf cbTipoPolimentoTroca.Value = "RIPADO" Then
        idChapa = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "RI")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "RIPADO")
        
    ElseIf cbTipoPolimentoTroca.Value = "POLIDO" Then
        idChapa = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "PO")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "POLIDO")
        
    ElseIf cbTipoPolimentoTroca.Value = "MATTE" Then
        idChapa = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "MA")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "MATTE")
        
    ElseIf cbTipoPolimentoTroca.Value = "RESIN PINTADO" Then
        idChapa = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "RP")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "RESIN PINTADO")
        
    ElseIf cbTipoPolimentoTroca.Value = "BRUTO" Then
        idChapa = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "BT")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "BRUTO")
    End If
    
    ' Pesquisa se chapa tem cadastro
    temChapa = daoChapa.temIdChapa(idChapa)
    
    If temChapa = True Then
        ' Seta chapa
        Set chapaTroca = daoChapa.pesquisarPorId(idChapa)
        
        Call carregarListTrocasQtdChapas(ListMateriaisParaTroca, chapa, tamanho)
        
        Call carregarListTrocasQtdChapas(ListTrocarPor, chapaTroca, tamanho)
        
    Else
        
    End If
End Sub
' Botão btnLTxtTrocarEstoque tela troca estoque
Private Sub btnLTxtTrocarEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Variaveis do metodo
    Dim chapaTroca As objChapa
    Dim chapaSerTrocada As objChapa
    Dim tamanhoTroca As objTamanho
    Dim tamanhoSerTrocada As objTamanho
    Dim tamanhoNovo As objTamanho
    Dim m2Diferenca As Double
    Dim totalCusto As Double
    Dim totalTamanho As Double
    Dim qtdDiferenca As Integer
    Dim i As Integer
    Dim j As Integer
    
    ' Verifica se tem algum item selecionado
    If Me.ListMateriaisParaTroca.ListCount = 0 Or Me.ListTrocarPor.ListCount = 0 Then
        ' Mensagem usuário
        errorStyle.Informativo ADICIONE_CHAPA_MENSAGEM, ADICIONE_CHAPA_TITULO
        Exit Sub
    End If
    
    ' Busca chapa e tamanho no banco para atualização
    Set chapaTroca = daoChapa.pesquisarPorId(ListMateriaisParaTroca.list(ListMateriaisParaTroca.ListCount - 1, 0))
    Set chapaSerTrocada = daoChapa.pesquisarPorId(ListTrocarPor.list(ListTrocarPor.ListCount - 1, 0))
    Set tamanhoTroca = daoTamanho.pesquisarPorIdTamanho(ListMateriaisParaTroca.list(ListMateriaisParaTroca.ListCount - 1, 4))
    Set tamanhoSerTrocada = daoTamanho.pesquisarPorIdTamanho(ListMateriaisParaTroca.list(ListMateriaisParaTroca.ListCount - 1, 4))
    Set tamanhoNovo = ObjectFactory.factoryTamanho(tamanhoNovo)
    
    ' Atualizações na memoria
    For i = 1 To chapaTroca.tamanhos.Count
        ' Seta o ojeto
        Set tamanho = chapaTroca.tamanhos(i)
        ' Comparação para atribuições
        If tamanho.id = tamanhoTroca.id Then
            ' Subtração
            qtdDiferenca = CInt(tamanho.qtdEstoque) - CDbl(ListMateriaisParaTroca.list(ListMateriaisParaTroca.ListCount - 1, 2))
            m2Diferenca = CDbl(tamanho.qtdM2) - CDbl(ListMateriaisParaTroca.list(ListMateriaisParaTroca.ListCount - 1, 3))
            
            ' Cofere se é pra trocar tudo
            If qtdDiferenca = 0 Then
                ' Atribuições
                chapaTroca.tamanhos.Remove (i)
                
                tamanhoSerTrocada.setChapa chapaSerTrocada
                chapaSerTrocada.tamanhos.Add tamanhoSerTrocada
                
                ' Atualizações custo
                For j = 1 To chapaTroca.tamanhos.Count
                    ' Seta o ojeto
                    Set tamanho = chapaTroca.tamanhos(j)

                    totalTamanho = CDbl(tamanho.qtdM2) * CDbl(tamanho.valorPolimento)

                    totalCusto = totalCusto + totalTamanho

                Next j

                ' Seta custo total
                chapaTroca.valorTotal = totalCusto
                
                ' Atualizações custo
                For j = 1 To chapaSerTrocada.tamanhos.Count
                    ' Seta o ojeto
                    Set tamanho = chapaSerTrocada.tamanhos(j)

                    totalTamanho = CDbl(tamanho.qtdM2) * CDbl(tamanho.valorPolimento)

                    totalCusto = totalCusto + totalTamanho

                Next j

                ' Seta custo total
                chapaSerTrocada.valorTotal = totalCusto
            Else
                ' Atribuições
                tamanho.qtdEstoque = qtdDiferenca
                tamanho.qtdM2 = m2Diferenca
                
                tamanhoNovo.id = "0"
                tamanhoNovo.compremento = tamanhoSerTrocada.compremento
                tamanhoNovo.altura = tamanhoSerTrocada.altura
                tamanhoNovo.qtdEstoque = qtdDiferenca
                tamanhoNovo.qtdM2 = m2Diferenca
                tamanhoNovo.valorPolimento = tamanhoSerTrocada.valorPolimento
                tamanhoNovo.espessura = tamanhoSerTrocada.espessura
                
                tamanhoNovo.setTipoMaterial tamanhoSerTrocada.tipoMaterial
                tamanhoNovo.setPolideira tamanhoSerTrocada.polideira
                tamanhoNovo.setEstoque tamanhoSerTrocada.estoque
                tamanhoNovo.setChapa chapaSerTrocada
                
                chapaSerTrocada.tamanhos.Add tamanhoNovo
                
                ' Atualizações custo
                For j = 1 To chapaTroca.tamanhos.Count
                    ' Seta o ojeto
                    Set tamanho = chapaTroca.tamanhos(j)

                    totalTamanho = CDbl(tamanho.qtdM2) * CDbl(tamanho.valorPolimento)

                    totalCusto = totalCusto + totalTamanho

                Next j

                ' Seta custo total
                chapaTroca.valorTotal = totalCusto
                
                ' Atualizações custo
                For j = 1 To chapaSerTrocada.tamanhos.Count
                    ' Seta o ojeto
                    Set tamanho = chapaSerTrocada.tamanhos(j)

                    totalTamanho = CDbl(tamanho.qtdM2) * CDbl(tamanho.valorPolimento)

                    totalCusto = totalCusto + totalTamanho

                Next j

                ' Seta custo total
                chapaSerTrocada.valorTotal = totalCusto
            End If
            ' Sai do for porque já achou o tamanho
            Exit For
        End If
    Next i
    
    ' Atualiazações no disco
    Call daoChapa.cadastrarEEditar(chapaTroca)
    Call daoChapa.cadastrarEEditar(chapaSerTrocada)
    
    ' Mensagem usuário
    errorStyle.Informativo TROCA_REALIZADA_MENSAGEM, TROCA_REALIZADA_TITULO
    
    ' Muda abra da multPage
    Me.MultiPageCEBC.Value = 4
    
    ' Atualiza a pesquisa
    Call pesquisarChapasFilter
End Sub
' Botão btnLTxtVoltarTrocaEstoque tela troca estoque
Private Sub btnLTxtVoltarTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Muda abra da multPage para tela estoque de chapas
    Me.MultiPageCEBC.Value = 4
End Sub

'-----------------------------------------------------------------TELA DESPACHE-----------------------------------
'                                                                 -------------
' Campo txtPesquisarMaterial de pesquisa de chapas tela despache
Private Sub txtPesquisarMaterial_Change()
    ' Converte para caixa alta
    txtPesquisarMaterial.Value = UCase(txtPesquisarMaterial.Value)
End Sub

' Campo txtPesquisarMaterial de pesquisa de chapas tela despache
Private Sub txtPesquisarMaterial_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Pesquisa e carrega ListBox
    Call pesquisarChapaDespachePorDescricao
    
    ' Limpa campos de tamanho da chapa
    Call limparCamposTamanho
End Sub

' Campo txtPesquisarPorNumeroBloco de pesquisa de chapas tela despache
Private Sub txtPesquisarPorNumeroBloco_Change()
    ' Converte para caixa alta
    txtPesquisarPorNumeroBloco.Value = UCase(txtPesquisarPorNumeroBloco.Value)
End Sub

' Campo txtPesquisarPorNumeroBloco de pesquisa de chapas tela despache
Private Sub txtPesquisarPorNumeroBloco_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Pesquisa e carrega ListBox
    Call pesquisarChapaDespachePorNumeroBloco
    
    ' Limpa campos de tamanho da chapa
    Call limparCamposTamanho
End Sub

' Campo txtPesquisarDespache de pesquisa de chapas tela despache
Private Sub txtPesquisarDespache_Change()
    ' Converte para caixa alta
    txtPesquisarDespache.Value = UCase(txtPesquisarDespache.Value)
End Sub

' Campo txtPesquisarDespache de pesquisa de chapas tela despache
Private Sub txtPesquisarDespache_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Pesquisa e carrega ListBox
    Call pesquisarDespacheSalvo
    ' Limpa campos de tamanho da chapa
    Call limparCamposTamanho
End Sub

' txtQuantidadeDespache tela despachar
Private Sub txtQuantidadeDespache_Change()
    ' Variaveis do metodo
    Dim textoDigitado As String
    Dim textoFormatado As String
    Dim linha As Integer
    Dim i As Integer
    
    ' Recebi o texto digitadado pelo usúario
    textoDigitado = txtQuantidadeDespache.Value
    
    ' Remove todos os caracteres não numéricos
    For i = 1 To Len(textoDigitado)
        If IsNumeric(Mid(textoDigitado, i, 1)) Then
            textoFormatado = textoFormatado & Mid(textoDigitado, i, 1)
        End If

        'Remove o zero na esquerda do texto
        If Len(textoFormatado) = 2 Then
            If Left(textoFormatado, 1) = 0 Then
                textoFormatado = Mid(textoFormatado, 2, 1)
            End If
        End If
    Next i
    
    If textoDigitado = "" Or textoDigitado = " " Then
        textoFormatado = "0"
    End If

    txtQuantidadeDespache.Value = textoFormatado
    
    ' Verifica se foi selecionado tamanho da chapa
    If ListTamanhosChapaDespache.ListIndex <> -1 Then
        ' Verifica a quantidade
        If txtQuantidadeDespache.Value = "0" Then
            ' Mensagem usuário
            errorStyle.EntrarErrorStyleTextBox txtQuantidadeDespache, SELECIONE_QTD_MENSAGEM, SELECIONE_QTD_TITULO
            Exit Sub
        End If
        ' Deixa na cor patrão
        errorStyle.sairErrorStyleTextBox txtQuantidadeDespache
        
        ' Captura a linha selecionada
        linha = ListTamanhosChapaDespache.ListIndex
        
        ' Retorna valor calculado e formatado
        txtQtdm2Despache.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM2( _
            ListTamanhosChapaDespache.list(linha, 1), ListTamanhosChapaDespache.list(linha, 2), txtQuantidadeDespache.Value), "0.0000"))
    End If
End Sub

' Limpar campos de tamanho da tela despeche
Private Sub limparCamposTamanho()
    ' Limpa campos de material
    txtNumeroBloco.Value = ""
    txtPesquisarMaterial.Value = ""
    txtPesquisarPorNumeroBloco.Value = ""
    txtPesquisarDespache.Value = ""
    txtMaterial.Value = ""
    txtIdSistemaChapaDespache.Value = ""
    txtQtdm2Despache.Value = "0,0000"
    
    ' Seta valor patrão para seleção das listas
    ListDespachado.ListIndex = -1
    ListEstoqueM2.ListIndex = -1
    ListTamanhosChapaDespache.ListIndex = -1
    ListTamanhosChapaDespache.Clear
    txtQuantidedadeChapa.Value = "0"
    txtQuantidadeDespache.Value = "0"
End Sub

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
    
    Dim qtdChapas As Integer
    Dim totalChapas As Integer
    
    ' Verifica se tem algum chapa selecionada
    If ListEstoqueM2.ListIndex = -1 Then ' Se não tiver dados
        ' Mensagem de retorno vazio
        errorStyle.Informativo ESCOLHA_CHAPA_MENSAGEM, ESCOLHA_CHAPA_TITULO
        ' Sai do metodo
        Exit Sub
    End If
    
    ' Verifica se tem estoque a chapa selecionada
    If ListTamanhosChapaDespache.list(ListTamanhosChapaDespache.ListIndex, 0) < 1 Then
        ' Mensagem de retorno vazio
        errorStyle.Informativo SEM_ESTOQUE_MENSAGEM, SEM_ESTOQUE_TITULO
        ' Sai do metodo
        Exit Sub
    End If
    
    ' Verifica se tem algum tamanho selecionado
    If ListTamanhosChapaDespache.ListIndex = -1 Then ' Se não tiver dados
        ' Mensagem de retorno vazio
        errorStyle.Informativo SELECIONE_TAMANHO_MENSAGEM, SELECIONE_TAMANHO_TITULO
        ' Sai do metodo
        Exit Sub
    End If
    
    If obVenda.Value = False Then
        ' Verifica o motorista
        If cbMotorista.Value = "" Or cbMotorista.Value = " " Then
            ' Deixa visivel o erro com mensagens
            errorStyle.EntrarErrorStyleComboBox cbMotorista, SELECIONE_MOTORISTA_MENSAGEM, SELECIONE_MOTORISTA_TITULO
            ' Para o fluxo do sistema para a correção
            Exit Sub
        End If
        ' Deixa na cor patrão
        errorStyle.sairErrorStyleComboBox cbMotorista
        
        ' Verifica o destino
        If cbDestino.Value = "" Or cbDestino.Value = " " Then
            ' Deixa visivel o erro com mensagens
            errorStyle.EntrarErrorStyleComboBox cbDestino, SELECIONE_DESTINO_MENSAGEM, SELECIONE_DESTINO_TITULO
            ' Para o fluxo do sistema para a correção
            Exit Sub
        End If
        ' Deixa na cor patrão
        errorStyle.sairErrorStyleComboBox cbDestino
    End If
    
    ' Verifica se a data é valida
    If IsDate(txtDataDespacho.Value) = False Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtDataDespacho, SELECIONE_DATA_MENSAGEM, SELECIONE_DATA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtDataDespacho
    
    ' Verifica a quantidade para despache
    If txtQuantidadeDespache.Value = "0" Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtQuantidadeDespache, SELECIONE_QTD_MENSAGEM, SELECIONE_QTD_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtQuantidadeDespache
   
    ' NOME CABEÇALHO DESPACHE     | COD | DESCRIÇÃO | QTD   | M²
    ' Tamanho do cabeçalho left   | 7   | 118       | 346,5 | 399
    ' Tamanho do cabeçalho width  | 110 | 228       | 52    | 81
    ' Tamanho das colunas da list
    ListDespachado.ColumnWidths = "110;228;52;81;"
 
    ' Adiciona uma linha
    ListDespachado.AddItem
    
    ' Adiciona os dados par despache
    ListDespachado.list(ListDespachado.ListCount - 1, 0) = txtIdSistemaChapaDespache.Value
    ListDespachado.list(ListDespachado.ListCount - 1, 1) = txtMaterial.Value
    ListDespachado.list(ListDespachado.ListCount - 1, 2) = txtQuantidadeDespache.Value
    ListDespachado.list(ListDespachado.ListCount - 1, 3) = txtQtdm2Despache.Value
    ListDespachado.list(ListDespachado.ListCount - 1, 4) = ListTamanhosChapaDespache.list( _
                                ListTamanhosChapaDespache.ListIndex, 5) ' Id tamanho
    ListDespachado.list(ListDespachado.ListCount - 1, 5) = ListTamanhosChapaDespache.list( _
                                ListTamanhosChapaDespache.ListIndex, 1) ' Comprimento
    ListDespachado.list(ListDespachado.ListCount - 1, 6) = ListTamanhosChapaDespache.list( _
                                ListTamanhosChapaDespache.ListIndex, 2) ' Altura
    ListDespachado.list(ListDespachado.ListCount - 1, 7) = ListTamanhosChapaDespache.list( _
                                ListTamanhosChapaDespache.ListIndex, 3) ' Espesura
    ListDespachado.list(ListDespachado.ListCount - 1, 8) = ListTamanhosChapaDespache.list( _
                                ListTamanhosChapaDespache.ListIndex, 6) ' Estoque
    ListDespachado.list(ListDespachado.ListCount - 1, 9) = txtNumeroBloco.Value
                                
    ' Atualiza a lista do estoque na memoria
    ListEstoqueM2.list(ListEstoqueM2.ListIndex, 2) = ListEstoqueM2.list( _
                ListEstoqueM2.ListIndex, 2) - ListDespachado.list(ListDespachado.ListCount - 1, 2)
    ListEstoqueM2.list(ListEstoqueM2.ListIndex, 3) = Format(ListEstoqueM2.list( _
                ListEstoqueM2.ListIndex, 3) - ListDespachado.list(ListDespachado.ListCount - 1, 3), "0.0000")
    
    qtdChapas = CInt(ListDespachado.list(ListDespachado.ListCount - 1, 2))
    
    totalChapas = CInt(lQtdChapasDespache.Caption) + qtdChapas
    
    lQtdChapasDespache.Caption = totalChapas
    
    ' Limpa campos de tamanho da chapa
    Call limparCamposTamanho
    
    ' Seta focu
    txtPesquisarMaterial.SetFocus
End Sub

' Botão btnLTxtDespachar tela despache
Private Sub btnLTxtDespachar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim resposta As VbMsgBoxResult ' Variavel para confirmação impressão
    Dim tamanhoChapa As Collection
    Dim totalChapasDespachadas As Integer
    Dim venda As String
    Dim i As Integer
    
    ' Verificoes
    If IsDate(txtDataDespacho.Value) = False Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtDataDespacho, SELECIONE_DATA_MENSAGEM, SELECIONE_DATA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtDataDespacho
    
    ' Verifica se tem algum item na lista
    If Me.ListDespachado.ListCount = -1 Then
        ' Mensagem usuário
        errorStyle.Informativo LIST_DESPACHE_SEM_DADOS_MENSAGEM, LIST_DESPACHE_SEM_DADOS_TITULO
        Exit Sub
    End If
    
    If obVenda.Value = False Then
        ' Verifica o motorista
        If cbMotorista.Value = "" Or cbMotorista.Value = " " Then
            ' Deixa visivel o erro com mensagens
            errorStyle.EntrarErrorStyleComboBox cbMotorista, SELECIONE_MOTORISTA_MENSAGEM, SELECIONE_MOTORISTA_TITULO
            ' Para o fluxo do sistema para a correção
            Exit Sub
        End If
        ' Deixa na cor patrão
        errorStyle.sairErrorStyleComboBox cbMotorista
        
        ' Seta objeto
        Set motorista = daoMotorista.pesquisarPorNome(cbMotorista.Value)
        
        ' Verifica o destino
        If cbDestino.Value = "" Or cbDestino.Value = " " Then
            ' Deixa visivel o erro com mensagens
            errorStyle.EntrarErrorStyleComboBox cbDestino, SELECIONE_DESTINO_MENSAGEM, SELECIONE_DESTINO_TITULO
            ' Para o fluxo do sistema para a correção
            Exit Sub
        End If
        ' Deixa na cor patrão
        errorStyle.sairErrorStyleComboBox cbDestino
        
        ' Seta objeto
        Set destino = daoDestino.pesquisarPorNome(cbDestino.Value)
    Else
        ' Seta objeto de venda
        Set motorista = daoMotorista.pesquisarPorNome("VENDA")
        Set destino = daoDestino.pesquisarPorNome("VENDA")
    End If
    
    ' Criacao dos objetos
    Set despache = ObjectFactory.factoryDespache(despache)
    Set listaObjeto = ObjectFactory.factoryLista(listaObjeto)
    
    
    ' Coloca em lista as chapas para despache
    For i = 0 To ListDespachado.ListCount - 1
        
        Set chapa = ObjectFactory.factoryChapa(chapa)
        Set tamanho = ObjectFactory.factoryTamanho(tamanho)
        Set tipoPolimento = ObjectFactory.factoryTipoPolimento(tipoPolimento)
        Set tamanhoChapa = ObjectFactory.factoryLista(tamanhoChapa)
        Set estoqueChapa = daoEstoqueChapa.pesquisarPorNome(ListDespachado.list(i, 8))
        Set bloco = daoBloco.pesquisarPorId(ListDespachado.list(i, 9), True)
        
        chapa.idSistema = ListDespachado.list(i, 0)
        chapa.nomeMaterial = ListDespachado.list(i, 1)
        chapa.numeroBlocoPedreira = bloco.numeroBlocoPedreira
        chapa.setBloco bloco
        
        tamanho.id = ListDespachado.list(i, 4)
        tamanho.qtdEstoque = ListDespachado.list(i, 2)
        tamanho.qtdM2 = ListDespachado.list(i, 3)
        tamanho.compremento = ListDespachado.list(i, 5)
        tamanho.altura = ListDespachado.list(i, 6)
        tamanho.espessura = ListDespachado.list(i, 7)
        tamanho.setEstoque estoqueChapa
        
        tamanhoChapa.Add tamanho
        chapa.setTamanhos tamanhoChapa
        
        listaObjeto.Add chapa
        
        Set chapa = Nothing
        Set tipoPolimento = Nothing
        Set tamanho = Nothing
        Set tamanhoChapa = Nothing
        Set estoqueChapa = Nothing
        Set bloco = Nothing
    Next i
    
    totalChapasDespachadas = CInt(lQtdChapasDespache.Caption)
    
    despache.carregarDespacheCadastro txtDataDespacho.Value, "SIM", totalChapasDespachadas, motorista, destino, listaObjeto

    ' Chama medoto para salvar no banco
    Call DaoDespache.cadastrarEEditar(despache)
    
    ' Mensagem de confirmação para gerar pdf da carga
    resposta = MsgBox(IMPRIMIR_CARREGO_MENSAGEM, vbQuestion + vbYesNo, IMPRIMIR_CARREGO_TITULO)
    
    
    ' Verifica a confirmação do usário para poder cadastrar
    If resposta = vbYes Then
        ' Gera um PDf
        Call ExportarArquivos.exportarCarregoPDF(despache)

    End If
    
    ' MENSAGEM DE SUCESSO
    
    ' Mensagem usuário
    errorStyle.Informativo SUCESSO_DESPACHE_MENSAGEM, SUCESSO_DESPACHE_TITULO
    
    ' Limpa os campos
    Call limparCamposTelaDespache
    
    Set despache = Nothing
    Set listaObjeto = Nothing
End Sub

' Botão btnLTxtSalvarDespache tela despache
Private Sub btnLTxtSalvarDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Dim resposta As VbMsgBoxResult ' Variavel para confirmação impressão
    Dim tamanhoChapa As Collection
    Dim totalChapasDespachadas As Integer
    Dim venda As String
    Dim i As Integer
    
    ' Verificoes
    If IsDate(txtDataDespacho.Value) = False Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtDataDespacho, SELECIONE_DATA_MENSAGEM, SELECIONE_DATA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtDataDespacho
    
    ' Verifica se tem algum item na lista
    If Me.ListDespachado.ListCount = -1 Then
        ' Mensagem usuário
        errorStyle.Informativo LIST_DESPACHE_SEM_DADOS_MENSAGEM, LIST_DESPACHE_SEM_DADOS_TITULO
        Exit Sub
    End If
    
    If obVenda.Value = False Then
        ' Verifica o motorista
        If cbMotorista.Value = "" Or cbMotorista.Value = " " Then
            ' Deixa visivel o erro com mensagens
            errorStyle.EntrarErrorStyleComboBox cbMotorista, SELECIONE_MOTORISTA_MENSAGEM, SELECIONE_MOTORISTA_TITULO
            ' Para o fluxo do sistema para a correção
            Exit Sub
        End If
        ' Deixa na cor patrão
        errorStyle.sairErrorStyleComboBox cbMotorista
        
        ' Seta objeto
        Set motorista = daoMotorista.pesquisarPorNome(cbMotorista.Value)
        
        ' Verifica o destino
        If cbDestino.Value = "" Or cbDestino.Value = " " Then
            ' Deixa visivel o erro com mensagens
            errorStyle.EntrarErrorStyleComboBox cbDestino, SELECIONE_DESTINO_MENSAGEM, SELECIONE_DESTINO_TITULO
            ' Para o fluxo do sistema para a correção
            Exit Sub
        End If
        ' Deixa na cor patrão
        errorStyle.sairErrorStyleComboBox cbDestino
        
        ' Seta objeto
        Set destino = daoDestino.pesquisarPorNome(cbDestino.Value)
    Else
        ' Seta objeto de venda
        Set motorista = daoMotorista.pesquisarPorNome("VENDA")
        Set destino = daoDestino.pesquisarPorNome("VENDA")
    End If
    
    ' Criacao dos objetos
    Set despache = ObjectFactory.factoryDespache(despache)
    Set listaObjeto = ObjectFactory.factoryLista(listaObjeto)
    
    
    ' Coloca em lista as chapas para despache
    For i = 0 To ListDespachado.ListCount - 1
        
        Set chapa = ObjectFactory.factoryChapa(chapa)
        Set tamanho = ObjectFactory.factoryTamanho(tamanho)
        Set tipoPolimento = ObjectFactory.factoryTipoPolimento(tipoPolimento)
        Set tamanhoChapa = ObjectFactory.factoryLista(tamanhoChapa)
        Set estoqueChapa = daoEstoqueChapa.pesquisarPorNome(ListDespachado.list(i, 8))
        Set bloco = daoBloco.pesquisarPorId(ListDespachado.list(i, 9), True)
        
        chapa.idSistema = ListDespachado.list(i, 0)
        chapa.nomeMaterial = ListDespachado.list(i, 1)
        chapa.numeroBlocoPedreira = bloco.numeroBlocoPedreira
        chapa.setBloco bloco
        
        tamanho.id = ListDespachado.list(i, 4)
        tamanho.qtdEstoque = ListDespachado.list(i, 2)
        tamanho.qtdM2 = ListDespachado.list(i, 3)
        tamanho.compremento = ListDespachado.list(i, 5)
        tamanho.altura = ListDespachado.list(i, 6)
        tamanho.espessura = ListDespachado.list(i, 7)
        tamanho.setEstoque estoqueChapa
        
        tamanhoChapa.Add tamanho
        chapa.setTamanhos tamanhoChapa
        
        listaObjeto.Add chapa
        
        Set chapa = Nothing
        Set tipoPolimento = Nothing
        Set tamanho = Nothing
        Set tamanhoChapa = Nothing
        Set estoqueChapa = Nothing
        Set bloco = Nothing
    Next i
    
    totalChapasDespachadas = CInt(lQtdChapasDespache.Caption)
    
    despache.carregarDespacheCadastro txtDataDespacho.Value, "NÃO", totalChapasDespachadas, motorista, destino, listaObjeto

    ' Chama medoto para salvar no banco
    Call DaoDespache.cadastrarEEditar(despache)
    
    ' Mensagem de confirmação para gerar pdf da carga
    resposta = MsgBox(IMPRIMIR_CARREGO_MENSAGEM, vbQuestion + vbYesNo, IMPRIMIR_CARREGO_TITULO)
    
    
    ' Verifica a confirmação do usário para poder cadastrar
    If resposta = vbYes Then
        ' Gera um PDf
        Call ExportarArquivos.exportarCarregoPDF(despache)

    End If
    
    ' MENSAGEM DE SUCESSO
    
    ' Mensagem usuário
    errorStyle.Informativo SUCESSO_DESPACHE_MENSAGEM, SUCESSO_DESPACHE_TITULO
    
    ' Limpa os campos
    Call limparCamposTelaDespache
    
    Set despache = Nothing
    Set listaObjeto = Nothing
End Sub

' Botão btnLTxtTirarListaDespache tela despache
Private Sub btnLTxtTirarListaDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Variaveis do metodo
    Dim linha As Double
    Dim totalChapas As Integer
    Dim i As Integer
    Dim qtdDespache As Integer
    Dim qtdEstoque As Integer
    Dim totalQtd As Integer
    Dim qtdM2Despache As Double
    Dim qtdM2Estoque As Double
    Dim totalM2 As Double
    Dim totalM2String As String
    
    ' Verifica se tem algum item selecionado
    If Me.ListDespachado.ListIndex = -1 Then
        ' Mensagem usuário
        errorStyle.Informativo ESCOLHA_CHAPA_MENSAGEM, ESCOLHA_CHAPA_TITULO
        Exit Sub
    End If
    
    ' Captura a linha clicada
    linha = ListDespachado.ListIndex
    
    ' Analisa se tem item na lista de despache
    If ListEstoqueM2.ListCount > 0 Then
        ' Percorre a lista
        For i = 0 To ListEstoqueM2.ListCount - 1
            ' Compara se já tem o material
            If ListEstoqueM2.list(i, 0) = ListDespachado.list(ListDespachado.ListIndex, 0) Then
                ' Atualiza quantidade na lista
                qtdDespache = CInt(ListDespachado.list(ListDespachado.ListIndex, 2))
                qtdEstoque = CInt(ListEstoqueM2.list(i, 2))
                totalM2 = qtdDespache + qtdEstoque
                
                ListEstoqueM2.list(i, 2) = totalM2
                
                ' Atualiza m² na lista
                qtdM2Despache = CDbl(ListDespachado.list(ListDespachado.ListIndex, 3))
                qtdM2Estoque = CDbl(ListEstoqueM2.list(i, 3))
                totalM2 = qtdM2Despache + qtdM2Estoque
                totalM2String = CStr(totalM2)
                
                ListEstoqueM2.list(i, 3) = Format(totalM2String, "0.0000")

                ' Atualiza label com total de chapas a serem carregadas
                totalChapas = CInt(lQtdChapasDespache.Caption) - qtdDespache
                
                lQtdChapasDespache.Caption = totalChapas
            End If
        Next i
    End If
    ' Remove o item selecionado
    ListDespachado.RemoveItem linha
End Sub

' Botão btnLTxtLimparDespache tela despache
Private Sub btnLTxtLimparDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Limpa os campos
    Call limparCamposTelaDespache
End Sub

'-----------------------------------------------------------------TELA CARREGOS-----------------------------------
'                                                                 -------------
' Botão btnLTxtPesquisarCarregos tela carregos
Private Sub btnLTxtPesquisarCarregos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço pesquisar por carregos, tela carregos"
    ' Seta o foco
    cbMotoristaL.SetFocus
End Sub
' Botão btnLTxtLimparListas tela carregos
Private Sub btnLTxtLimparListas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço limpar dados filtro, tela carregos"
'    ' Seta o foco
'    cbMotoristaL.SetFocus
End Sub
' Botão btnLImgExportarCarregoPDF tela carregos
Private Sub btnLImgExportarCarregoPDF_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço exportar carregos em pdf, tela carregos"
    ' Seta o foco
'    cbMotoristaL.SetFocus
End Sub
' Botão btnLTxtEditarCarrego tela carregos
Private Sub btnLTxtEditarCarrego_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço editar carrego, tela carregos"
    ' Seta o foco
    cbMotoristaL.SetFocus
End Sub
' Botão btnLTxtVoltarCArrego tela carregos
Private Sub btnLTxtVoltarCArrego_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço voltar, tela carregos"
    ' Seta o foco
    cbMotoristaL.SetFocus
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

'-----------------------------------------------------------------CLIQUE LIST-------------------------------------------------------
'                                                                 -----------
' Clique na ListTamanhosChapaDespache na tela despache
Private Sub ListTamanhosChapaDespache_Click()
    ' Variaveis do metodo
    Dim linha As Integer
    Dim qtdEstoque As Integer
    Dim qtdSaida As Integer
    
    ' Captura a linha selecionada
    linha = ListTamanhosChapaDespache.ListIndex

    ' Verifica se foi selecionado tamanho da chapa
    If ListTamanhosChapaDespache.ListIndex <> -1 Then
        ' Verifica a quantidade
        If txtQuantidadeDespache.Value = "0" Then
            ' Mensagem usuário
            errorStyle.EntrarErrorStyleTextBox txtQuantidadeDespache, SELECIONE_QTD_MENSAGEM, SELECIONE_QTD_TITULO
            Exit Sub
        End If
        ' Deixa na cor patrão
        errorStyle.sairErrorStyleTextBox txtQuantidadeDespache
        
        qtdEstoque = CInt(ListTamanhosChapaDespache.list(linha, 0))
        qtdSaida = CInt(txtQuantidadeDespache.Value)
        
        ' Verifica o estoque
        If qtdSaida > qtdEstoque Then
            ' Mensagem usuário
            errorStyle.EntrarErrorStyleTextBox txtQuantidadeDespache, SEM_ESTOQUE_MENSAGEM, SEM_ESTOQUE_TITULO
            Exit Sub
        End If
        ' Deixa na cor patrão
        errorStyle.sairErrorStyleTextBox txtQuantidadeDespache
        
        ' Captura a linha selecionada
        linha = ListTamanhosChapaDespache.ListIndex
        
        ' Retorna valor calculado e formatado
        txtQtdm2Despache.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularM2( _
            ListTamanhosChapaDespache.list(linha, 1), ListTamanhosChapaDespache.list(linha, 2), txtQuantidadeDespache.Value), "0.0000"))
    End If
End Sub

' Clique na ListEstoqueM2 na tela despache
Private Sub ListEstoqueM2_Click()
    ' Variaveis do metodo
    Dim estoqueDespache As Integer
    Dim linha As Integer
    Dim i As Integer
    ' Captura linha selecionada
    linha = ListEstoqueM2.ListIndex

    ' Pesquisa pela chapa
    Set chapa = daoChapa.pesquisarPorId(ListEstoqueM2.list(linha, 0))
    
    ' Soma total disponivel no estoque
    For i = 1 To chapa.tamanhos.Count
        Set tamanho = chapa.tamanhos.Item(i)
        ' Soma
        estoqueDespache = estoqueDespache + CInt(tamanho.qtdEstoque)
    Next i
    
    ' Seta valores
    txtMaterial.Value = ListEstoqueM2.list(linha, 1)
    txtIdSistemaChapaDespache.Value = ListEstoqueM2.list(linha, 0)
    txtQuantidedadeChapa.Value = estoqueDespache
    txtNumeroBloco.Value = ListEstoqueM2.list(linha, 4)
    
    ' Seta tamanhos da chapa pesquisada
    Call carregarTamanhosChapasTelaDespache(ListTamanhosChapaDespache, chapa.tamanhos)
    
    Set chapa = Nothing
    
    ' Seta foco para ser adicionada a quantiidade
    txtQuantidadeDespache.SetFocus
End Sub

'-----------------------------------------------------------------PESQUISAR-------------------------------------------------------
'                                                                 ---------
' Pesquisa chapas tela despachar por descrição
Private Sub pesquisarChapaDespachePorDescricao()
    ' Se campo estiver em branco, apenas sai do medoto
    If txtPesquisarMaterial.Value = "" Or txtPesquisarMaterial.Value = " " Then
        Exit Sub
    Else ' Se foi digitado algo se faz a pesquisa no banco de dados
        
        ' Pesquisa no banco de dados
        Set listaObjeto = daoChapa.listarChapasFilter(txtPesquisarMaterial.Value, "", "", "", "", "NÃO")
        
        ' Verifica se tem algum dado a pesquisa
        If listaObjeto.Count = -1 Or listaObjeto.Count = 0 Then ' Se não tiver dados
            ' Mensagem de retorno vazio
            errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
            ' Sai do metodo
            Exit Sub
        End If
        ' Carrega os dados
        Call carregarListsTelaDespache(ListEstoqueM2, listaObjeto)
    End If
End Sub

' Pesquisa chapas tela despachar por número bloco
Private Sub pesquisarChapaDespachePorNumeroBloco()
    ' Se campo estiver em branco, apenas sai do medoto
    If txtPesquisarPorNumeroBloco.Value = "" Or txtPesquisarPorNumeroBloco.Value = " " Then
        Exit Sub
    Else ' Se foi digitado algo se faz a pesquisa no banco de dados
        
        ' Pesquisa no banco de dados
        Set listaObjeto = daoChapa.listarChapasFilter("", txtPesquisarPorNumeroBloco.Value, "", "", "", "NÃO")
        
        ' Verifica se tem algum dado a pesquisa
        If listaObjeto.Count = -1 Or listaObjeto.Count = 0 Then ' Se não tiver dados
            ' Mensagem de retorno vazio
            errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
            ' Sai do metodo
            Exit Sub
        End If
        ' Carrega os dados
        Call carregarListsTelaDespache(ListEstoqueM2, listaObjeto)
    End If
End Sub

' Pesquisa lista chapas salva para despache tela despachar
Private Sub pesquisarDespacheSalvo()

End Sub

' Pesquisa blocos com filtros tela estoque m³
Private Sub pesquisarChapasFilter()
    ' Variaveis do metodo
    Dim listaChapas As Collection
    Dim estoqueZero As String
    
    ' Atribuição das variaveies
    If obEstoqueZeroNao.Value = True Then
        estoqueZero = "NÃO"
    Else
        estoqueZero = "SIM"
    End If
    
    ' Faz pesquisa com filtros no banco de dados e retorna uma lista
    Set listaChapas = daoChapa.listarChapasFilter(txtMaterialChapaPesquisa.Value, txtIdBlocoChapaPesquisa.Value, _
                        txtIdchapaEstoque.Value, cbPolideiraChapaPesquisa.Value, cbTipoPolimentoPesquisa.Value, estoqueZero)
            
    ' Carrega a lista
    Call carregarList(ListEstoqueChapas, listaChapas)
    
    ' Libera espeço na memoria
    Set listaChapas = Nothing
End Sub

' Pesquisa blocos com filtros tela estoque m³
Private Sub pesquisarBlocosFilter()
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
    
    ' Validando a data
    If IsDate(txtDataInicioBlocoPesquisa.Value) = False Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtDataInicioBlocoPesquisa, ADICIONE_DATA_MENSAGEM, ADICIONE_DATA_TITULO
        'Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtDataInicioBlocoPesquisa

    ' Validando a data
    If IsDate(txtDataFinalBlocoPesquisa.Value) = False Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtDataFinalBlocoPesquisa, ADICIONE_DATA_MENSAGEM, ADICIONE_DATA_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtDataFinalBlocoPesquisa
    
    ' Atribuição das variaveies
    dataInicial = txtDataInicioBlocoPesquisa.Value
    dataFinal = txtDataFinalBlocoPesquisa.Value
    idBlocoPedreira = txtIdBlocoPesquisa.Value
    descricaoBloco = txtMaterialBlocoPesquisa.Value
    pedreiraBloco = cbPedreiraBlocoPesquisa.Value
    serrariaBloco = cbSerrariaBlocoPesquisa.Value
    temNota = cbTemNota.Value
    
    ' Status filter
    statusPedreira = ""
    statusSerraria = ""
    statusChapasBrutas = ""
    statusEmProcesso = ""
    statusEstoque = ""
    statusFechado = ""
    
    ' Status para pesquisa e formatação
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
            
    ' Mensagem para o usuario escolher algum Status
    If chbPedreida.Value = False And chbSerraria.Value = False And chbChapasBrutas.Value = False _
            And chbEmProcesso.Value = False And chbEstoque.Value = False And chbFechado.Value = False Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleOptionButton obPedreiraESerrada, ADICIONE_STATUS_MENSAGEM, ADICIONE_STATUS_TITULO
        ' Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleOptionButton obPedreiraESerrada
    
    ' Faz pesquisa com filtros no banco de dados e retoeno uma lista
    Set listaBlocos = daoBloco.listarBlocosFilter(dataInicial, dataFinal, idBlocoPedreira, _
            descricaoBloco, pedreiraBloco, serrariaBloco, temNota, statusPedreira, statusSerraria, _
            statusChapasBrutas, statusEmProcesso, statusEstoque, statusFechado)
            
    ' Carrega a lista
    Call carregarList(ListEstoqueM3, listaBlocos)
    
    ' Libera espeço na memoria
    Set listaBlocos = Nothing
End Sub

'-----------------------------------------------------------------DESABILITA  E HABILITA CAMPOS-----------------------------------
'                                                                 -----------------------------
' Desabilita campos da tela lançamento e edição de chapas
Private Sub desabilitaCamposChapas()
'    txtIdChapaSistema.Enabled = False
'    txtDescricaoChapa.Enabled = False
    txtEstoqueChapa.Enabled = False
    txtIdBlocoPedreiraChapa.Enabled = False
    txtDecricaoBlocoChapa.Enabled = False
    txtQtdDisponivelChapaBloco.Enabled = False
    txtNBlocoPedreiraChapa.Enabled = False
    txtTipoMaterialChapa.Enabled = False
    cbPolideiraChapa.Enabled = False
    cbTipoPolimentoChapa.Enabled = False
    cbTipoMaterialChapaC.Enabled = False
    cbEstoqueChapaC.Enabled = False
    txtCompLiquidoChapa.Enabled = False
    txtAlturaLiquidoChapa.Enabled = False
    txtQtdChapaC.Enabled = False
    txtCompBrutoChapa.Enabled = False
    txtAlturaBrutaChapa.Enabled = False
    txtEspChapa.Enabled = False
    txtQtsM2Chapa.Enabled = False
    txtCustoChapa.Enabled = False
    txtTotalChapas.Enabled = False
    cbTiposMateriaisChapas.Enabled = False
    txtCompTipoMateriaisChapa.Enabled = False
    txtAltTipoMateriaisChapa.Enabled = False
    txtQtdM2TipoMateriaisChapas.Enabled = False
    txtQtdTipoMateriaisChapas.Enabled = False
    txtEspTiposMateriaisChapa.Enabled = False
End Sub

' Habilita campos da tela lançamento e edição de chapas
Private Sub habilitaCamposChapas()
'    txtIdChapaSistema.Enabled = False
'    txtDescricaoChapa.Enabled = True
    txtEstoqueChapa.Enabled = True
    txtIdBlocoPedreiraChapa.Enabled = False
    txtDecricaoBlocoChapa.Enabled = False
    txtQtdDisponivelChapaBloco.Enabled = False
    txtNBlocoPedreiraChapa.Enabled = False
    txtTipoMaterialChapa.Enabled = False
    cbPolideiraChapa.Enabled = True
    cbTipoPolimentoChapa.Enabled = True
    cbTipoMaterialChapaC.Enabled = True
    cbEstoqueChapaC.Enabled = True
'    txtCompLiquidoChapa.Enabled = True
'    txtAlturaLiquidoChapa.Enabled = True
'    txtQtdChapaC.Enabled = True
    txtCompBrutoChapa.Enabled = True
    txtAlturaBrutaChapa.Enabled = True
'    txtEspChapa.Enabled = True
    txtQtsM2Chapa.Enabled = True
    txtCustoChapa.Enabled = True
    txtTotalChapas.Enabled = True
'    cbTiposMateriaisChapas.Enabled = True
    txtCompTipoMateriaisChapa.Enabled = True
    txtAltTipoMateriaisChapa.Enabled = True
    txtQtdM2TipoMateriaisChapas.Enabled = True
    txtQtdTipoMateriaisChapas.Enabled = True
    txtEspTiposMateriaisChapa.Enabled = True
End Sub
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
    
    ' Dimensões bloco e médias chapas
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
    txtIdBlocoEditar.Enabled = False
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
    
    ' Dimensões bloco e médias chapas
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
    txtCustoMaterialBlocoEditar.Enabled = True
    txtTotalM2PolimentoBlocoEditar.Enabled = True
    txtTotalBlocoEditar.Enabled = True
End Sub

'-----------------------------------------------------------------LIMPAR CAMPOS-----------------------------------
'                                                                 -------------
' Limpa os campos da tela despache
Private Sub limparCamposTelaDespache()
    ' Limpa os campos
    lQtdChapasDespache.Caption = "0"
    txtNumeroBloco.Value = ""
    txtPesquisarMaterial.Value = ""
    txtPesquisarPorNumeroBloco.Value = ""
    txtPesquisarDespache.Value = ""
    txtMaterial.Value = ""
    txtIdSistemaChapaDespache.Value = ""
    txtQuantidedadeChapa.Value = "0"
    txtQtdm2Despache.Value = "0,0000"
    txtQuantidadeDespache.Value = "0"
    ListDespachado.Clear
    ListEstoqueM2.Clear
    cbMotorista.Clear
    cbDestino.Clear
    ListTamanhosChapaDespache.Clear
End Sub

' Limpa os campos de chapa da tela cadastroEdição de chapa
Private Sub limparCamposChapa()
    txtIdChapaSistema.Value = ""
    txtDescricaoChapa.Value = ""
    cbTipoPolimentoChapa.Value = "POLIDO"
    txtCompBrutoChapa.Value = "0,000"
    txtAlturaBrutaChapa.Value = "0,000"
    txtQtsM2Chapa.Value = "0,000"
    txtIdBlocoPedreiraChapa.Value = ""
    txtDecricaoBlocoChapa.Value = ""
    txtQtdDisponivelChapaBloco.Value = ""
    txtTipoMaterialChapa.Value = ""
    txtNBlocoPedreiraChapa.Value = ""
End Sub
' Limpa os campos de tamanho da tela cadastroEdição de chapa
Private Sub limparCamposTamanhoChapa()
    cbTipoMaterialChapaC.Value = "EXTRA"
    cbPolideiraChapa.Value = ""
    cbEstoqueChapaC.Value = "CASA DO GRANITO"
    txtCompTipoMateriaisChapa.Value = "0,0000"
    txtAltTipoMateriaisChapa.Value = "0,0000"
    txtQtdM2TipoMateriaisChapas.Value = "0,0000"
    txtEspTiposMateriaisChapa.Value = "02"
    txtQtdTipoMateriaisChapas.Value = "0"
    txtCustoChapa.Value = "0,00"
    txtTotalChapas.Value = "0,00"
End Sub
' Limpa os campos de pesquisa da tela estoque M³
Private Sub limparCamposTrocaEstoque()
    txtMaterialParaTroca01.Value = ""
    txtEspParaTroca01.Value = ""
    txtTipoMaterialParaTroca01.Value = ""
    txtCompMaterialParaTroca.Value = ""
    txtAltMaterialParaTroca.Value = ""
    txtTotalM2T.Value = ""
    txtQtdDispovelMaterialParaTroca01.Value = ""
    txtQtdMaterialParaTroca02.Value = 0
    cbTipoPolimentoTroca.Clear
    ListMateriaisParaTroca.Clear
    ListTrocarPor.Clear
    
End Sub
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
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim i As Integer
    
    ' Criando a lista
    Set listaObjetos = daoPedreira.listarPedreiras

    ' limpa a lista para carregamento
    cbPedreiras.Clear
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se não tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop através dos itens da coleção
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set pedreira = listaObjetos(i)
            ' Carregamento para lista
            cbPedreiras.AddItem pedreira.nome
            ' Libera espaço memoria
            Set pedreira = Nothing
        Next i
    End If
    ' Libera espaço da memoria
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
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se não tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
       ' Loop através dos itens da coleção
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set serraria = listaObjetos(i)
            ' Carregamento para lista
            cbSerrarias.AddItem serraria.nome
            ' Libera espaço memoria
            Set serraria = Nothing
        Next i
    End If
    ' Libera espaço da memoria
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
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se não tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop através dos itens da coleção
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set tipoMaterial = listaObjetos(i)
            ' Carregamento para lista
            cbTiposMateriais.AddItem tipoMaterial.nome
            ' Libera espaço memoria
            Set tipoMaterial = Nothing
        Next i
        
        ' Deixar um item já selecionado
        If cbTiposMateriais.name = "cbTipoMaterial" Then
            Call selecaoItem("cbTipoMaterial", "EXTRA")
            
        ElseIf cbTiposMateriais.name = "cbTipoMaterialL" Then
            Call selecaoItem("cbTipoMaterialL", "EXTRA")
        End If
        
    End If
    ' Libera espaço da memoria
    Set listaObjetos = Nothing
End Sub

' Carrega a combobox tem nota
Private Sub carregarTemNota(cbTemNota As MSForms.comboBox)
    ' limpa a lista para carregamento
    cbTemNota.Clear
    
    ' Deixar um item já selecionado
    If Me.MultiPageCEBC.Value = 1 Then
        cbTemNota.AddItem ""
    End If
    
    ' Carregamento para lista
    cbTemNota.AddItem "SIM"
    cbTemNota.AddItem "NÃO"
    
    ' Deixar um item já selecionado
    If cbTemNota.name = "cbNotaC" Then
        Call selecaoItem("cbNotaC", "NÃO")
    
    ElseIf cbTemNota.name = "cbTemNotaAvulso" Then
        Call selecaoItem("cbTemNotaAvulso", "NÃO")
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
    cbPolideiras.AddItem ""
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se não tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop através dos itens da coleção
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set polideira = listaObjetos(i)
            ' Carregamento para lista
            cbPolideiras.AddItem polideira.nome
            ' Libera espaço memoria
            Set polideira = Nothing
        Next i
    End If
    ' Libera espaço da memoria
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
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se não tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop através dos itens da coleção
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set tipoPolimento = listaObjetos(i)
            ' Carregamento para lista
            cbTiposPolimento.AddItem tipoPolimento.nome
            ' Libera espaço memoria
            Set tipoPolimento = Nothing
        Next i
        
        ' Deixar um item já selecionado
        If cbTiposPolimento.name = "cbTipoPolimentoL" Then
            Call selecaoItem("cbTipoPolimentoL", "POLIDO")
        End If
        
    End If
    ' Libera espaço da memoria
    Set listaObjetos = Nothing
End Sub

' Carrega a combobox de tipo polimento com algum tipos
Private Sub carregarTiposPolimentoAlgum(cbTiposPolimento As MSForms.comboBox, lista As Collection)
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim polimento As Variant
    Dim totalLista As Integer
    Dim temNaLista As Boolean
    Dim i As Integer
    Dim j As Integer
    
    ' Criando a lista
    Set listaObjetos = daoTipoPolimento.listarTipoPolideiras
    ' Inicia com false
    temNaLista = False
    ' limpa a lista para carregamento
    cbTiposPolimento.Clear
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se não tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop através dos itens da coleção
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set tipoPolimento = listaObjetos(i)
            ' Laço nos polimentos já cadastrados
            For j = 1 To lista.Count
                ' Seta para comparação
                polimento = lista(j)
                
                ' Compara se tem na lista
                If tipoPolimento.nome = polimento Then
                    temNaLista = True
                    Exit For
                End If
            Next j
            ' Se tiver na lista adiciona no combox
            If temNaLista = False Then
                ' Carregamento para lista
                cbTiposPolimento.AddItem tipoPolimento.nome
                ' Volta com false para proxima verificação
                temNaLista = False
                
            Else
                temNaLista = False
            End If
            
            ' Libera espaço memoria
            Set tipoPolimento = Nothing
        Next i
    End If
    ' Libera espaço da memoria
    Set listaObjetos = Nothing
End Sub

' Carrega a combobox de estoque tela edição de bloco
Private Sub carregarEstoque(cbTiposEstoque As MSForms.comboBox)
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim i As Integer
    
    ' Criando a lista
    Set listaObjetos = daoEstoqueM3.listarEstoqueM3

    ' limpa a lista para carregamento
    cbTiposEstoque.Clear
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se não tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop através dos itens da coleção
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set estoque = listaObjetos(i)
            ' Carregamento para lista
            cbTiposEstoque.AddItem estoque.nome
            ' Libera espaço memoria
            Set estoque = Nothing
        Next i
    End If
    ' Libera espaço da memoria
    Set listaObjetos = Nothing
    Set estoque = Nothing
End Sub

' Carrega a combobox de estoque tela chapa
Private Sub carregarEstoqueChapas(cbTiposEstoque As MSForms.comboBox)
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim i As Integer
    
    ' Criando a lista
    Set listaObjetos = daoEstoqueChapa.listarEstoqueChapas

    ' limpa a lista para carregamento
    cbTiposEstoque.Clear
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se não tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop através dos itens da coleção
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set estoqueChapa = listaObjetos(i)
            ' Carregamento para lista
            cbTiposEstoque.AddItem estoqueChapa.nome
            ' Libera espaço memoria
            Set estoqueChapa = Nothing
        Next i
        
    End If
    ' Libera espaço da memoria
    Set listaObjetos = Nothing
    Set estoqueChapa = Nothing
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
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim i As Integer
    
    ' Criando a lista
    Set listaObjetos = daoStatus.listarStatus

    ' limpa a lista para carregamento
    cbStatus.Clear
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se não tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop através dos itens da coleção
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set statusObj = listaObjetos(i)
            ' Carregamento para lista
            cbStatus.AddItem statusObj.nome
            ' Libera espaço memoria
            Set statusObj = Nothing
        Next i
    End If
    ' Libera espaço da memoria
    Set listaObjetos = Nothing
End Sub

' Carrega a combobox de motoristas
Private Sub carregarMotoristas(cbMotorista As MSForms.comboBox)
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim i As Integer
    
    ' Criando a lista
    Set listaObjetos = daoMotorista.listarMotoristas

    ' limpa a lista para carregamento
    cbMotorista.Clear
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se não tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop através dos itens da coleção
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set motorista = listaObjetos(i)
            ' Carregamento para lista
            cbMotorista.AddItem motorista.nome
            ' Libera espaço memoria
            Set motorista = Nothing
        Next i
    End If
    ' Libera espaço da memoria
    Set listaObjetos = Nothing
End Sub

' Carrega a combobox de destinos
Private Sub carregarDestinos(cbDestino As MSForms.comboBox)
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim i As Integer
    
    ' Criando a lista
    Set listaObjetos = daoDestino.listarDestinos

    ' limpa a lista para carregamento
    cbDestino.Clear
    
    ' Verifica se tem algum dado a pesquisa
    If listaObjetos.Count = -1 Or listaObjetos.Count = 0 Then ' Se não tiver dados
        ' Mensagem de erro
        errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        Exit Sub
    Else
        ' Loop através dos itens da coleção
        For i = 1 To listaObjetos.Count
            ' Seta o ojeto
            Set destino = listaObjetos(i)
            ' Carregamento para lista
            cbDestino.AddItem destino.nome
            ' Libera espaço memoria
            Set destino = Nothing
        Next i
    End If
    ' Libera espaço da memoria
    Set listaObjetos = Nothing
End Sub


'-----------------------------------------------------------------CARREAGMENTO DAS LIST-----------------------------------
'                                                                 ---------------------
' Carrega a lista bloco
Private Sub carregarList(ListBox As MSForms.ListBox, listaCollection As Collection)
   'Variaveis do metodo
    Dim objeto As objBloco
    Dim objetoChapa As objChapa
    Dim objetoTamanho As objTamanho
    Dim i As Integer
    Dim j As Integer
    Dim qtdChapas As Integer
    Dim qtdEstoque As Integer
    Dim mediaComp As Double
    Dim mediaAlt As Double
    Dim totalM2 As Double
    Dim valorPolimento As String
    Dim esp As String
    
    ' Limpar a ListBox
    ListBox.Clear
    
    ' NOME CABEÇALHO CHAPAS       | COD | DECRCIÇÃO | QTD  | COMP  | ALT  | M²    | TIPO     | ESP | VALOR | TOTAL |
    ' NOME CABEÇALHO BLOCOS       | COD | DECRCIÇÃO | COMP | ALT   | LARG | QTD   | VALOR M³ | ADD | FRETE | TOTAL |
    ' Tamanho do cabeçalho left   | 7   | 193       | 444  | 496,5 | 549  | 601,5 | 654      | 745 | 820,5 | 896   |
    ' Tamanho do cabeçalho width  | 185 | 250       | 52   | 52    | 52   | 52    | 90       | 75  | 75    | 74,5  |
    ' Tamanho das colunas da list
    ListBox.ColumnWidths = "185;250;52;52;52;52;90;75;75;74;"
    
    ' Verifica se tem algum dado a pesquisa
    If listaCollection.Count = -1 Or listaCollection.Count = 0 Then ' Se não tiver dados
        If paginaAnterior = 1 Or paginaAnterior = 4 Then
            Exit Sub
        ElseIf paginaAnterior <> 1 Then ' Ativa mensagem se a pagina anterior não for a do menu
            ' Mensagem de retorno
            errorStyle.Informativo SEM_DADOS_MENSAGEM, SEM_DADOS_TITULO
        End If
    Else
        ' Direciona lista
        If ListBox.name = "ListMateriais" Or ListBox.name = "ListEstoqueChapas" Then
            ' Loop através dos itens da coleção
            For i = 1 To listaCollection.Count
                ' Seta o ojeto
                Set objetoChapa = listaCollection(i)
                
                ' Adiciona uma linha
                ListBox.AddItem
                
                ' Adiciona os dados do bloco
                ListBox.list(ListBox.ListCount - 1, 0) = objetoChapa.idSistema
                ListBox.list(ListBox.ListCount - 1, 1) = objetoChapa.nomeMaterial
                
                For j = 1 To objetoChapa.tamanhos.Count
                    Set tamanho = objetoChapa.tamanhos.Item(j)
                    
                    qtdEstoque = qtdEstoque + CInt(tamanho.qtdEstoque)
                    mediaComp = mediaComp + CDbl(tamanho.compremento)
                    mediaAlt = mediaAlt + CDbl(tamanho.altura)
                    totalM2 = totalM2 + CDbl(tamanho.qtdM2)
                    valorPolimento = tamanho.valorPolimento
                    esp = tamanho.espessura
                    
                    ' Soma total de chapa da pesquisa
                    qtdChapas = qtdChapas + CInt(tamanho.qtdEstoque)
                    
                    ' Libera espaço da memoria
                    Set objetoTamanho = Nothing
                Next j
                
                ListBox.list(ListBox.ListCount - 1, 2) = qtdEstoque
                
                ' Media comprimento
                mediaComp = mediaComp / objetoChapa.tamanhos.Count
                ListBox.list(ListBox.ListCount - 1, 3) = _
                                        M_METODOS_GLOBAL.formatarComPontos(Format(mediaComp, "0.0000"))
                
                ' Media altura
                mediaAlt = mediaAlt / objetoChapa.tamanhos.Count
                ListBox.list(ListBox.ListCount - 1, 4) = _
                                        M_METODOS_GLOBAL.formatarComPontos(Format(mediaAlt, "0.0000"))
                                        
                ListBox.list(ListBox.ListCount - 1, 5) = _
                                        M_METODOS_GLOBAL.formatarComPontos(Format(totalM2, "0.0000"))
                ListBox.list(ListBox.ListCount - 1, 6) = objetoChapa.tipoPolimento.nome
                ListBox.list(ListBox.ListCount - 1, 7) = esp
                ListBox.list(ListBox.ListCount - 1, 8) = _
                                        M_METODOS_GLOBAL.formatarComPontos(Format(valorPolimento, "currency"))
                ListBox.list(ListBox.ListCount - 1, 9) = _
                                        M_METODOS_GLOBAL.formatarComPontos(Format(objetoChapa.valorTotal, "currency"))
                
                ' Limpando as variaveis
                qtdEstoque = 0
                mediaComp = 0
                mediaAlt = 0
                totalM2 = 0
                valorPolimento = ""
                esp = "02"
                
                ' Libera espaço da memoria
                Set objetoChapa = Nothing
            Next i
            ' Total de chapas
            lqtdChapasListaEstoque.Caption = qtdChapas
        Else
            ' Loop através dos itens da coleção
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
                                        M_METODOS_GLOBAL.formatarComPontos(Format(objeto.precoM3Bloco, "currency"))
                ListBox.list(ListBox.ListCount - 1, 7) = _
                                        M_METODOS_GLOBAL.formatarComPontos(Format(objeto.valoresAdicionais, "currency"))
                ListBox.list(ListBox.ListCount - 1, 8) = _
                                        M_METODOS_GLOBAL.formatarComPontos(Format(objeto.freteBloco, "currency"))
                ListBox.list(ListBox.ListCount - 1, 9) = _
                                        M_METODOS_GLOBAL.formatarComPontos(Format(objeto.valorBloco, "currency"))
                
                ' Total de blocos pesquisados
                lQtdBlocos.Caption = i
                ' Soma a qtd de chapas
                qtdChapas = qtdChapas + CInt(objeto.qtdChapas)
                ' Libera espaço da memoria
                Set objeto = Nothing
            Next i
            ' Total de chapas
            lQtdChapas.Caption = qtdChapas
        End If
        
        
    End If
    ' Libera espaço da memoria
    Set listaObjeto = Nothing
End Sub

' Carrega a lista ListTamanhosChapas tela edicao chapa
Private Sub carregarListTamanhosChapas(ListBox As MSForms.ListBox, listaCollection As Collection)
    'Variaveis do metodo
    Dim tamanho As objTamanho
    Dim totalChapas As Integer
    Dim i As Integer
    
    ' Limpar a ListBox
    ListBox.Clear
    
    ' NOME CABEÇALHO TAMANHOS     | TIPO  | COMP | ALT | M²  | QTD    | ESP    | CUSTO  | POLIDEIRA | ESTOQUE
    ' Tamanho do cabeçalho left   | 7,05  | 146  | 195 | 244 | 295,05 | 331,05 | 361,55 | 437,5     | 551
    ' Tamanho do cabeçalho width  | 138,5 | 48   | 48  | 50  | 35     | 30     | 75     | 112       | 114,5
    ' Tamanho das colunas da list
    ListBox.ColumnWidths = "138,5;48;48;50;35;30;75;112,5;114,5"
    
    ' Verifica se tem algum dado a pesquisa
    If listaCollection.Count = -1 Or listaCollection.Count = 0 Then ' Se não tiver dados
        
    Else
        ' Loop através dos itens da coleção
        For i = 1 To listaCollection.Count
            ' Seta o ojeto
            Set tamanho = listaCollection(i)
            
            ' Adiciona uma linha
            ListBox.AddItem
            
            'Adiciona os dados do bloco
            ListBox.list(ListBox.ListCount - 1, 0) = tamanho.tipoMaterial.nome
            ListBox.list(ListBox.ListCount - 1, 1) = M_METODOS_GLOBAL.formatarComPontos(Format(tamanho.compremento, "0.0000"))
            ListBox.list(ListBox.ListCount - 1, 2) = M_METODOS_GLOBAL.formatarComPontos(Format(tamanho.altura, "0.0000"))
            ListBox.list(ListBox.ListCount - 1, 3) = M_METODOS_GLOBAL.formatarComPontos(Format(tamanho.qtdM2, "0.0000"))
            ListBox.list(ListBox.ListCount - 1, 4) = tamanho.qtdEstoque
            ListBox.list(ListBox.ListCount - 1, 5) = tamanho.espessura
            ListBox.list(ListBox.ListCount - 1, 6) = M_METODOS_GLOBAL.formatarComPontos(Format(tamanho.valorPolimento, "currency"))
            ListBox.list(ListBox.ListCount - 1, 7) = tamanho.polideira.nome
            ListBox.list(ListBox.ListCount - 1, 8) = tamanho.estoque
            ListBox.list(ListBox.ListCount - 1, 9) = tamanho.id
            
            ' Soma as chapas
            totalChapas = totalChapas + CInt(tamanho.qtdEstoque)
            
            ' Libera espaço na memoria
            Set tamanho = Nothing
        Next i
        
        ' Seta total da tela
        txtEstoqueChapa.Value = totalChapas
    End If
End Sub

' Carrega a lista ListMateriaisParaTroca tela troca de estoque
Private Sub carregarListTrocasQtdChapas(ListBox As MSForms.ListBox, chapaTroca As objChapa, tamanhoTroca As objTamanho)
    
    ' Variaveis do metodo
    Dim qtd As Integer
    Dim m2 As Double
    Dim comp As Double
    Dim alt As Double
    
    ' Atribiucoes
    comp = CDbl(txtCompMaterialParaTroca.Value)
    alt = CDbl(txtAltMaterialParaTroca.Value)
    qtd = CDbl(txtQtdMaterialParaTroca02.Value)
    
    m2 = comp * alt * qtd
    
    ' Limpar a ListBox
    ListBox.Clear
    
    ' NOME CABEÇALHO TAMANHOS     | COD | DESCRIÇÃO | QTD   | M²
    ' Tamanho do cabeçalho left   | 7   | 118       | 346,5 | 399
    ' Tamanho do cabeçalho width  | 110 | 228       | 52    | 81
    ' Tamanho das colunas da list
    ListBox.ColumnWidths = "110;228;52;81;"
    
    ' Adiciona uma linha
    ListBox.AddItem
    
    'Adiciona os dados do bloco
    ListBox.list(ListBox.ListCount - 1, 0) = chapaTroca.idSistema
    ListBox.list(ListBox.ListCount - 1, 1) = chapaTroca.nomeMaterial
    ListBox.list(ListBox.ListCount - 1, 2) = txtQtdMaterialParaTroca02.Value
    ListBox.list(ListBox.ListCount - 1, 3) = m2
    ListBox.list(ListBox.ListCount - 1, 4) = tamanhoTroca.id
    
    ' Libera espaço na memoria
    Set chapaTroca = Nothing
End Sub

' Carrega as listas da tela Despache
Private Sub carregarListsTelaDespache(ListBox As MSForms.ListBox, listaColletion As Collection)
    ' Variaveis do metodo
    Dim qtd As Double
    Dim m2 As Double
    Dim i As Integer
    Dim j As Integer
    
    ' Limpar a ListBox
    ListBox.Clear
    
    ' NOME CABEÇALHO TAMANHOS     | COD | DESCRIÇÃO | QTD   | M²
    ' Tamanho do cabeçalho left   | 7   | 118       | 346,5 | 399
    ' Tamanho do cabeçalho width  | 110 | 228       | 52    | 81
    ' Tamanho das colunas da list
    ListBox.ColumnWidths = "110;228;52;81"
    
    For i = 1 To listaColletion.Count
        ' Seta chapa
        Set chapa = listaColletion.Item(i)
        
        ' Adiciona uma linha
        ListBox.AddItem
        
        ' Laço para soma do estoque e m²
        For j = 1 To chapa.tamanhos.Count
            ' Seta tamanho
            Set tamanho = chapa.tamanhos.Item(j)
            ' Soma
            qtd = qtd + CDbl(tamanho.qtdEstoque)
            m2 = m2 + CDbl(tamanho.qtdM2)
        Next j
        
        'Adiciona os dados da chapa
        ListBox.list(ListBox.ListCount - 1, 0) = chapa.idSistema
        ListBox.list(ListBox.ListCount - 1, 1) = chapa.nomeMaterial
        ListBox.list(ListBox.ListCount - 1, 2) = qtd
        ListBox.list(ListBox.ListCount - 1, 3) = Format(m2, "0.0000")
        ListBox.list(ListBox.ListCount - 1, 4) = chapa.bloco.idSistema
        
        ' Analisa se tem item na lista de despache
        If ListDespachado.ListCount > 0 Then
            ' Percorre a lista
            For j = 0 To ListDespachado.ListCount - 1
                ' Compara se já tem o material
                If ListDespachado.list(j, 0) = ListBox.list(i - 1, 0) Then
                    ' Atualiza quantidade na lista
                    ListBox.list(i - 1, 2) = ListBox.list(i - 1, 2) - ListDespachado.list(j, 2)
                    ' Atualiza m² na lista
                    ListBox.list(i - 1, 3) = Format(ListBox.list(i - 1, 3) - ListDespachado.list(j, 3), "0.0000")
                End If
            Next j
        End If
        
        ' Libera espaço na memoria
        Set chapa = Nothing
        qtd = 0
        m2 = 0
    Next i
    ' Libera espaço na memoria
    Set listaColletion = Nothing
End Sub

' Carrega as listas de tamanhos da chapa da tela Despache
Private Sub carregarTamanhosChapasTelaDespache(ListBox As MSForms.ListBox, listaColletion As Collection)
    ' Variaveis do metodo
    Dim i As Integer
    Dim j As Integer
    
    ' Limpar a ListBox
    ListBox.Clear
    
    ' NOME CABEÇALHO TAMANHOS     | QTD   | COMP | ALT   | ESP  | TM
    ' Tamanho do cabeçalho left   | 210,5 | 257  | 317,5 | 378  | 423
    ' Tamanho do cabeçalho width  | 45    | 60   | 60    | 44,5 | 44,5
    ' Tamanho das colunas da list
    ListBox.ColumnWidths = "45;60;60;44,5;44,5"
    
    For i = 1 To listaColletion.Count
        ' Seta chapa
        Set tamanho = listaColletion.Item(i)
        
        ' Adiciona uma linha
        ListBox.AddItem
        
        'Adiciona os dados da chapa
        ListBox.list(ListBox.ListCount - 1, 0) = tamanho.qtdEstoque
        ListBox.list(ListBox.ListCount - 1, 1) = M_METODOS_GLOBAL.formatarComPontos(Format(tamanho.compremento, "0.0000"))
        ListBox.list(ListBox.ListCount - 1, 2) = M_METODOS_GLOBAL.formatarComPontos(Format(tamanho.altura, "0.0000"))
        ListBox.list(ListBox.ListCount - 1, 3) = tamanho.espessura
        ' Seta sigla do polimento
        Select Case tamanho.tipoMaterial.nome
            Case "EXTRA"
                ListBox.list(ListBox.ListCount - 1, 4) = "EX"
            Case "SEMI-EXTRA"
                ListBox.list(ListBox.ListCount - 1, 4) = "SEX"
            Case "PEÇA"
                ListBox.list(ListBox.ListCount - 1, 4) = "PE"
            Case "COMERCIAL A"
                ListBox.list(ListBox.ListCount - 1, 4) = "CA"
            Case "COMERCIAL B"
                ListBox.list(ListBox.ListCount - 1, 4) = "CB"
            Case "COMERCIAL C"
                ListBox.list(ListBox.ListCount - 1, 4) = "CC"
            Case "COMERCIAL D"
                ListBox.list(ListBox.ListCount - 1, 4) = "CD"
            Case "COMERCIAL E"
                ListBox.list(ListBox.ListCount - 1, 4) = "CE"
            Case "COMERCIAL F"
                ListBox.list(ListBox.ListCount - 1, 4) = "CF"
            Case "STARNDER"
                ListBox.list(ListBox.ListCount - 1, 4) = "ST"
        End Select
        ListBox.list(ListBox.ListCount - 1, 5) = tamanho.id
        ListBox.list(ListBox.ListCount - 1, 6) = tamanho.estoque.nome
        
        ' Analisa se tem item na lista de despache
        If ListDespachado.ListCount > 0 Then
            ' Percorre a lista
            For j = 0 To ListDespachado.ListCount - 1
                ' Compara se já tem o material
                If ListDespachado.list(j, 4) = ListBox.list(ListBox.ListCount - 1, 5) Then
                    ' Atualiza quantidade na lista
                    ListBox.list(i - 1, 0) = ListBox.list(i - 1, 0) - ListDespachado.list(j, 2)
'                    ' Atualiza m² na lista
'                    ListBox.list(i - 1, 3) = Format(ListBox.list(i - 1, 3) - ListDespachado.list(j, 3), "0.0000")
                End If
            Next j
        End If
        ' Libera espaço na memoria
        Set tamanho = Nothing
    Next i
    ' Libera espaço na memoria
    Set listaColletion = Nothing
End Sub

