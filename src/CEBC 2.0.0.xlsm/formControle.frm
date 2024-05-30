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
Dim paginaAnterior As Integer
Dim status() As String
Dim listaObjeto As Collection

' Variaveis de objetos
Dim bloco As objBloco
Dim chapa As objChapa
Dim pedreira As objPedreira
Dim polideira As objPolideira
Dim serraria As objSerraria
Dim tamanho As objTamanho
Dim tipoMaterial As objTipoMaterial
Dim tipoPolimento As objTipoPolimento
Dim statusObj As objStatus
Dim estoque As objEstoque
Dim estoqueChapa As objEstoqueChapa

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
        ' Mensagem que já existe cadastro de chapa com esse numero da pedreira
        errorStyle.Informativo ESCOLHA_CHAPA_MENSAGEM, ESCOLHA_CHAPA_TITULO
    Else
        ' Direciona para tela lançamento e edição de chapa
        Me.MultiPageCEBC.Value = 6
        ' Carrega combox da tela lançamento e edição de chapa
        Call carregarTiposMateriais(Me.cbTipoMaterialChapaC)
        Call carregarEstoqueChapas(Me.cbEstoqueChapaC)
        
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
        
        chapaCadastro.carregarChapa idChapa, descricaoChapa, valorTotalSerrada, bloco.qtdChapas, bloco.qtdM2Serrada, _
                        bloco.compBrutoChapaBruta, bloco.altBrutoChapaBruta, bloco.numeroBlocoPedreira, tipoPolimento, _
                        bloco, tamanhos
                            
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
Private Sub txtTotalChapaBlocoEditar_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Deixa só a digitação de numero
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
    ' Chama Serviço
    MsgBox "Chama Serviço esportar estoque m², tela estoque m²"
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
    If Me.ListEstoqueChapas.ListIndex = -1 Then
        ' Mensagem usuário
        errorStyle.Informativo ESCOLHA_CHAPA_MENSAGEM, ESCOLHA_CHAPA_TITULO
        Exit Sub
    End If
    
    ' Muda abra da multPage para tela editar bloco
    Me.MultiPageCEBC.Value = 6
    ' Seta paniga anterior para futuras condições
    paginaAnterior = 4
    
    ' Analisar quais tipos de polimentos vão ser carregador
    Set chapaPesquisa = daoChapa.pesquisarPorId(Me.ListEstoqueChapas.list(Me.ListEstoqueChapas.ListIndex, 0))
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
                        tipoPolimento, bloco, polideira, tamanhos
                        
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
    txtTotalBlocoAvulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(M_METODOS_GLOBAL.calcularValorBloco( _
            txtTotalM2Avulso.Value, txtValorMetroAvulso.Value), "0.00"))
            
    'Se m² for diferente de 0 calcula o custo do material
    'Seta o custo do material m²
    txtCustoSimplesM2Avulso.Value = M_METODOS_GLOBAL.formatarComPontos(Format(Util.custoMaterialM2(txtTotalBlocoAvulso.Value, _
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
    Dim blocoPesquisa As objBloco
    Dim resposta As VbMsgBoxResult ' Variavel para confirmação na hora de cadastrar
    Dim nomeStatus As String
    Dim nomeMaterial As String
    Dim valorTotalBloco As String
    Dim cadastro As Boolean
    
    ' Patrão true
    cadastro = True
    
    ' Captura do status
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
    Set blocoPesquisa = daoBloco.pesquisarPorId(txtIdBlocoAvulso.Value)
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
        Set tipoMaterial = daoTipoMaterial.pesquisarPorNome(cbTipoMaterialL.Value)
        Set tipoPolimento = daoTipoPolimento.pesquisarPorNome(cbTipoPolimentoL.Value)
        Set statusObj = daoStatus.pesquisarPorNome("ESTOQUE")
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
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
        'Variaveis no metodo
    Dim idBloco As String
    Dim descricao As String
    Dim adicionais As String
    Dim precoBloco As String
    Dim valorFreteBloco As String
    Dim valorM3 As String
    Dim quantidadeM3 As String
    Dim idBlocoPedreira As String
    Dim largura As String
    Dim altura As String
    Dim comprimento As String
    Dim dataCadastro As String
    Dim observacao As String
    Dim tipoMaterial As String
    Dim idEstoque As String
    Dim nomePedreira As String
    Dim nomeStatus As String
    Dim nomeSerraria As String
    Dim temNota As String
    
    'Variaveis para as chapa
    Dim idChapa As String
    Dim descricaoChapa As String
    Dim totalChapas As String
    Dim alturaChapas As String
    Dim comprimentoChapas As String
    Dim espessura As String
    Dim m2Chapas As String
    Dim tipoPolimento As String
    Dim custoMaterial As String
    Dim idEstoqueChapa As String
    Dim nomePolideira As String
    
    'Capturando informações do usuario para atualização no bloco
    idBloco = UCase(TextBoxIdBlocoAvulsoSistema.Value)
    observacao = UCase(TextBoxObsBlocoL.Value)
    precoBloco = UCase(TextBoxValorBlocoAvulso.Value)
    adicionais = UCase(txtAdicionaisAvulso.Value)
    valorFreteBloco = UCase(TextBoxValorFreteAvulso.Value)
    valorSerrada = UCase("0,00")
    valorPolimento = UCase("0,00")
    valoresAdicionais = UCase("0,00")
    custoSimples = UCase(TextBoxCustoSimplesM2Avulso.Value)
    nomePedreira = "IMPORTADO"
    nomeSerraria = "IMPORTADO"
    idBlocoPedreira = UCase(TextBoxIdBlocoAvulso.Value)
    nomeStatus = "ESTOQUE"
    dataCadastro = UCase(TextBoxDataCadastroChapaAvulsa.Value)
    descricao = "BLOCO " & TextBoxMaterialAvulso.Value
    temNota = ComboBoxTemNotaAvulso.Value
    idEstoque = "1"
    largura = "0,0000"
    altura = "0,0000"
    comprimento = "0,0000"
    valorM3 = "0,0000"
    quantidadeM3 = "0,0000"
    
    'Capturando informações do usuario para atualização na chapa
    totalChapas = UCase(TextBoxQuantidadeChapasAvulsas.Value)
    alturaChapas = UCase(TextBoxAlturaChapaAvulsa.Value)
    comprimentoChapas = UCase(TextBoxComprimentoChapaAvulsa.Value)
    m2Chapas = UCase(TextBoxTtalM2Avulso.Value)
    espessura = UCase(txtEspessuraAvulso.Value)
    idEstoqueChapa = "1"
    tipoMaterial = ComboBoxTipoMaterialL.Value
    idChapa = Util.formatarIdChapa(idBloco, "A")
    descricaoChapa = UCase(TextBoxMaterialAvulso.Value)
    nomePolideira = "IMPORTADO"
    
    If OptionButtonAvulso.Value = True Then
        nomePedreira = "AVULSO"
        nomeSerraria = "AVULSO"
        nomePolideira = "AVULSO"
    End If
    
    'Validações
    
    'Verificando o Id do bloco
    If idBlocoPedreira = "" Then

        'Deixa o cursor na ser adicionado o id
        TextBoxIdBlocoAvulso.SetFocus

        'Altera cor para melhor visualização
        TextBoxIdBlocoAvulso.BackColor = RGB(255, 182, 193)

        'Mensagem de erro
        MsgBox "Adicone o número do bloco!", vbCritical, "ID do bloco não informado"

        'Para o fluxo do sistema para a correção
        Exit Sub
    End If

    'Volta a cor patrão
    TextBoxIdBlocoAvulso.BackColor = RGB(255, 255, 255)

    'Verificando o nome do bloco
    If descricaoChapa = "" Or descricaoChapa = " A" Then

        'Deixa o cursor na ser adicionado o nome do bloco
        TextBoxMaterialAvulso.SetFocus

        'Altera cor para melhor visualização
        TextBoxMaterialAvulso.BackColor = RGB(255, 182, 193)

        'Mensagem de erro
        MsgBox "Adicone a descrição da chapa!", vbCritical, "Descrição da chapa não informada"

        'Para o fluxo do sistema para a correção
        Exit Sub
    End If

    'Volta a cor patrão
    TextBoxMaterialAvulso.BackColor = RGB(255, 255, 255)

    'Verificando o tipo do material
    If tipoMaterial = "" Then

        'Deixa o cursor na ser adicionado o tipo do material
        ComboBoxTipoMaterialL.SetFocus

        'Altera cor para melhor visualização
        ComboBoxTipoMaterialL.BackColor = RGB(255, 182, 193)

        'Mensagem de erro
        MsgBox "Selecione o tipo do material!", vbCritical, "Tipo de material não informado"

        'Para o fluxo do sistema para a correção
        Exit Sub
    End If

    'Volta a cor patrão
    ComboBoxTipoMaterialL.BackColor = RGB(255, 255, 255)

    'Captura o tipo de polimento, cria o id e descrição da chapa e valida ComboBoxTipoPolimentoL
    If ComboBoxTipoPolimentoL.Value = "" Then

        'Deixa o cursor na ser adicionado o tipo de polimento
        ComboBoxTipoPolimentoL.SetFocus

        'Altera cor para melhor visualização
        ComboBoxTipoPolimentoL.BackColor = RGB(255, 182, 193)

        'Mensagem de erro
        MsgBox "Adiciona o tipo de polimento!", vbCritical, "Tipo de polimento não informado"

        'Para o fluxo do sistema para a correção
        Exit Sub

    ElseIf ComboBoxTipoPolimentoL.Value = "POLIDO" Then

        idChapa = Util.formatarIdChapa(idBloco, "PO")
        descricaoChapa = Util.formatarNomeChapa(descricao, "POLIDO")
        tipoPolimento = "POLIDO"

    ElseIf ComboBoxTipoPolimentoL.Value = "BI POLIDO" Then

        idChapa = Util.formatarIdChapa(idBloco, "BPO")
        descricaoChapa = Util.formatarNomeChapa(descricao, "BI POLIDO")
        tipoPolimento = "BI POLIDO"

    ElseIf ComboBoxTipoPolimentoL.Value = "ESCOVADO" Then

        idChapa = Util.formatarIdChapa(idBloco, "ES")
        descricaoChapa = Util.formatarNomeChapa(descricao, "ESCOVADO")
        tipoPolimento = "ESCOVADO"

    ElseIf ComboBoxTipoPolimentoL.Value = "BI ESCOVADO" Then

        idChapa = Util.formatarIdChapa(idBloco, "BES")
        descricaoChapa = Util.formatarNomeChapa(descricao, "BI ESCOVADO")
        tipoPolimento = "BI ESCOVADO"

    ElseIf ComboBoxTipoPolimentoL.Value = "LEVIGADO" Then

        idChapa = Util.formatarIdChapa(idBloco, "LE")
        descricaoChapa = Util.formatarNomeChapa(descricao, "LEVIGADO")
        tipoPolimento = "LEVIGADO"

    ElseIf ComboBoxTipoPolimentoL.Value = "FLAMIADO" Then

        idChapa = Util.formatarIdChapa(idBloco, "FL")
        descricaoChapa = Util.formatarNomeChapa(descricao, "FLAMIADO")
        tipoPolimento = "FLAMIADO"

    ElseIf ComboBoxTipoPolimentoL.Value = "RIPADO" Then

        idChapa = Util.formatarIdChapa(idBloco, "RI")
        descricaoChapa = Util.formatarNomeChapa(descricao, "RIPADO")
        tipoPolimento = "RIPADO"

    ElseIf ComboBoxTipoPolimentoL.Value = "RIPADO" Then

        idChapa = Util.formatarIdChapa(idBloco, "MA")
        descricaoChapa = Util.formatarNomeChapa(descricao, "MATTE")
        tipoPolimento = "MATTE"
        
    ElseIf ComboBoxTipoPolimentoL.Value = "RIPADO" Then

        idChapa = Util.formatarIdChapa(idBloco, "RP")
        descricaoChapa = Util.formatarNomeChapa(descricao, "RESIN PINTADO")
        tipoPolimento = "RESIN PINTADO"
    End If

    'Altera cor para melhor visualização
    ComboBoxTipoPolimentoL.BackColor = RGB(255, 255, 255)

    'Verificando o tipo de polimento e valida ComboBoxTipoPolimentoL
    If precoBloco = "0,00" Or precoBloco = "" Then

        'Deixa o cursor na ser adicionado o tipo de polimento
        TextBoxValorBlocoAvulso.SetFocus

        'Altera cor para melhor visualização
        TextBoxValorBlocoAvulso.BackColor = RGB(255, 182, 193)

        'Mensagem de erro
        MsgBox "Adiciona o custo do bloco!", vbCritical, "Custo do bloco não informado"

        'Para o fluxo do sistema para a correção
        Exit Sub

    End If

    'Altera cor para melhor visualização
    TextBoxValorBlocoAvulso.BackColor = RGB(255, 255, 255)

    'Verificando o valor do frete
    If valorFreteBloco = "0,00" Or valorFreteBloco = "" Then

        'Deixa o cursor na ser adicionado o frete
        TextBoxValorFreteAvulso.SetFocus

        'Altera cor para melhor visualização
        TextBoxValorFreteAvulso.BackColor = RGB(255, 182, 193)

        'Mensagem de erro
        MsgBox "Adiciona o valor do frete!", vbCritical, "Valor do frete não informado"

        'Para o fluxo do sistema para a correção
        Exit Sub
    End If

    'Altera cor para melhor visualização
    TextBoxValorFreteAvulso.BackColor = RGB(255, 255, 255)

    'Verificando a quantidades de chapas
    If totalChapas = "0" Or totalChapas = "" Then

        'Deixa o cursor na ser adicionado a quantidade
        TextBoxQuantidadeChapasAvulsas.SetFocus

        'Altera cor para melhor visualização
        TextBoxQuantidadeChapasAvulsas.BackColor = RGB(255, 182, 193)

        'Mensagem de erro
        MsgBox "Adiciona a quantidade de chapas!", vbCritical, "Quantidade de chapas não informado"

        'Para o fluxo do sistema para a correção
        Exit Sub
    End If

    'Volta a cor patrão
    TextBoxQuantidadeChapasAvulsas.BackColor = RGB(255, 255, 255)

        'Verificando o comprimento
    If comprimentoChapas = "0,0000" Or comprimentoChapas = "" Then

        'Deixa o cursor na ser adicionado a quantidade
        TextBoxComprimentoChapaAvulsa.SetFocus

        'Altera cor para melhor visualização
        TextBoxComprimentoChapaAvulsa.BackColor = RGB(255, 182, 193)

        'Mensagem de erro
        MsgBox "Adiciona o comprimento!", vbCritical, "Comprimneto não informado"

        'Para o fluxo do sistema para a correção
        Exit Sub
    End If

    'Volta a cor patrão
    TextBoxComprimentoChapaAvulsa.BackColor = RGB(255, 255, 255)

    'Verificando a altura
    If alturaChapas = "0,0000" Or alturaChapas = "" Then

        'Deixa o cursor na ser adicionado a quantidade
        TextBoxAlturaChapaAvulsa.SetFocus

        'Altera cor para melhor visualização
        TextBoxAlturaChapaAvulsa.BackColor = RGB(255, 182, 193)

        'Mensagem de erro
        MsgBox "Adiciona a altura!", vbCritical, "Altura não informado"

        'Para o fluxo do sistema para a correção
        Exit Sub
    End If

    'Volta a cor patrão
    TextBoxAlturaChapaAvulsa.BackColor = RGB(255, 255, 255)

    'Verificando a espessura
    If espessura = "" Then

        'Deixa o cursor na ser adicionado a espessura
        txtEspessuraAvulso.SetFocus

        'Altera cor para melhor visualização
        txtEspessuraAvulso.BackColor = RGB(255, 182, 193)

        'Mensagem de erro
        MsgBox "Adiciona a espessura!", vbCritical, "Espessura não informado"

        'Para o fluxo do sistema para a correção
        Exit Sub
    End If

    'Volta a cor patrão
    txtEspessuraAvulso.BackColor = RGB(255, 255, 255)
    
    'Mensagem de confirmação
    
    resposta = MsgBox("Confira se o número e descrição/material do bloco estão corretos, pois a junção deles irá criar o ID do bloco no sistema. ID do bloco não pederá ser alterado posteriormente. Tudo conferido e podemos seguir com o cadastro?", vbQuestion + vbYesNo, "Atenção - Confirmação")
    
    'Verifica a confirmação do usário para poder cadastrar
    If resposta = vbYes Then
        
        Call cadastrarBlocoComSerraria(idBloco, descricao, adicionais, precoBloco, valorM3, quantidadeM3, _
            idBlocoPedreira, largura, altura, comprimento, dataCadastro, observacao, _
            tipoMaterial, valorFreteBloco, idEstoque, nomePedreira, nomeStatus, nomeSerraria, temNota)
        
        'Para o processo de cadastro
        If PARAR_PROCESSO = True Then
            
            'Retorna o valor patrão
            PARAR_PROCESSO = False ' Variavel Globol que foi declarada em M_GLOBAL
            
            Exit Sub
        End If
        
        'Cadastra chapa
        Call cadastrarChapa(idChapa, descricaoChapa, custoSimples, custoSimples, totalChapas, m2Chapas, _
                comprimentoChapas, alturaChapas, espessura, idBlocoPedreira, tipoPolimento, idEstoqueChapa, _
                tipoMaterial, nomePolideira, idBloco)
                
        'Para o processo de cadastro
        If PARAR_PROCESSO = True Then

            'Retorna o valor patrão
            PARAR_PROCESSO = False ' Variavel Globol que foi declarada em M_GLOBAL

            Exit Sub
        End If
        
        'Carregar o lisBox com chapas cadastradas
        carregarListBoxMaterias
        
        'Mensagem de cadastro realizado com sucesso
        MsgBox "Chapas cadastradas com sucesso!", vbInformation, "Cadastrado de Chapas avulso"
    
    Else
    
        ' Coloque o código a ser executado se o usuário clicar em "Não" aqui.
        MsgBox "A ação foi cancelada."
        
        'Para o processo
        Exit Sub
    End If
    
    'Limpa os campos
    Call btnLimparCadastroAvulso_Click
    
    'Deixa o cursor no TextBoxIdBlocoAvulso
    TextBoxIdBlocoAvulso.SetFocus

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ' Seta o foco
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
    
    ' Dados chapa
    txtIdChapaSistema.Value = chapa.idSistema
    txtDescricaoChapa.Value = chapa.nomeMaterial
    txtEstoqueChapa.Value = chapa.qtdEstoque
    
    ' Dimensões e custos
    Call selecaoItem("cbPolideiraChapa", chapa.polideira.nome)
    Call selecaoItem("cbTipoPolimentoChapa", chapa.tipoPolimento.nome)
    Call selecaoItem("cbEstoqueChapaC", chapa.estoque.nome)
    txtQtdChapaC.Value = chapa.qtdEstoque
    txtCompBrutoChapa.Value = M_METODOS_GLOBAL.formatarComPontos(Format(chapa.compBruto, "0.0000"))
    txtAlturaBrutaChapa.Value = M_METODOS_GLOBAL.formatarComPontos(Format(chapa.altBruto, "0.0000"))
    txtQtsM2Chapa.Value = M_METODOS_GLOBAL.formatarComPontos(Format(chapa.qtdM2Bruto, "0.0000"))
    txtCustoChapa.Value = M_METODOS_GLOBAL.formatarComPontos(Format(chapa.custoPolimento, "0.00"))
    txtTotalChapas.Value = M_METODOS_GLOBAL.formatarComPontos(Format(chapa.custoTotal, "0.00"))
    
    ' Carrega lista com tamanhos das chapas
    Call carregarListTamanhosChapas(ListTamanhosChapas, chapa.tamanhos) ' Irá enviar id chapa para carregamento
        
    ' Verifica se a lista só tem um tamanho
    If chapa.tamanhos.Count = 1 Then
        For i = 1 To chapa.tamanhos.Count
            ' Seta o ojeto
            Set tamanho = chapa.tamanhos(i)
            
            ' Tamanho único
            Call selecaoItem("cbTipoMaterialChapaC", tamanho.tipoMaterial.nome)
            txtCompLiquidoChapa.Value = M_METODOS_GLOBAL.formatarComPontos(Format(tamanho.compremento, "0.0000"))
            txtAlturaLiquidoChapa.Value = M_METODOS_GLOBAL.formatarComPontos(Format(tamanho.altura, "0.0000"))
            txtEspChapa.Value = tamanho.espessura

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
    Dim idchapas As String
    Dim descricaoChapa As String
    Dim codFinal As String
    Dim posicao As Integer
    Dim idBloco As String
    Dim descricao As String
    
    '
'    posicao = Len(cbTipoPolimentoChapa.Value) - 2
'    codFinal = Mid(cbTipoPolimentoChapa.Value, posicao, 3)
'    ' Comparação para logica
'    If cbTipoPolimentoChapa.Value = "BRUTO" Then
'
'    End If
    ' Id do bloco
    idBloco = txtIdBlocoPedreiraChapa.Value
    descricao = txtDecricaoBlocoChapa.Value
     
    'Captura o tipo de polimento, cria o id e descrição da chapa
    If cbTipoPolimentoChapa.Value = "BI POLIDO" Then
        idchapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "BPO")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "BI POLIDO")
        
    ElseIf cbTipoPolimentoChapa.Value = "ESCOVADO" Then
        idchapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "ES")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "ESCOVADO")
        
    ElseIf cbTipoPolimentoChapa.Value = "BI ESCOVADO" Then
        idchapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "BES")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "BI ESCOVADO")
        
    ElseIf cbTipoPolimentoChapa.Value = "LEVIGADO" Then
        idchapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "LE")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "LEVIGADO")
        
    ElseIf cbTipoPolimentoChapa.Value = "FLAMIADO" Then
        idchapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "FL")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "FLAMIADO")
        
    ElseIf cbTipoPolimentoChapa.Value = "RIPADO" Then
        idchapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "RI")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "RIPADO")
        
    ElseIf cbTipoPolimentoChapa.Value = "POLIDO" Then
        idchapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "PO")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "POLIDO")
        
    ElseIf cbTipoPolimentoChapa.Value = "MATTE" Then
        idchapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "MA")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "MATTE")
        
    ElseIf cbTipoPolimentoChapa.Value = "RESIN PINTADO" Then
        idchapas = M_METODOS_GLOBAL.formatarIdChapa(idBloco, "RP")
        descricaoChapa = M_METODOS_GLOBAL.formatarNomeChapa(descricao, "RESIN PINTADO")
    End If
    
    ' Seta id e descrição
    txtIdChapaSistema.Value = idchapas
    txtDescricaoChapa.Value = descricaoChapa
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
    ' Seta o foco
    cbPolideiraChapa.SetFocus
End Sub

' Botão btnLTxtVoltarChapa tela lançamento e edição chapa
Private Sub btnLTxtVoltarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Volta para tela que chamou
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = paginaAnterior
        
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
' Botão btnLTxtAdicionarTrocaEstoque tela troca estoque
Private Sub btnLTxtAdicionarTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Chama Serviço
    MsgBox "Chama Serviço adicionar material para troca, tela troca estoque"
    ' Seta o foco
    txtMaterialParaTroca02.SetFocus
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
    ' Seta o foco
    txtMaterialChapaPesquisa.SetFocus
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
    ' Seta o foco
    txtMaterial.SetFocus
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

'-----------------------------------------------------------------PESQUISAR-------------------------------------------------------
'                                                                 ---------
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
    
    'Validando a data
    If IsDate(txtDataInicioBlocoPesquisa.Value) = False Then
        ' Deixa visivel o erro com mensagens
        errorStyle.EntrarErrorStyleTextBox txtDataInicioBlocoPesquisa, ADICIONE_DATA_MENSAGEM, ADICIONE_DATA_TITULO
        'Para o fluxo do sistema para a correção
        Exit Sub
    End If
    ' Deixa na cor patrão
    errorStyle.sairErrorStyleTextBox txtDataInicioBlocoPesquisa

    'Validando a data
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
    txtIdChapaSistema.Enabled = False
    txtDescricaoChapa.Enabled = False
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
    txtIdChapaSistema.Enabled = False
    txtDescricaoChapa.Enabled = True
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
    txtCompLiquidoChapa.Enabled = True
    txtAlturaLiquidoChapa.Enabled = True
    txtQtdChapaC.Enabled = True
    txtCompBrutoChapa.Enabled = True
    txtAlturaBrutaChapa.Enabled = True
    txtEspChapa.Enabled = True
    txtQtsM2Chapa.Enabled = True
    txtCustoChapa.Enabled = True
    txtTotalChapas.Enabled = True
    cbTiposMateriaisChapas.Enabled = True
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
        Call selecaoItem("cbTipoMaterial", "EXTRA")
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
    If Me.MultiPageCEBC.Value = 2 Then
        Call selecaoItem("cbNotaC", "NÃO")
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

'-----------------------------------------------------------------CARREAGMENTO DAS LIST-----------------------------------
'                                                                 ---------------------
' Carrega a lista bloco
Private Sub carregarList(ListBox As MSForms.ListBox, listaCollection As Collection)
   'Variaveis do metodo
    Dim objeto As objBloco
    Dim objetoChapa As objChapa
    Dim i As Integer
    Dim qtdChapas As Integer
    
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
                'ListBox.list(ListBox.ListCount - 1, 2) = objetoChapa.qtdEstoque
'                ListBox.list(ListBox.ListCount - 1, 3) = _
'                                        M_METODOS_GLOBAL.formatarComPontos(Format(objetoChapa.compBruto, "0.0000"))
'                ListBox.list(ListBox.ListCount - 1, 4) = _
'                                        M_METODOS_GLOBAL.formatarComPontos(Format(objetoChapa.altBruto, "0.0000"))
'                ListBox.list(ListBox.ListCount - 1, 5) = _
'                                        M_METODOS_GLOBAL.formatarComPontos(Format(objetoChapa.qtdM2Bruto, "0.0000"))
'                ListBox.list(ListBox.ListCount - 1, 6) = objetoChapa.tipoPolimento.nome
'                ListBox.list(ListBox.ListCount - 1, 7) = "02"
                ListBox.list(ListBox.ListCount - 1, 8) = _
                                        M_METODOS_GLOBAL.formatarComPontos(Format(0, "currency"))
                ListBox.list(ListBox.ListCount - 1, 9) = _
                                        M_METODOS_GLOBAL.formatarComPontos(Format(objetoChapa.valorTotal, "currency"))
                
                ' Libera espaço da memoria
                Set objetoChapa = Nothing
            Next i
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
    Dim i As Integer
    
    ' Limpar a ListBox
    ListBox.Clear
    
    ' NOME CABEÇALHO BLOCOS       | TIPO  | ESP   | COMP  | ALT | M²  | QTD |
    ' Tamanho do cabeçalho left   | 192,5 | 331,5 | 362,5 | 411 | 460 | 511 |
    ' Tamanho do cabeçalho width  | 138,5 | 30    | 48    | 48  | 50  | 30  |
    ' Tamanho das colunas da list
    ListBox.ColumnWidths = "140,5;30;48;48;50;35;"
    
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
            ListBox.list(ListBox.ListCount - 1, 1) = tamanho.espessura
            ListBox.list(ListBox.ListCount - 1, 2) = M_METODOS_GLOBAL.formatarComPontos(Format(tamanho.compremento, "0.0000"))
            ListBox.list(ListBox.ListCount - 1, 3) = M_METODOS_GLOBAL.formatarComPontos(Format(tamanho.altura, "0.0000"))
            ListBox.list(ListBox.ListCount - 1, 4) = M_METODOS_GLOBAL.formatarComPontos(Format(tamanho.qtdM2, "0.0000"))
            ListBox.list(ListBox.ListCount - 1, 5) = tamanho.qtdEstoque
            
            ' Libera espaço na memoria
            Set tamanho = Nothing
        Next i
    End If
End Sub
