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
' Variaveis para manipula��o com varios metodos
Dim tamanhoColunasList As String
'Inicializa��o do formControle
Private Sub UserForm_Initialize()
    ' Variaveis para o metodo
    Dim obj As Object
    Dim i As Long
    Dim j As Long
    Dim l As Long
    Dim m As Long
    Dim nameObj As String
    Dim nameObjInicio As String
    
    ' Resevando espa�o em memoria para manipula��o das variaveis
    ReDim botoesMenu(1 To Me.Controls.Count)
    ReDim botoesImg(1 To Me.Controls.Count)
    ReDim botoesText(1 To Me.Controls.Count)
    ReDim frameEfeito(1 To Me.Controls.Count)
    
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
    
    tamanhoColunasList = "185;250;52;52;52;52;90;75;75;74;"
    
    ' Retira os nomes de cima da multPage
    Me.MultiPageCEBC.Style = fmTabStyleNone
    
    ' Chama metodo para carregar lista e blocos cadastros do dia atual
    Call carregarListCadastradosHoje
End Sub

'-----------------------------------------------------------------MENU DO SISTEMA-----------------------------------
'                                                                 ---------------
'Efeito para clique nas label btnLMenuHome do menu
Private Sub btnLMenuHome_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 0
End Sub
'Efeito para clique nas label btnLMenuBloco do menu
Private Sub btnLMenuBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 1
End Sub
'Efeito para clique nas label btnLMenuChapa do menu
Private Sub btnLMenuChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 4
End Sub
'Efeito para clique nas label btnLMenuDespachar do menu
Private Sub btnLMenuDespachar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 8
End Sub
'Efeito para clique nas label btnLMenuCarrago do menu
Private Sub btnLMenuCarrago_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 9
End Sub
'Efeito para clique nas label btnLMenuCadastros do menu
Private Sub btnLMenuCadastros_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 10
End Sub
'Efeito para clique nas label btnLMenuUsuarios do menu
Private Sub btnLMenuUsuarios_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 11
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
'Bot�o btnLTxtPesquisarBlocos tela estoque m�
Private Sub btnLTxtPesquisarBlocos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o pesquiar, tela estoque m�"
End Sub
'Bot�o btnLTxtLimparFiltrosBlocos tela estoque m�
Private Sub btnLTxtLimparFiltrosBlocos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o limpar filtros, tela estoque m�"
End Sub
'Bot�o btnLImgExportarEstoqueM3 tela estoque m�
Private Sub btnLImgExportarEstoqueM3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o exportar estoque m�, tela estoque m�"
End Sub
'Bot�o btnLTxtNovoBloco tela estoque m�
Private Sub btnLTxtNovoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Coloca data atual na txtDataCadastro na tela cadastro de bloco
    txtDataCadastro.Value = Date
    
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 2
End Sub
'Bot�o btnLTxtEditarBloco tela estoque m�
Private Sub btnLTxtEditarBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 3
End Sub
'Bot�o btnLTxtADDEstoque tela estoque m�
Private Sub btnLTxtADDEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
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
    
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 6
End Sub

'-----------------------------------------------------------------TELA CADASTRO DE BLOCOS-----------------------------------
'                                                                 -----------------------
' txtIdBloco tela cadastro de bloco
Private Sub txtIdBloco_Change()
    ' Coloca tudo em caixa alta
    txtIdBloco.Value = UCase(txtIdBloco.Value)
    
    ' Cria o c�digo para o sistema
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
    
    ' Cria o c�digo para o sistema
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
' Carrega a lista listCadastradosHoje tela cadastro de bloco
Private Sub carregarListCadastradosHoje()
'    'Variaveis do metodo
'    Dim Data As String
'    Dim listaBlocos As Variant
'    Dim totalDia As Double
'    Dim tamanhoLista As Integer
'    Dim i As Long
    
'    'Formata a data
'    Data = Util.ConverterFormatoData(txtDataCadastro.Value)
'
'    'Criando Cole��o para manipula��
'    listaBlocos = BlocosDAO.listarBlocosEntreDatas(Data, Data)

    ' Limpar a ListBox
    Me.listCadastradosHoje.Clear
    
    ' Tamanho das colunas, String fica no metodo de inicializa��o do sistema UserForm_Initialize
    Me.listCadastradosHoje.ColumnWidths = tamanhoColunasList

'    'Captura o tamanho da matriz
'    tamanhoLista = TamanhoDaMatriz(listaBlocos)
    
'    'Sem n�o tiver dados
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
            listCadastradosHoje.AddItem
            
            'Adiciona os dados do bloco
            listCadastradosHoje.List(listCadastradosHoje.ListCount - 1, 0) = "37766-50793-MOON-LIGHT-BL"
            listCadastradosHoje.List(listCadastradosHoje.ListCount - 1, 1) = "BLOCO BRANCO DALLAS MOON-LIGHT"
            listCadastradosHoje.List(listCadastradosHoje.ListCount - 1, 2) = "3,0000"
            listCadastradosHoje.List(listCadastradosHoje.ListCount - 1, 3) = "2,0000"
            listCadastradosHoje.List(listCadastradosHoje.ListCount - 1, 4) = "2,0000"
            listCadastradosHoje.List(listCadastradosHoje.ListCount - 1, 5) = "71"
            listCadastradosHoje.List(listCadastradosHoje.ListCount - 1, 6) = "1.500,00"
            listCadastradosHoje.List(listCadastradosHoje.ListCount - 1, 7) = "15.000,00"
            listCadastradosHoje.List(listCadastradosHoje.ListCount - 1, 8) = "5.000,00"
            listCadastradosHoje.List(listCadastradosHoje.ListCount - 1, 9) = "150.000,00"
                
'            'Soma o total do dia
'            lTotalDia = lTotalDia + Util.formatarComPontos(Format(Util.calcularValorBloco(CStr(listaBlocos(i, 11)), _
'                CStr(listaBlocos(i, 9))), "0.00"))
'        Next i
    
'        lTotalDia.Caption = listCadastradosHoje.List(listCadastradosHoje.ListCount - 1, 9)
'    End If

End Sub
'Bot�o btnLImgCadastrarPedreira tela cadastrar bloco
Private Sub btnLImgCadastrarPedreira_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar pedreira, tela cadastrar bloco"
End Sub
'Bot�o btnLImgCadastrarSerrariaCB tela cadastrar bloco
Private Sub btnLImgCadastrarSerrariaCB_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar serraria, tela cadastrar bloco"
End Sub
'Bot�o btnLImgCadastroTipoMaterial tela cadastrar bloco
Private Sub btnLImgCadastroTipoMaterial_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar tipo material, tela cadastrar bloco"
End Sub
'Bot�o btnLTxtCadastrarBloco tela cadastrar bloco
Private Sub btnLTxtCadastrarBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar bloco, tela cadastrar bloco"
End Sub
'Bot�o btnLTextLimparCadastroBloco tela cadastrar bloco
Private Sub btnLTxtLimparCadastroBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o limpar campos, tela cadastro de blocos"
End Sub

'-----------------------------------------------------------------TELA DESPACHE-----------------------------------
'                                                                 -------------
'Bot�o btnLImgCadastrarMotoristaDespache tela despache
Private Sub btnLImgCadastrarMotoristaDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastro motorista, tela despache"
End Sub
'Bot�o btnLImgCadastrarDestinoDespache tela despache
Private Sub btnLImgCadastrarDestinoDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastro destino, tela despache"
End Sub
'Bot�o btnLTxtAdicionar tela despache
Private Sub btnLTxtAdicionar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o adicionar chapa, tela despache"
End Sub
'Bot�o btnLTxtDespachar tela despache
Private Sub btnLTxtDespachar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o despachar, tela despache"
End Sub
'Bot�o btnLTxtLimparDespache tela despache
Private Sub btnLTxtLimparDespache_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o limpar dados, tela despache"
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
    MsgBox "Chama Servi�o pesquisar chapa, tela estoque m�"
End Sub
'Bot�o btnLTxtLimparFiltrosChapas tela estoque m�
Private Sub btnLTxtLimparFiltrosChapas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o limpar filtros, tela estoque m�"
End Sub
'Bot�o btnLImgExportarEstoqueM2 tela estoque m�
Private Sub btnLImgExportarEstoqueM2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o esportar estoque m�, tela estoque m�"
End Sub
'Bot�o btnLTxtNovoAvulso tela estoque m�
Private Sub btnLTxtNovoAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 5
End Sub
'Bot�o btnLTxtEditarChapa tela estoque m�
Private Sub btnLTxtEditarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 6
End Sub
'Bot�o btnLTxtTrocaEstoque tela estoque m�
Private Sub btnLTxtTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Muda abra da multPage
    Me.MultiPageCEBC.Value = 7
End Sub

'-----------------------------------------------------------------TELA CADASTRO AVULSO-----------------------------------
'                                                                 --------------------
'Bot�o btnLImgCadastrarMaterialAvulso tela cadastro avulso
Private Sub btnLImgCadastrarMaterialAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar tipo material, tela cadastro avulso"
End Sub
'Bot�o btnLImgCadastrarPolimentoAvulso tela cadastro avulso
Private Sub btnLImgCadastrarPolimentoAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar tipo polimento, tela cadastro avulso"
End Sub
'Bot�o btnLTxtCadastrarChapaAvulso tela cadastro avulso
Private Sub btnLTxtCadastrarChapaAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar chapas avulsos, tela cadastro avulso"
End Sub
'Bot�o btnLTxtLimparCadastroChapaAvulso tela cadastro avulso
Private Sub btnLTxtLimparCadastroChapaAvulso_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o limpar cadastro chapas avulsos, tela cadastro avulso"
End Sub

'-----------------------------------------------------------------TELA CARREGOS-----------------------------------
'                                                                 -------------
'Bot�o btnLTxtPesquisarCarregos tela carregos
Private Sub btnLTxtPesquisarCarregos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o pesquisar por carregos, tela carregos"
End Sub
'Bot�o btnLTxtLimparListas tela carregos
Private Sub btnLTxtLimparListas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o limpar dados filtro, tela carregos"
End Sub
'Bot�o btnLImgExportarCarregoPDF tela carregos
Private Sub btnLImgExportarCarregoPDF_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o exportar carregos em pdf, tela carregos"
End Sub
'Bot�o btnLTxtEditarCarrego tela carregos
Private Sub btnLTxtEditarCarrego_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o editar carrego, tela carregos"
End Sub
'Bot�o btnLTxtVoltarCArrego tela carregos
Private Sub btnLTxtVoltarCArrego_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o voltar, tela carregos"
End Sub

'-----------------------------------------------------------------TELA EDITAR BLOCO-----------------------------------
'                                                                 -----------------
'Bot�o btnLTxtSalvarEdicaoBloco tela carregos
Private Sub btnLTxtSalvarEdicaoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o editar bloco, tela editar bloco"
End Sub
'Bot�o btnLTxtVoltarEdicaoBloco tela carregos
Private Sub btnLTxtVoltarEdicaoBloco_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o voltar, tela editar bloco"
End Sub

'-----------------------------------------------------------------TELA LAN�AMENTO E EDI��O CHAPA-----------------------------------
'                                                                 ------------------------------
'Bot�o btnLImgCadastrarPolideiraChapa tela lan�amento e edi��o chapa
Private Sub btnLImgCadastrarPolideiraChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar Polideira, tela lan�amento e edi��o chapa"
End Sub
'Bot�o btnLImgCadastrarTipoPolideiraChapa tela lan�amento e edi��o chapa
Private Sub btnLImgCadastrarTipoPolideiraChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar tipo polimento, tela lan�amento e edi��o chapa"
End Sub
'Bot�o btnLImgCadastrarTipoMaterialChapa tela lan�amento e edi��o chapa
Private Sub btnLImgCadastrarTipoMaterialChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o tipo material, tela lan�amento e edi��o chapa"
End Sub
'Bot�o btnLImgCadastrarTipoMaterialChapaTamanhos tela lan�amento e edi��o chapa
Private Sub btnLImgCadastrarTipoMaterialChapaTamanhos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o tipo material, tela lan�amento e edi��o chapa"
End Sub
'Bot�o btnLTxtAdicionarTamanhoChapa tela lan�amento e edi��o chapa
Private Sub btnLTxtAdicionarTamanhoChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o adicionar tamanhos, tela lan�amento e edi��o chapa"
End Sub
'Bot�o btnLTxtEditarTamanhoChapa tela lan�amento e edi��o chapa
Private Sub btnLTxtEditarTamanhoChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o editar tamanho chapa, tela lan�amento e edi��o chapa"
End Sub
'Bot�o btnLTxtTirarDaLista tela lan�amento e edi��o chapa
Private Sub btnLTxtTirarDaLista_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o tira tamanho da lista, tela lan�amento e edi��o chapa"
End Sub
'Bot�o btnLTxtSalvarChapa tela lan�amento e edi��o chapa
Private Sub btnLTxtSalvarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o salva altera��o da chapa, tela lan�amento e edi��o chapa"
End Sub
'Bot�o btnLTxtVoltarChapa tela lan�amento e edi��o chapa
Private Sub btnLTxtVoltarChapa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o voltar, tela lan�amento e edi��o chapa"
End Sub


'-----------------------------------------------------------------TELA TROCA ESTOQUE-----------------------------------
'                                                                 ------------------
'Bot�o btnLTxtAdicionarTrocaEstoque tela troca estoque
Private Sub btnLTxtAdicionarTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o adicionar material para troca, tela troca estoque"
End Sub
'Bot�o btnLTxtTrocarEstoque tela troca estoque
Private Sub btnLTxtTrocarEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o troca de estoque, tela troca estoque"
End Sub
'Bot�o btnLTxtVoltarTrocaEstoque tela troca estoque
Private Sub btnLTxtVoltarTrocaEstoque_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o voltar, tela troca estoque"
End Sub


'-----------------------------------------------------------------TELA CADASTROS DIVERSOS-----------------------------------
'                                                                 -----------------------
'Bot�o btnLTxtSalvarPedreira tela cadastros diversos
Private Sub btnLTxtSalvarPedreira_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar pedreira, tela cadastros diversos"
End Sub
'Bot�o btnLTxtSalvarSerraria tela cadastros diversos
Private Sub btnLTxtSalvarSerraria_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar serraria, tela cadastros diversos"
End Sub
'Bot�o btnLTxtSalvarPolideira tela cadastros diversos
Private Sub btnLTxtSalvarPolideira_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar polideira, tela cadastros diversos"
End Sub
'Bot�o btnLTxtSalvarTipoMaterial tela cadastros diversos
Private Sub btnLTxtSalvarTipoMaterial_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar tipo material, tela cadastros diversos"
End Sub
'Bot�o btnLTxtSalvarTipoPolimento tela cadastros diversos
Private Sub btnLTxtSalvarTipoPolimento_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar tipo polimento, tela cadastros diversos"
End Sub
'Bot�o btnLTxtSalvarMotorista tela cadastros diversos
Private Sub btnLTxtSalvarMotorista_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar motorista, tela cadastros diversos"
End Sub
'Bot�o btnLTxtSalvarMotorista tela cadastros diversos
Private Sub btnLTxtSalvarDestino_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar destino, tela cadastros diversos"
End Sub


'-----------------------------------------------------------------TELA USUARIO-----------------------------------
'                                                                 ------------
'Bot�o btnLTxtSalvarUsuario tela usuarios
Private Sub btnLTxtSalvarUsuario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o cadastrar ou editar usu�rio, tela usuarios"
End Sub
'Bot�o btnLTxtListUsuario tela usuarios
Private Sub btnLTxtListUsuario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o carrega lista com usu�rios, tela usuarios"
End Sub
'Bot�o btnLTxtListUsuarioLog tela usuarios
Private Sub btnLTxtListUsuarioLog_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Chama Servi�o
    MsgBox "Chama Servi�o carrega lista com log dos usu�rios, tela usuarios"
End Sub
