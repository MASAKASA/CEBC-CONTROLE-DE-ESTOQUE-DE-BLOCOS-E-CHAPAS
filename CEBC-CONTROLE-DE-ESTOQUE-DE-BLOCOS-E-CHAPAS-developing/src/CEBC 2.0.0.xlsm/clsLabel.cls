VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents efeitoBotoesMenu As MSForms.Label
Attribute efeitoBotoesMenu.VB_VarHelpID = -1
Public WithEvents efeitoBotoesImagem As MSForms.Label
Attribute efeitoBotoesImagem.VB_VarHelpID = -1
Public WithEvents efeitoBotoesTexto As MSForms.Label
Attribute efeitoBotoesTexto.VB_VarHelpID = -1

' Variavel do menu
Public rMenuOnClick As Integer
Public gMenuOnClick As Integer
Public bMenuOnClick As Integer
Public rMenuOffClick As Integer
Public gMenuOffClick As Integer
Public bMenuOffClick As Integer
Public leftMenuOnClick As Integer
Public WidthMenuOnClick As Integer
Public leftMenuOffClick As Integer
Public WidthMenuOffClick As Integer
Public FontSizeBig As Integer
Public FontSizeSmall As Integer
'Efeito ao clicar com o mouse
Private Sub efeitoBotoesMenu_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ' OBS.:
    '   1 - TODA VEZ QUE ADICIONAR MAIS UM ITEM NO MENU LEMBRAR DE
    '       TROCAR A COR DO FUNDO DA LABEL QUE FICA POR TR�S DA LABEL
    '       PRINCIPAL(QUE � A DO BOT�O)
    '
    '   2 - SE CASO TROCAR A CORES E EFEITOS DOS BOT�ES DO MENU, IR NO
    '       MOUSEDOWN DO BOT�O ADDESTOQUE NA TELA ESTOQUE M3 PARA TROCAR
    '       AS CONFIGURA��O DO BOT�ES BLOCO E CHAPA DO MENU
    
    ' Atribuicoes
    rMenuOnClick = 200
    gMenuOnClick = 230
    bMenuOnClick = 255
    rMenuOffClick = 0
    gMenuOffClick = 100
    bMenuOffClick = 200
    FontSizeBig = 32
    FontSizeSmall = 20
    leftMenuOnClick = 15
    WidthMenuOnClick = 172
    leftMenuOffClick = 2
    WidthMenuOffClick = 189
    
    'Efeitos paras os label bot�es do menu
    Select Case efeitoBotoesMenu.name
        
        ' Bot�es no menu
        Case "btnLMenuHome"
            ' Bot�o home
            formControle.Controls("btnLMenuHome").BackColor = RGB(rMenuOnClick, gMenuOnClick, bMenuOnClick)
            formControle.Controls("btnLMenuHome").Font.Size = FontSizeBig
            formControle.Controls("btnLMenuHome").Font.Size = FontSizeSmall
            formControle.Controls("btnLMenuHome").Left = leftMenuOnClick
            formControle.Controls("btnLMenuHome").Width = WidthMenuOnClick
            formControle.Controls("btnLMenuHome").TextAlign = fmTextAlignCenter
            
            ' Demais bot�es no menu
            formControle.Controls("btnLMenuBloco").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuBloco").Left = leftMenuOffClick
            formControle.Controls("btnLMenuBloco").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuChapa").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuChapa").Left = leftMenuOffClick
            formControle.Controls("btnLMenuChapa").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuDespachar").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuDespachar").Left = leftMenuOffClick
            formControle.Controls("btnLMenuDespachar").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuDespachar").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuCarrago").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuCarrago").Left = leftMenuOffClick
            formControle.Controls("btnLMenuCarrago").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuCarrago").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuCadastros").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuCadastros").Left = leftMenuOffClick
            formControle.Controls("btnLMenuCadastros").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuCadastros").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuUsuarios").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuUsuarios").Left = leftMenuOffClick
            formControle.Controls("btnLMenuUsuarios").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuUsuarios").TextAlign = fmTextAlignLeft
            
        Case "btnLMenuBloco"
            ' Bot�o bloco
            formControle.Controls("btnLMenuBloco").BackColor = RGB(rMenuOnClick, gMenuOnClick, bMenuOnClick)
            formControle.Controls("btnLMenuBloco").Font.Size = FontSizeBig
            formControle.Controls("btnLMenuBloco").Font.Size = FontSizeSmall
            formControle.Controls("btnLMenuBloco").Left = leftMenuOnClick
            formControle.Controls("btnLMenuBloco").Width = WidthMenuOnClick
            formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignCenter
            
            ' Demais bot�es no menu
            formControle.Controls("btnLMenuHome").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuHome").Left = leftMenuOffClick
            formControle.Controls("btnLMenuHome").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuHome").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuChapa").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuChapa").Left = leftMenuOffClick
            formControle.Controls("btnLMenuChapa").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuDespachar").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuDespachar").Left = leftMenuOffClick
            formControle.Controls("btnLMenuDespachar").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuDespachar").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuCarrago").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuCarrago").Left = leftMenuOffClick
            formControle.Controls("btnLMenuCarrago").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuCarrago").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuCadastros").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuCadastros").Left = leftMenuOffClick
            formControle.Controls("btnLMenuCadastros").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuCadastros").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuUsuarios").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuUsuarios").Left = leftMenuOffClick
            formControle.Controls("btnLMenuUsuarios").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuUsuarios").TextAlign = fmTextAlignLeft
            
        Case "btnLMenuChapa"
            ' Bot�o chapa
            formControle.Controls("btnLMenuChapa").BackColor = RGB(rMenuOnClick, gMenuOnClick, bMenuOnClick)
            formControle.Controls("btnLMenuChapa").Font.Size = FontSizeBig
            formControle.Controls("btnLMenuChapa").Font.Size = FontSizeSmall
            formControle.Controls("btnLMenuChapa").Left = leftMenuOnClick
            formControle.Controls("btnLMenuChapa").Width = WidthMenuOnClick
            formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignCenter
            
            ' Demais bot�es no menu
            formControle.Controls("btnLMenuHome").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuHome").Left = leftMenuOffClick
            formControle.Controls("btnLMenuHome").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuHome").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuBloco").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuBloco").Left = leftMenuOffClick
            formControle.Controls("btnLMenuBloco").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuDespachar").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuDespachar").Left = leftMenuOffClick
            formControle.Controls("btnLMenuDespachar").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuDespachar").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuCarrago").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuCarrago").Left = leftMenuOffClick
            formControle.Controls("btnLMenuCarrago").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuCarrago").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuCadastros").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuCadastros").Left = leftMenuOffClick
            formControle.Controls("btnLMenuCadastros").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuCadastros").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuUsuarios").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuUsuarios").Left = leftMenuOffClick
            formControle.Controls("btnLMenuUsuarios").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuUsuarios").TextAlign = fmTextAlignLeft
            
        Case "btnLMenuDespachar"
            ' Bot�o despachar
            formControle.Controls("btnLMenuDespachar").BackColor = RGB(rMenuOnClick, gMenuOnClick, bMenuOnClick)
            formControle.Controls("btnLMenuDespachar").Font.Size = FontSizeBig
            formControle.Controls("btnLMenuDespachar").Font.Size = FontSizeSmall
            formControle.Controls("btnLMenuDespachar").Left = leftMenuOnClick
            formControle.Controls("btnLMenuDespachar").Width = WidthMenuOnClick
            formControle.Controls("btnLMenuDespachar").TextAlign = fmTextAlignCenter
            
            ' Demais bot�es no menu
            formControle.Controls("btnLMenuHome").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuHome").Left = leftMenuOffClick
            formControle.Controls("btnLMenuHome").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuHome").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuBloco").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuBloco").Left = leftMenuOffClick
            formControle.Controls("btnLMenuBloco").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuChapa").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuChapa").Left = leftMenuOffClick
            formControle.Controls("btnLMenuChapa").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuCarrago").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuCarrago").Left = leftMenuOffClick
            formControle.Controls("btnLMenuCarrago").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuCarrago").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuCadastros").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuCadastros").Left = leftMenuOffClick
            formControle.Controls("btnLMenuCadastros").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuCadastros").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuUsuarios").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuUsuarios").Left = leftMenuOffClick
            formControle.Controls("btnLMenuUsuarios").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuUsuarios").TextAlign = fmTextAlignLeft
            
        Case "btnLMenuCarrago"
            ' Bot�o carrego
            formControle.Controls("btnLMenuCarrago").BackColor = RGB(rMenuOnClick, gMenuOnClick, bMenuOnClick)
            formControle.Controls("btnLMenuCarrago").Font.Size = FontSizeBig
            formControle.Controls("btnLMenuCarrago").Font.Size = FontSizeSmall
            formControle.Controls("btnLMenuCarrago").Left = leftMenuOnClick
            formControle.Controls("btnLMenuCarrago").Width = WidthMenuOnClick
            formControle.Controls("btnLMenuCarrago").TextAlign = fmTextAlignCenter
            
            ' Demais bot�es no menu
            formControle.Controls("btnLMenuHome").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuHome").Left = leftMenuOffClick
            formControle.Controls("btnLMenuHome").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuHome").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuBloco").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuBloco").Left = leftMenuOffClick
            formControle.Controls("btnLMenuBloco").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuChapa").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuChapa").Left = leftMenuOffClick
            formControle.Controls("btnLMenuChapa").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuDespachar").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuDespachar").Left = leftMenuOffClick
            formControle.Controls("btnLMenuDespachar").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuDespachar").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuCadastros").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuCadastros").Left = leftMenuOffClick
            formControle.Controls("btnLMenuCadastros").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuCadastros").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuUsuarios").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuUsuarios").Left = leftMenuOffClick
            formControle.Controls("btnLMenuUsuarios").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuUsuarios").TextAlign = fmTextAlignLeft
            
        Case "btnLMenuUsuarios"
            ' Bot�o usuarios
            formControle.Controls("btnLMenuUsuarios").BackColor = RGB(rMenuOnClick, gMenuOnClick, bMenuOnClick)
            formControle.Controls("btnLMenuUsuarios").Font.Size = FontSizeBig
            formControle.Controls("btnLMenuUsuarios").Font.Size = FontSizeSmall
            formControle.Controls("btnLMenuUsuarios").Left = leftMenuOnClick
            formControle.Controls("btnLMenuUsuarios").Width = WidthMenuOnClick
            formControle.Controls("btnLMenuUsuarios").TextAlign = fmTextAlignCenter
            
            ' Demais bot�es no menu
            formControle.Controls("btnLMenuHome").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuHome").Left = leftMenuOffClick
            formControle.Controls("btnLMenuHome").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuHome").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuBloco").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuBloco").Left = leftMenuOffClick
            formControle.Controls("btnLMenuBloco").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuChapa").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuChapa").Left = leftMenuOffClick
            formControle.Controls("btnLMenuChapa").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuDespachar").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuDespachar").Left = leftMenuOffClick
            formControle.Controls("btnLMenuDespachar").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuDespachar").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuCarrago").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuCarrago").Left = leftMenuOffClick
            formControle.Controls("btnLMenuCarrago").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuCarrago").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuCadastros").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuCadastros").Left = leftMenuOffClick
            formControle.Controls("btnLMenuCadastros").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuCadastros").TextAlign = fmTextAlignLeft
            
        Case "btnLMenuCadastros"
            ' Bot�o usuarios
            formControle.Controls("btnLMenuCadastros").BackColor = RGB(rMenuOnClick, gMenuOnClick, bMenuOnClick)
            formControle.Controls("btnLMenuCadastros").Font.Size = FontSizeBig
            formControle.Controls("btnLMenuCadastros").Font.Size = FontSizeSmall
            formControle.Controls("btnLMenuCadastros").Left = leftMenuOnClick
            formControle.Controls("btnLMenuCadastros").Width = WidthMenuOnClick
            formControle.Controls("btnLMenuCadastros").TextAlign = fmTextAlignCenter
            
            ' Demais bot�es no menu
            formControle.Controls("btnLMenuHome").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuHome").Left = leftMenuOffClick
            formControle.Controls("btnLMenuHome").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuHome").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuBloco").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuBloco").Left = leftMenuOffClick
            formControle.Controls("btnLMenuBloco").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuChapa").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuChapa").Left = leftMenuOffClick
            formControle.Controls("btnLMenuChapa").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuDespachar").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuDespachar").Left = leftMenuOffClick
            formControle.Controls("btnLMenuDespachar").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuDespachar").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuCarrago").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuCarrago").Left = leftMenuOffClick
            formControle.Controls("btnLMenuCarrago").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuCarrago").TextAlign = fmTextAlignLeft
            
            formControle.Controls("btnLMenuUsuarios").BackColor = RGB(rMenuOffClick, gMenuOffClick, bMenuOffClick)
            formControle.Controls("btnLMenuUsuarios").Left = leftMenuOffClick
            formControle.Controls("btnLMenuUsuarios").Width = WidthMenuOffClick
            formControle.Controls("btnLMenuUsuarios").TextAlign = fmTextAlignLeft
            
    End Select
End Sub
'Efeito de passagem por cima do mouse
Private Sub efeitoBotoesMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Efeitos paras os label bot�es do menu
    efeitoBotoesMenu.Font.Size = 22
End Sub
'Efeito ao clicar com o mouse
Private Sub efeitoBotoesImagem_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Efeitos paras os label bot�es com imagens
    Select Case efeitoBotoesImagem.name
        ' Tela cadastrar bloco
        Case "btnLImgCadastrarPedreira"
            efeitoBotoesImagem.top = 24
        Case "btnLImgCadastrarSerrariaCB"
            efeitoBotoesImagem.top = 78
        Case "btnLImgCadastroTipoMaterial"
            efeitoBotoesImagem.top = 24
        
        ' Tela despache
        Case "btnLImgCadastrarMotoristaDespache"
            efeitoBotoesImagem.top = 28
        Case "btnLImgCadastrarDestinoDespache"
            efeitoBotoesImagem.top = 76
            
        ' Tela estoque m�
        Case "btnLImgExportarEstoqueM3"
            efeitoBotoesImagem.top = 18
            
        ' Tela estoque m�
        Case "btnLImgExportarEstoqueM2"
            efeitoBotoesImagem.top = 6
            
        ' Tela cadastro chapas avulso
        Case "btnLImgCadastrarMaterialAvulso"
            efeitoBotoesImagem.top = 24
        Case "btnLImgCadastrarPolimentoAvulso"
            efeitoBotoesImagem.top = 78
            
        ' Tela carregos
        Case "btnLImgExportarCarregoPDF"
            efeitoBotoesImagem.top = 24
            
        ' Tela lan�amento e edi��o chapa
        Case "btnLImgCadastrarPolideiraChapa"
            efeitoBotoesImagem.top = 24
        Case "btnLImgCadastrarTipoPolideiraChapa"
            efeitoBotoesImagem.top = 78
        Case "btnLImgCadastrarTipoMaterialChapa"
            efeitoBotoesImagem.top = 132
        Case "btnLImgCadastrarTipoMaterialChapaTamanhos"
            efeitoBotoesImagem.top = 24
    End Select
End Sub
'Efeito de passagem por cima do mouse
Private Sub efeitoBotoesImagem_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Efeitos paras os label bot�es com imagens
    Select Case efeitoBotoesImagem.name
        ' Tela cadastrar bloco
        Case "btnLImgCadastrarPedreira"
            efeitoBotoesImagem.top = 22
         Case "btnLImgCadastrarSerrariaCB"
            efeitoBotoesImagem.top = 76
        Case "btnLImgCadastroTipoMaterial"
            efeitoBotoesImagem.top = 22
        
        ' Tela despache
        Case "btnLImgCadastrarMotoristaDespache"
            efeitoBotoesImagem.top = 28
        Case "btnLImgCadastrarDestinoDespache"
            efeitoBotoesImagem.top = 76
        Case "btnLImgExportarListaDespache"
            efeitoBotoesImagem.top = 4
            
        ' Tela estoque m�
        Case "btnLImgExportarEstoqueM3"
            efeitoBotoesImagem.top = 16
            
        ' Tela estoque m�
        Case "btnLImgExportarEstoqueM2"
            efeitoBotoesImagem.top = 4
            
        ' Tela cadastro chapas avulso
        Case "btnLImgCadastrarMaterialAvulso"
            efeitoBotoesImagem.top = 22
        Case "btnLImgCadastrarPolimentoAvulso"
            efeitoBotoesImagem.top = 76
            
        ' Tela carregos
        Case "btnLImgExportarCarregoPDF"
            efeitoBotoesImagem.top = 22
            
        ' Tela lan�amento e edi��o chapa
        Case "btnLImgCadastrarPolideiraChapa"
            efeitoBotoesImagem.top = 22
        Case "btnLImgCadastrarTipoPolideiraChapa"
            efeitoBotoesImagem.top = 76
        Case "btnLImgCadastrarTipoMaterialChapa"
            efeitoBotoesImagem.top = 130
        Case "btnLImgCadastrarTipoMaterialChapaTamanhos"
            efeitoBotoesImagem.top = 22
    End Select
End Sub
'Efeito ao clicar com o mouse
Private Sub efeitoBotoesTexto_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Efeitos paras os label bot�es com texto
    efeitoBotoesTexto.Font.Size = 32
    efeitoBotoesTexto.Font.Size = 16
End Sub
'Efeito de passagem por cima do mouse
Private Sub efeitoBotoesTexto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Efeitos paras os label bot�es com texto
    efeitoBotoesTexto.Font.Size = 18
End Sub
