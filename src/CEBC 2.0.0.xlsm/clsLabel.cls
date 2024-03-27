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
'Efeito ao clicar com o mouse
Private Sub efeitoBotoesMenu_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Efeitos paras os label botões do menu
    Select Case efeitoBotoesMenu.name
        
        ' Botões no menu
        Case "btnLMenuHome"
            ' Botão home
            formControle.Controls("btnLMenuHome").Font.Size = 32
            formControle.Controls("btnLMenuHome").Font.Size = 20
            formControle.Controls("btnLMenuHome").Left = 15
            formControle.Controls("btnLMenuHome").Width = 200
            formControle.Controls("btnLMenuHome").TextAlign = fmTextAlignCenter
            
            ' Demais botões no menu
            formControle.Controls("btnLMenuBloco").Left = 5
            formControle.Controls("btnLMenuBloco").Width = 187
            formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuChapa").Left = 5
            formControle.Controls("btnLMenuChapa").Width = 187
            formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuDespachar").Left = 5
            formControle.Controls("btnLMenuDespachar").Width = 187
            formControle.Controls("btnLMenuDespachar").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuCarrago").Left = 5
            formControle.Controls("btnLMenuCarrago").Width = 187
            formControle.Controls("btnLMenuCarrago").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuUsuarios").Left = 5
            formControle.Controls("btnLMenuUsuarios").Width = 187
            formControle.Controls("btnLMenuUsuarios").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuCadastros").Left = 5
            formControle.Controls("btnLMenuCadastros").Width = 187
            formControle.Controls("btnLMenuCadastros").TextAlign = fmTextAlignLeft
            
        Case "btnLMenuBloco"
            ' Botão bloco
            formControle.Controls("btnLMenuBloco").Font.Size = 32
            formControle.Controls("btnLMenuBloco").Font.Size = 20
            formControle.Controls("btnLMenuBloco").Left = 15
            formControle.Controls("btnLMenuBloco").Width = 200
            formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignCenter
            
            ' Demais botões no menu
            formControle.Controls("btnLMenuHome").Left = 5
            formControle.Controls("btnLMenuHome").Width = 187
            formControle.Controls("btnLMenuHome").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuChapa").Left = 5
            formControle.Controls("btnLMenuChapa").Width = 187
            formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuDespachar").Left = 5
            formControle.Controls("btnLMenuDespachar").Width = 187
            formControle.Controls("btnLMenuDespachar").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuCarrago").Left = 5
            formControle.Controls("btnLMenuCarrago").Width = 187
            formControle.Controls("btnLMenuCarrago").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuUsuarios").Left = 5
            formControle.Controls("btnLMenuUsuarios").Width = 187
            formControle.Controls("btnLMenuUsuarios").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuCadastros").Left = 5
            formControle.Controls("btnLMenuCadastros").Width = 187
            formControle.Controls("btnLMenuCadastros").TextAlign = fmTextAlignLeft
            
        Case "btnLMenuChapa"
            ' Botão chapa
            formControle.Controls("btnLMenuChapa").Font.Size = 32
            formControle.Controls("btnLMenuChapa").Font.Size = 20
            formControle.Controls("btnLMenuChapa").Left = 15
            formControle.Controls("btnLMenuChapa").Width = 200
            formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignCenter
            
            ' Demais botões no menu
            formControle.Controls("btnLMenuHome").Left = 5
            formControle.Controls("btnLMenuHome").Width = 187
            formControle.Controls("btnLMenuHome").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuBloco").Left = 5
            formControle.Controls("btnLMenuBloco").Width = 187
            formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuDespachar").Left = 5
            formControle.Controls("btnLMenuDespachar").Width = 187
            formControle.Controls("btnLMenuDespachar").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuCarrago").Left = 5
            formControle.Controls("btnLMenuCarrago").Width = 187
            formControle.Controls("btnLMenuCarrago").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuUsuarios").Left = 5
            formControle.Controls("btnLMenuUsuarios").Width = 187
            formControle.Controls("btnLMenuUsuarios").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuCadastros").Left = 5
            formControle.Controls("btnLMenuCadastros").Width = 187
            formControle.Controls("btnLMenuCadastros").TextAlign = fmTextAlignLeft
            
        Case "btnLMenuDespachar"
            ' Botão despachar
            formControle.Controls("btnLMenuDespachar").Font.Size = 32
            formControle.Controls("btnLMenuDespachar").Font.Size = 20
            formControle.Controls("btnLMenuDespachar").Left = 15
            formControle.Controls("btnLMenuDespachar").Width = 200
            formControle.Controls("btnLMenuDespachar").TextAlign = fmTextAlignCenter
            
            ' Demais botões no menu
            formControle.Controls("btnLMenuHome").Left = 5
            formControle.Controls("btnLMenuHome").Width = 187
            formControle.Controls("btnLMenuHome").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuBloco").Left = 5
            formControle.Controls("btnLMenuBloco").Width = 187
            formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuChapa").Left = 5
            formControle.Controls("btnLMenuChapa").Width = 187
            formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuCarrago").Left = 5
            formControle.Controls("btnLMenuCarrago").Width = 187
            formControle.Controls("btnLMenuCarrago").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuUsuarios").Left = 5
            formControle.Controls("btnLMenuUsuarios").Width = 187
            formControle.Controls("btnLMenuUsuarios").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuCadastros").Left = 5
            formControle.Controls("btnLMenuCadastros").Width = 187
            formControle.Controls("btnLMenuCadastros").TextAlign = fmTextAlignLeft
            
        Case "btnLMenuCarrago"
            ' Botão carrego
            formControle.Controls("btnLMenuCarrago").Font.Size = 32
            formControle.Controls("btnLMenuCarrago").Font.Size = 20
            formControle.Controls("btnLMenuCarrago").Left = 15
            formControle.Controls("btnLMenuCarrago").Width = 200
            formControle.Controls("btnLMenuCarrago").TextAlign = fmTextAlignCenter
            
            ' Demais botões no menu
            formControle.Controls("btnLMenuHome").Left = 5
            formControle.Controls("btnLMenuHome").Width = 187
            formControle.Controls("btnLMenuHome").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuBloco").Left = 5
            formControle.Controls("btnLMenuBloco").Width = 187
            formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuChapa").Left = 5
            formControle.Controls("btnLMenuChapa").Width = 187
            formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuDespachar").Left = 5
            formControle.Controls("btnLMenuDespachar").Width = 187
            formControle.Controls("btnLMenuDespachar").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuUsuarios").Left = 5
            formControle.Controls("btnLMenuUsuarios").Width = 187
            formControle.Controls("btnLMenuUsuarios").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuCadastros").Left = 5
            formControle.Controls("btnLMenuCadastros").Width = 187
            formControle.Controls("btnLMenuCadastros").TextAlign = fmTextAlignLeft
            
        Case "btnLMenuUsuarios"
            ' Botão usuarios
            formControle.Controls("btnLMenuUsuarios").Font.Size = 32
            formControle.Controls("btnLMenuUsuarios").Font.Size = 20
            formControle.Controls("btnLMenuUsuarios").Left = 15
            formControle.Controls("btnLMenuUsuarios").Width = 200
            formControle.Controls("btnLMenuUsuarios").TextAlign = fmTextAlignCenter
            
            ' Demais botões no menu
            formControle.Controls("btnLMenuHome").Left = 5
            formControle.Controls("btnLMenuHome").Width = 187
            formControle.Controls("btnLMenuHome").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuBloco").Left = 5
            formControle.Controls("btnLMenuBloco").Width = 187
            formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuChapa").Left = 5
            formControle.Controls("btnLMenuChapa").Width = 187
            formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuDespachar").Left = 5
            formControle.Controls("btnLMenuDespachar").Width = 187
            formControle.Controls("btnLMenuDespachar").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuCarrago").Left = 5
            formControle.Controls("btnLMenuCarrago").Width = 187
            formControle.Controls("btnLMenuCarrago").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuCadastros").Left = 5
            formControle.Controls("btnLMenuCadastros").Width = 187
            formControle.Controls("btnLMenuCadastros").TextAlign = fmTextAlignLeft
            
        Case "btnLMenuCadastros"
            ' Botão usuarios
            formControle.Controls("btnLMenuCadastros").Font.Size = 32
            formControle.Controls("btnLMenuCadastros").Font.Size = 20
            formControle.Controls("btnLMenuCadastros").Left = 15
            formControle.Controls("btnLMenuCadastros").Width = 200
            formControle.Controls("btnLMenuCadastros").TextAlign = fmTextAlignCenter
            
            ' Demais botões no menu
            formControle.Controls("btnLMenuHome").Left = 5
            formControle.Controls("btnLMenuHome").Width = 187
            formControle.Controls("btnLMenuHome").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuBloco").Left = 5
            formControle.Controls("btnLMenuBloco").Width = 187
            formControle.Controls("btnLMenuBloco").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuChapa").Left = 5
            formControle.Controls("btnLMenuChapa").Width = 187
            formControle.Controls("btnLMenuChapa").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuDespachar").Left = 5
            formControle.Controls("btnLMenuDespachar").Width = 187
            formControle.Controls("btnLMenuDespachar").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuCarrago").Left = 5
            formControle.Controls("btnLMenuCarrago").Width = 187
            formControle.Controls("btnLMenuCarrago").TextAlign = fmTextAlignLeft
            formControle.Controls("btnLMenuUsuarios").Left = 5
            formControle.Controls("btnLMenuUsuarios").Width = 187
            formControle.Controls("btnLMenuUsuarios").TextAlign = fmTextAlignLeft
    End Select
End Sub
'Efeito de passagem por cima do mouse
Private Sub efeitoBotoesMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Efeitos paras os label botões do menu
    efeitoBotoesMenu.Font.Size = 22
End Sub
'Efeito ao clicar com o mouse
Private Sub efeitoBotoesImagem_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Efeitos paras os label botões com imagens
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
            efeitoBotoesImagem.top = 70
        Case "btnLImgCadastrarDestinoDespache"
            efeitoBotoesImagem.top = 70
        Case "btnLImgExportarListaDespache"
            efeitoBotoesImagem.top = 6
            
        ' Tela estoque m³
        Case "btnLImgExportarEstoqueM3"
            efeitoBotoesImagem.top = 18
            
        ' Tela estoque m²
        Case "btnLImgExportarEstoqueM2"
            efeitoBotoesImagem.top = 6
            
        ' Tela cadastro chapas avulso
        Case "btnLImgCadastrarMaterialAvulso"
            efeitoBotoesImagem.top = 24
        Case "btnLImgCadastrarPolimentoAvulso"
            efeitoBotoesImagem.top = 78
    End Select
End Sub
'Efeito de passagem por cima do mouse
Private Sub efeitoBotoesImagem_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Efeitos paras os label botões com imagens
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
            efeitoBotoesImagem.top = 70
        Case "btnLImgCadastrarDestinoDespache"
            efeitoBotoesImagem.top = 70
        Case "btnLImgExportarListaDespache"
            efeitoBotoesImagem.top = 4
            
        ' Tela estoque m³
        Case "btnLImgExportarEstoqueM3"
            efeitoBotoesImagem.top = 16
            
        ' Tela estoque m²
        Case "btnLImgExportarEstoqueM2"
            efeitoBotoesImagem.top = 4
            
        ' Tela cadastro chapas avulso
        Case "btnLImgCadastrarMaterialAvulso"
            efeitoBotoesImagem.top = 22
        Case "btnLImgCadastrarPolimentoAvulso"
            efeitoBotoesImagem.top = 76
    End Select
End Sub
'Efeito ao clicar com o mouse
Private Sub efeitoBotoesTexto_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Efeitos paras os label botões com texto
    efeitoBotoesTexto.Font.Size = 32
    efeitoBotoesTexto.Font.Size = 16
End Sub
'Efeito de passagem por cima do mouse
Private Sub efeitoBotoesTexto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Efeitos paras os label botões com texto
    efeitoBotoesTexto.Font.Size = 18
End Sub
