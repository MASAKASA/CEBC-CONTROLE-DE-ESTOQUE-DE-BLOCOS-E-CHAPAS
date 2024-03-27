VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formLogin 
   Caption         =   "LOGIN"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "formLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lLogologin_Click()

End Sub

'Efeito de label tela login
Private Sub lUsuarioExemplo_Click()
    LUsuario.Visible = True
    lUsuarioExemplo.Visible = False
    txtUsuario.SetFocus
End Sub
'Efeito e coloca em caixa alta o texto em txtUsuario tela login
Private Sub txtUsuario_Change()
    LUsuario.Visible = True
    lUsuarioExemplo.Visible = False
    
    If txtUsuario.Value = "" Then
        LUsuario.Visible = False
        lUsuarioExemplo.Visible = True
    End If
    
    txtUsuario.Value = UCase(txtUsuario.Value)
End Sub
'Efeito ao sair da caixa txtUsuario de texto tela login
Private Sub txtUsuario_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtUsuario.Value = "" Then
        LUsuario.Visible = False
        lUsuarioExemplo.Visible = True
    End If
End Sub
'Efeito de label tela login
Private Sub lSenhaExemplo_Click()
    lSenha.Visible = True
    lSenhaExemplo.Visible = False
    txtSenha.SetFocus
End Sub
'Efeito e coloca em caixa alta o texto em txtSenha tela login
Private Sub txtSenha_Change()
    lSenha.Visible = True
    lSenhaExemplo.Visible = False
    
    If txtSenha.Value = "" Then
        lSenha.Visible = False
        lSenhaExemplo.Visible = True
    End If
    
    txtSenha.Value = UCase(txtSenha.Value)
End Sub
'Efeito para as imagem e colocar * para senha
Private Sub txtSenha_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    txtSenha.PasswordChar = "*"
    
    imgLCadeadoFechado.Visible = False
    imgLCadeadoAberto.Visible = True
End Sub
'Efeito para as imagem e aparecer a senha
Private Sub txtSenha_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    txtSenha.PasswordChar = ""
    
    imgLCadeadoFechado.Visible = True
    imgLCadeadoAberto.Visible = False
End Sub
'Efeito para as imagem e voltar com os *para senha
Private Sub txtSenha_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    txtSenha.PasswordChar = "*"
    
    imgLCadeadoFechado.Visible = False
    imgLCadeadoAberto.Visible = True
End Sub
'Efeito ao sair da caixa txtSenha de texto tela login
Private Sub txtSenha_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtSenha.Value = "" Then
        lSenha.Visible = False
        lSenhaExemplo.Visible = True
    End If
End Sub
'Botão login tela login
Private Sub btnLEntrar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Efeito ao botão
    btnLEntrar.Font.Size = 32
    btnLEntrar.Font.Size = 16
End Sub
'Efeito de passagem do mouse botão entrar tela login
Private Sub btnLEntrar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Efeito ao botão
    btnLEntrar.Font.Size = 18
End Sub
'Ação para quando o mouse sair de cima dos botões
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Efeito ao botão
    btnLEntrar.Font.Size = 16
End Sub




