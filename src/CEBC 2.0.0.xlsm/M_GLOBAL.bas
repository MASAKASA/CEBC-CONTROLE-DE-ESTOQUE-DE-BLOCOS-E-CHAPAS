Attribute VB_Name = "M_GLOBAL"
Option Explicit

' Variavel para controlar inicio e fim do relogio
Global INICIAR_RELOGIO As Boolean

' Variaveis para manipulação de banco de dados
Global BD As New ADODB.Connection
Global rs As New ADODB.Recordset
Global ARQ As String
Global CS As String
Global SENHA_BD As String
' Sub para abrir a conexão com banco de dados
Sub conctarBanco()

    ' SENHA_BD = "MAsa0608#"
    ' Caminho onde esta o banco de dados
    ARQ = ThisWorkbook.Path & "\BD\" & "BD_CEBC.accdb;"
    
    ' String de conexão com banco de dados
    CS = "Provider=Microsoft.ACE.OLEDB.12.0;" _
        & "Data Source=" & ARQ _
        & "Persist Security Info=False"
        
        ' Provider=Microsoft.ACE.OLEDB.12.0;Data Source=G:\Meu Drive\SEU DENIS\CEBC - CONTROLE DE ESTOQUE DE BLOCOS E CHAPAS\BD\BD_CEBC.accdb;
        ' Jet OLEDB:Database Password=MAsa0608#;
   
    ' Comando para abrir a conexão com banco de dados
    BD.Open CS
End Sub
' Sub para fechar a conexão com banco de dados
Sub fecharConexaoBanco()
    BD.Close
End Sub
'Sub para fechar Recordset
Sub fecharRecordset()
    'If de verificação se esta aberto
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
End Sub

    
