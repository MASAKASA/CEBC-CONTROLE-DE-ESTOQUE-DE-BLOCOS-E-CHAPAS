Attribute VB_Name = "M_CONEXAO_BD"
Option Explicit

' Abrir a conexão com banco de dados
Sub conctarBanco()
    ' Variaveis do metodo
    Dim caminhoBD As String
    Dim cs As String
    Dim senhaBD As String
       
    ' senhaBD = "MAsa0608#"
    ' Caminho onde esta o banco de dados
    caminhoBD = ThisWorkbook.Path & "\BD\" & "BD_CEBC_2.0.0.accdb;"
    
    ' String de conexão com banco de dados
    cs = "Provider=Microsoft.ACE.OLEDB.12.0;" _
        & "Data Source=" & caminhoBD _
        & "Persist Security Info=False"
        
        ' Provider=Microsoft.ACE.OLEDB.12.0;Data Source=G:\Meu Drive\SEU DENIS\CEBC - CONTROLE DE ESTOQUE DE BLOCOS E CHAPAS\BD\BD_CEBC.accdb;
        ' Jet OLEDB:Database Password=MAsa0608#;
        
    ' Verifica se a conexão já existe
    If Not CONEXAO_BD Is Nothing Then
        ' Verifica se a conexão está aberta
        If CONEXAO_BD.State = 1 Then ' 1 representa adStateOpen
            ' A conexão está aberta não faz nada
            Exit Sub
        Else
            ' Abre a conexão com banco de dados
            CONEXAO_BD.Open cs
        End If
    Else
        ' Abre a conexão com banco de dados
        CONEXAO_BD.Open cs
    End If
End Sub

' Fechar a conexão com banco de dados
Sub fecharConexaoBanco()
    ' Verifica se a conexão já existe
    If Not CONEXAO_BD Is Nothing Then
        ' Verifica se a conexão está aberta
        If CONEXAO_BD.State = 1 Then ' 1 representa adStateOpen
            ' Fecha a conexão com banco de dados
            CONEXAO_BD.Close
            ' Libera espaço na memoria
            Set CONEXAO_BD = Nothing
        End If
    End If
End Sub


    
