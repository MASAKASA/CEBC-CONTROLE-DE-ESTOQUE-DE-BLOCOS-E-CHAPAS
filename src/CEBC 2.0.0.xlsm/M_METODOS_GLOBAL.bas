Attribute VB_Name = "M_METODOS_GLOBAL"
' Retorna data inicial patr�o
Function dataInicial() As String
    dataInicial = "01/01/2000"
End Function

' Retorna data funal patr�o
Function dataFinal() As String
    dataFinal = "31/12/2031"
End Function

' Formata a data
Function ConverterFormatoData(data) As String
    'Variaveeis do medoto
    Dim dataOriginal As String
    Dim dataConvertida As String
    
    ' Defina a data original no formato "dd/mm/yyyy"
    dataOriginal = data
    
    ' Divida a data em dia, m�s e ano
    Dim partesData() As String
    partesData = Split(dataOriginal, "/")
    
    ' Reorganize as partes da data no formato "yyyy-mm-dd"
    dataConvertida = partesData(2) & "-" & partesData(1) & "-" & partesData(0)

    ConverterFormatoData = dataConvertida
End Function

' Verifica se a planilha existe
Function PlanilhaExiste(NomeDaPlanilha As String) As Boolean
    Dim Planilha As Worksheet
    On Error Resume Next
    Set Planilha = Worksheets(NomeDaPlanilha)
    On Error GoTo 0
    PlanilhaExiste = Not (Planilha Is Nothing)
End Function

' Conta o tamanho da matriz
Function TamanhoDaMatriz(matriz As Variant) As Integer

    'Dim minhaMatriz() As Integer ' Declare sua matriz aqui
    Dim tamanho As Integer
    
    ' Obtenha o tamanho da matriz
    tamanho = UBound(matriz) - LBound(matriz) + 1
    
    TamanhoDaMatriz = tamanho
End Function

' Inicia a contagem do regolio
Sub relogioIniciar()

    'Variaveis do metodo
    Dim dataAtual As Date
    Dim dia As Integer
    Dim mes As Integer
    Dim ano As Integer
    Dim diaSemana As String
    Dim mesExtenso As String
    
    'Captura data atual
    dataAtual = VBA.Date
    
    'Captura dia, m�s e ano
    ano = VBA.Year(dataAtual)
    mes = VBA.Month(dataAtual)
    dia = VBA.Day(dataAtual)
    
    mesExtenso = UCase(VBA.Format(dataAtual, "mmmm"))
    
    'Captura serial da data
    dataAtual = VBA.DateSerial(ano, mes, dia)
    
    'Captura dia da semana
    diaSemana = VBA.WeekdayName(VBA.Weekday(dataAtual), False)
    
    If INICIAR_RELOGIO = True Then
            
        UserFormControle.lData.Caption = UCase(diaSemana) & ", " & dia & " DE " & mesExtenso & " DE " & ano _
                & " - " & VBA.Format(VBA.Now, "hh:mm:ss")
        
        
        Application.OnTime VBA.Now + VBA.TimeValue("00:00:01"), "relogioIniciar"
    End If
End Sub

' Finaliza a contagem do relogio
Sub pararRelogio()

    INICIAR_RELOGIO = False
End Sub

' Retorna a data formatada
Function formataDataPesquisa(data) As String
    
    'Variaveis do metodo
    Dim textoDigitado As String
    Dim textoFormatado As String
    
    'Atribui a data
    textoDigitado = data
    textoFormatado = ""
    
    'Remove todos os caracteres n�o num�ricos
    For i = 1 To Len(textoDigitado)
        If IsNumeric(Mid(textoDigitado, i, 1)) Then
            textoFormatado = textoFormatado & Mid(textoDigitado, i, 1)
        End If
        
        'Coloca barra na data
        If Len(textoFormatado) = 2 Then
                
            textoFormatado = Mid(textoFormatado, 1, 2) & "/"
            
        End If
        
        If Len(textoFormatado) = 5 Then
                
            textoFormatado = Mid(textoFormatado, 1, 5) & "/"
            
        End If
    Next i
    
    'Retorno
    formataDataPesquisa = textoFormatado
    
End Function

' Formata o id da chada
Function formatarIdChapa(idBloco, tipoPolimento) As String
    
    'Variaveis do metodo
    Dim descricaoBloco As String
    Dim idChapa As String
    Dim posicaoUnderline As Integer
    Dim i As Long
    
    'Atribui��es das variaveis
    descricaoBloco = idBloco
    idChapa = ""
    posicaoUnderline = 0
    
    'Acha a posi��o do �ltimo do tra�o
    For i = Len(descricaoBloco) To 1 Step -1
        
        'Acha a posi��o do �ltimo do tra�o
        If Mid(descricaoBloco, i, 1) = "-" Then
            If posicaoUnderline = 0 Then
                posicaoUnderline = i
            End If
        End If
    Next i
    
    'Cria a id da chapa
    idChapa = Mid(descricaoBloco, 1, posicaoUnderline) & tipoPolimento
    
    formatarIdChapa = idChapa
    
End Function

' Formata o nome da chada
Function formatarNomeChapa(nomeBloco, tipoPolimento) As String
    
    'Variaveis do metodo
    Dim descricaoBloco As String
    Dim descricaoChapa As String
    Dim posicaoUnderline As Integer
    
    'Atribui��es das variaveis
    descricaoBloco = nomeBloco
    descricaoChapa = ""
    
    'Cria a id da chapa
    descricaoChapa = Mid(descricaoBloco, 7, Len(descricaoBloco)) & " " & tipoPolimento
    
    formatarNomeChapa = descricaoChapa
    
End Function

' Calcula o custo do material
Function custoBloco(valorBloco, valorFrete, valorSerrada, valorPolimento, valoresAdicionais) As Double
    
    'Variaveis do metodo
    Dim bloco As Double
    Dim frete As Double
    Dim serrada As Double
    Dim polimento As Double
    Dim adicionais As Double
    Dim custo As Double
    
    'Atribuii��es das variaveis
    bloco = CDbl(valorBloco)
    frete = CDbl(valorFrete)
    serrada = CDbl(valorSerrada)
    polimento = CDbl(valorPolimento)
    adicionais = CDbl(valoresAdicionais)
    
    'Custo do material
    custo = bloco + frete + serrada + polimento + adicionais
    
    'Retorno custo por m�
    custoBloco = custo
    
End Function

' Calcula o custo do material
Function custoMaterialM2(valorBloco, valorFrete, valoresAdicionais, valorSerrada, valorPolimento, qtdM2) As Double
    
    'Variaveis do metodo
    Dim bloco As Double
    Dim frete As Double
    Dim adicionais As Double
    Dim serrada As Double
    Dim polimento As Double
    Dim m2 As Double
    Dim custoBloco As Double
    Dim custoM2 As Double
    
    'Atribuii��es das variaveis
    bloco = CDbl(valorBloco)
    frete = CDbl(valorFrete)
    adicionais = CDbl(valoresAdicionais)
    serrada = CDbl(valorSerrada)
    polimento = CDbl(valorPolimento)
    m2 = CDbl(qtdM2)
    
    'Para evitat erro na divis�o se m� for igual a zero
    If m2 = 0 Then
        m2 = 1
    End If
    
    'Custo do material
    custoBloco = bloco + frete + adicionais + serrada + polimento
    custoM2 = custoBloco / m2
    
    'Retorno custo por m�
    custoMaterialM2 = custoM2
    
End Function

' Fun��o para retornar metros com quatro digitos ap�s a virgula
Function formatarMetros(quantidade As String)
    
    'Variaveis do metodo
    Dim textoDigitado As String
    Dim textoFormatado As String
    Dim primeiroCaractere As String
    Dim caractere_0 As String
    Dim i As Integer
    
    'Obt�m o texto digitado pelo usu�rio
    textoDigitado = quantidade
    textoFormatado = ""
    
    'Remove todos os caracteres n�o num�ricos
    For i = 1 To Len(textoDigitado)
        If IsNumeric(Mid(textoDigitado, i, 1)) Then
            textoFormatado = textoFormatado & Mid(textoDigitado, i, 1)
        End If
        
        'Remove o zero na esquerda do texto
        If Len(textoFormatado) = 6 Then
            If Left(textoFormatado, 1) = 0 Then
                textoFormatado = Mid(textoFormatado, 2, 5)
            End If
        End If
    Next i
    
    'Adiciona um caractere de virgula e mant�m a m�scara
    If Len(textoFormatado) = 6 Then
        textoFormatado = Left(textoFormatado, 1) & Mid(textoFormatado, 2, 1) & "," & Mid(textoFormatado, 3, 5)
    End If
    
    'Adiciona um caractere de virgula e mant�m a m�scara
    If Len(textoFormatado) = 5 Then
        textoFormatado = Left(textoFormatado, 1) & "," & Mid(textoFormatado, 2, 5)
    End If
    
    'Captura o primeiro caractere para compara��o
    primeiroCaractere = Mid(textoFormatado, 1, 1)

    'Adicionaos a direita e mant�m a m�scara
    If Len(textoFormatado) = 4 Then
        textoFormatado = "0," & textoFormatado

    ElseIf Len(textoFormatado) = 3 Then
        textoFormatado = "0,0" & textoFormatado

    ElseIf Len(textoFormatado) = 2 Then
        textoFormatado = "0,00" & textoFormatado

    ElseIf Len(textoFormatado) = 1 Then
        textoFormatado = "0,000" & textoFormatado

    ElseIf Len(textoFormatado) = 0 Then
        textoFormatado = "0,0000" & textoFormatado
    End If
    
    'Retorna valor formatado
    formatarMetros = textoFormatado
    
End Function

' Fun��o para retornar valor com dois digitos ap�s a virgula
Function formatarValor(valor As String)

    'Variaveis do metodo
    Dim textoDigitado As String
    Dim textoFormatado As String
    Dim textoEditado As String
    Dim primeiroCaractere As String
    Dim caractere_0 As String
    Dim posicaoVirgula As Long
    Dim i As Integer
    
    'Obt�m o texto digitado pelo usu�rio
    textoDigitado = valor
    textoFormatado = ""
    textoEditado = ""
    
    'Remove todos os caracteres n�o num�ricos
    For i = 1 To Len(textoDigitado)
        If IsNumeric(Mid(textoDigitado, i, 1)) Then
            textoFormatado = textoFormatado & Mid(textoDigitado, i, 1)
        End If
        
        'Remove o zero na esquerda do texto
        If Len(textoFormatado) = 4 Then
            If Left(textoFormatado, 1) = 0 Then
                textoFormatado = Mid(textoFormatado, 2, 3)
            End If
        End If
    Next i
    
    'Adiciona um caractere de virgula e mant�m a m�scara
    If Len(textoFormatado) = 4 Then
        textoFormatado = Left(textoFormatado, 1) & Mid(textoFormatado, 2, 1) & "," & Mid(textoFormatado, 3, 3)
    End If
    
    'Adiciona um caractere de virgula e mant�m a m�scara
    If Len(textoFormatado) = 3 Then
        textoFormatado = Left(textoFormatado, 1) & "," & Mid(textoFormatado, 2, 3)
    End If
    
    'Captura o primeiro caractere para compara��o
    primeiroCaractere = Mid(textoFormatado, 1, 1)

    'Faz a formata��o do texto se for editado
    If Len(textoFormatado) < 4 Then
        
        'Adiciona os zeros a direita e mant�m a m�scara
        If Len(textoFormatado) = 2 Then
            textoEditado = "0," & textoFormatado
            textoFormatado = textoEditado
            
        ElseIf Len(textoFormatado) = 1 Then
            textoEditado = "0,0" & textoFormatado
            textoFormatado = textoEditado
            
        ElseIf Len(textoFormatado) = 0 Then
            textoEditado = "0,00" & textoFormatado
            textoFormatado = textoEditado
            
        End If
    End If
    
    'Coloca vigula com duas casas decimais se n�o for uma edi��o
    If Len(textoFormatado) >= 5 Then
        'Captura a posi��o da virgula no texto
        posicaoVirgula = InStr(textoFormatado, ",")
        
        If posicaoVirgula = 0 Then
            textoFormatado = Mid(textoFormatado, 1, Len(textoFormatado) - 2) & "," & Mid(textoFormatado, _
                    Len(textoFormatado) - 1, 2)
            
            If Len(textoFormatado) = 7 Then
                textoFormatado = Mid(textoFormatado, 1, 1) & "." & Mid(textoFormatado, 2, 6)
            End If
        End If
    End If
    
    'Retorna valor formatado
    formatarValor = textoFormatado
End Function

' Formata e calcula a subtra��o no m� da chapa
Function subtracaoM2(m2Estoque As String, m2Despache As String) As Double
    
    'Variaveis do metodo
    Dim textoFormatado As String
    Dim estoque As Double
    Dim despache As Double
    Dim totalM2 As Double
    
    'Convertendo os valores
    estoque = CDbl(m2Estoque)
    despache = CDbl(m2Despache)
    
    'Receber e calcular o total do bloco
    totalM2 = estoque - despache
    
    'Retornar valor calcuculado e formatado
    subtracaoM2 = totalM2
    
End Function
'Formata e calcula o total do bloco
Function calcularValorBloco(totalMetro As String, valorMetro As String) As Double
    
    'Variaveis do metodo
    Dim textoFormatado As String
    Dim totalM As Double
    Dim valorM As Double
    Dim totalBloco As Double
    
    'Convertendo os valores
    totalM = CDbl(totalMetro)
    valorM = CDbl(valorMetro)
    
    'Receber e calcular o total do bloco
    totalBloco = totalM * valorM
    
    'Retornar valor calcuculado e formatado
    calcularValorBloco = totalBloco
    
End Function

' Calcula o m�
Function calcularM3(compr As String, alt As String, larg As String) As Double
    
    'Variaveis do medoto
    Dim comprimento As Double
    Dim altura As Double
    Dim largura As Double
    Dim totalM3 As Double
    
    'Convertendo os valores
    comprimento = CDbl(compr)
    altura = CDbl(alt)
    largura = CDbl(larg)
    
    'Calculando o metro m�
    totalM3 = comprimento * altura * largura
    
    'Retornando o m�
    calcularM3 = totalM3
    
End Function

' Calcula o m�
Function calcularM2(compr As String, alt As String, qtd As String) As Double
    
    'Variaveis do medoto
    Dim comprimento As Double
    Dim altura As Double
    Dim quantidade As Double
    Dim totalM2 As Double
    
    If qtd = "" Then
        qtd = 0
    End If
    'Convertendo os valores
    comprimento = CDbl(compr)
    altura = CDbl(alt)
    quantidade = CDbl(qtd)
    
    'Calculando o metro m�
    totalM2 = comprimento * altura * quantidade
    
    'Retornando o m�
    calcularM2 = totalM2
    
End Function

' Formata com pontos para melhor visualia��o
Function formatarComPontos(texto As String) As String

    'Variareis do metodo
    Dim textoFormatado As String
    Dim textoEdicao As String
    Dim numerosInteiros As String
    Dim numerosDecimais As String
    Dim posicaoVirgula As Long
    Dim i As Integer
    Dim temPonto As Boolean
    
    temPonto = False
    
    'Recebi os valores para manipula��o
    textoFormatado = ""
    textoEdicao = texto
    posicaoVirgula = InStr(textoEdicao, ",") ' Captura a posi��o da virgula no texto
    numerosInteiros = Mid(textoEdicao, 1, posicaoVirgula - 1) ' Captura os n�meros inteiros
    numerosDecimais = Mid(textoEdicao, posicaoVirgula + 1, Len(textoEdicao)) ' Captura os n�meros decimais
    
    'Coloca os pontos nos n�meros inteiros
    Select Case Len(numerosInteiros)
       Case 4
            textoFormatado = Mid(numerosInteiros, 1, 1) & "." & Mid(numerosInteiros, 2, 3) & "," & numerosDecimais
       Case 5
            
            'Percorre os caracteres para saber se tem .
            For i = 1 To Len(textoEdicao)
            
                If IsNumeric(Mid(textoEdicao, i, 1)) Then
                                   
                Else
                    If Mid(textoEdicao, i, 1) = "." Then
                        temPonto = True
                    End If
                End If
            Next i
            
            If temPonto = True Then
                textoFormatado = Mid(numerosInteiros, 1, 2) & Mid(numerosInteiros, 3, 3) & "," & numerosDecimais
            Else
                textoFormatado = Mid(numerosInteiros, 1, 2) & "." & Mid(numerosInteiros, 3, 3) & "," & numerosDecimais
            End If
            
       Case 6
            textoFormatado = Mid(numerosInteiros, 1, 3) & "." & Mid(numerosInteiros, 4, 3) & "," & numerosDecimais
       Case 7
            textoFormatado = Mid(numerosInteiros, 1, 1) & "." & Mid(numerosInteiros, 2, 3) & "." & Mid(numerosInteiros, 5, 3) & "," & numerosDecimais
       Case 8
            textoFormatado = Mid(numerosInteiros, 1, 2) & "." & Mid(numerosInteiros, 3, 3) & "." & Mid(numerosInteiros, 6, 3) & "," & numerosDecimais
       Case 9
            textoFormatado = Mid(numerosInteiros, 1, 3) & "." & Mid(numerosInteiros, 4, 3) & "." & Mid(numerosInteiros, 7, 3) & "," & numerosDecimais
       Case 10
            textoFormatado = Mid(numerosInteiros, 1, 1) & "." & Mid(numerosInteiros, 2, 3) & "." & Mid(numerosInteiros, 5, 3) & "." & Mid(numerosInteiros, 8, 3) & "," & numerosDecimais
       Case Else
           textoFormatado = Mid(numerosInteiros, 1, 3) & "," & numerosDecimais
    End Select
        
    formatarComPontos = textoFormatado

End Function

' Extrai a �ltima palavra do texto
Function ExtrairUltimaPalavra(texto As String) As String
    
    'Variaveis do metodo
    Dim palavras() As String
    Dim ultimaPalavra As String
    
    palavras = Split(texto, " ")
    If UBound(palavras) >= 0 Then
        ultimaPalavra = palavras(UBound(palavras))
    End If
    
    ExtrairUltimaPalavra = ultimaPalavra
End Function
