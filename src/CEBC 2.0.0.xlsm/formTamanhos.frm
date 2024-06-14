VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formTamanhos 
   Caption         =   "TAMANHOS CHAPA"
   ClientHeight    =   8595.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13665
   OleObjectBlob   =   "formTamanhos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formTamanhos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Variaveis para manipulação com os botões e frames
Dim botoesText() As clsLabel
Dim frameEfeito() As clsFrame
Dim errorStyle As clsErrorStyle

' Inicialização do form
Private Sub UserForm_Initialize()

    ' Variaveis para o metodo
    Dim obj As Object
    Dim i As Long
    Dim j As Long
    Dim l As Long
    Dim m As Long
    Dim nameObj As String
    Dim nameObjInicio As String
    Dim chapaTamanho As objChapa
    
    ' Carrega tradução do sistema
'    Call M_TRADUCAO.carregarTraducaoErros
    
    ' Resevando espaço em memoria para manipulação das variaveis
    ReDim botoesText(1 To Me.Controls.Count)
    ReDim frameEfeito(1 To Me.Controls.Count)
    Set errorStyle = New clsErrorStyle
    
    ' Separa os botões e frames
    For Each obj In Me.Controls
        
        ' Atribuições das variaveis para manipulações
        nameObj = obj.name
        nameObjInicio = Mid(nameObj, 1, 7)
        
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
    ReDim Preserve botoesText(1 To l)
    ReDim Preserve frameEfeito(1 To m)
    
    Set chapaTamanho = formControle.chapa
    
    ' Carrega os campos
    txtMaterial.Value = formControle.chapa.nomeMaterial
    txtIdChapaSitema.Value = formControle.chapa.idSistema
    txtNumeroPedreira.Value = formControle.chapa.numeroBlocoPedreira
    txtIdBlocoSistema.Value = formControle.chapa.bloco.idSistema
    cblTipoPoliment.AddItem formControle.chapa.tipoPolimento.nome
    cblTipoPoliment.ListIndex = 0
    txtTotal.Value = formControle.chapa.valorTotal
    
    Call carregarListTamanhosChapas(ListTamanhosChapas, formControle.chapa.tamanhos)
    
    If formControle.paginaAnterior = 4 Then
        btnLTxtAdicionarTamanho.Visible = False
    End If
End Sub

' Seta os valores para ser alterados
Private Sub btnLTxtTrocar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    ' Variaveis no metodo
    Dim listaPolimentosJaCadastras As Collection
    Dim tamanhoChapa As objTamanho
    Dim chapaPesquisa As objChapa
    
    ' Verifica se tem algum dado a pesquisa
    If Me.ListTamanhosChapas.ListIndex = -1 Then ' Se não tiver dados
        ' Mensagem para selecioanar um tamanho
        errorStyle.Informativo SELECIONE_TEM_MENSAGEM, SELECIONE_TEM_TITULO
        Exit Sub
    End If
    
    ' Seta tamanho e tipos de polimentos
    Set tamanhoChapa = daoTamanho.pesquisarPorIdTamanho(Me.ListTamanhosChapas.list(Me.ListTamanhosChapas.ListIndex, 8))
    Set listaPolimentosJaCadastras = ObjectFactory.factoryLista(listaPolimentosJaCadastras)
    
    ' Seta o ojeto
    Set chapaPesquisa = daoChapa.pesquisarPorId(txtIdChapaSitema.Value)
    ' Seta os polimentos já cadastrados
    listaPolimentosJaCadastras.Add chapaPesquisa.tipoPolimento.nome
 
    ' Carrega só os tipos deferentes
    Call carregarTiposPolimentoAlgum(formControle.cbTipoPolimentoTroca, listaPolimentosJaCadastras)
    
    ' Carrega os campos
    formControle.txtMaterialParaTroca01.Value = formControle.chapa.nomeMaterial
    formControle.txtEspParaTroca01.Value = tamanhoChapa.espessura
    formControle.txtTipoMaterialParaTroca01.Value = tamanhoChapa.tipoMaterial.nome
    formControle.txtCompMaterialParaTroca.Value = tamanhoChapa.compremento
    formControle.txtAltMaterialParaTroca.Value = tamanhoChapa.altura
    formControle.txtQtdDispovelMaterialParaTroca01.Value = tamanhoChapa.qtdEstoque
    formControle.txtTotalM2T.Value = tamanhoChapa.qtdM2
    
    ' Seta tamanho
    Set formControle.tamanho = daoTamanho.pesquisarPorIdTamanho(Me.ListTamanhosChapas.list(Me.ListTamanhosChapas.ListIndex, 8))
    
    ' Fecha formulario
    Unload formTamanhos
End Sub
' Fecha formolario
Private Sub btnLTxtVoltarTamanhos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ' Muda abra da multPage
    formControle.MultiPageCEBC.Value = formControle.paginaAnterior
    ' Fecha formulario
    Unload formTamanhos
End Sub

' Efeito de pasagem do mouse
Private Sub fTiraEfeitoBotoesAcoesChapaTamanhos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    btnLTxtAdicionarTamanho.Font.Size = 16
    btnLTxtVoltarTamanhos.Font.Size = 16
End Sub

' Carrega a lista ListTamanhosChapas tela edicao chapa
Private Sub carregarListTamanhosChapas(ListBox As MSForms.ListBox, listaCollection As Collection)
    'Variaveis do metodo
    Dim tamanho As objTamanho
    Dim qtdChapas As Integer
    Dim i As Integer
    
    ' Limpar a ListBox
    ListBox.Clear
    
    ' NOME CABEÇALHO TAMANHOS     | TIPO | COMP  | ALT   | M²    | ESP | CUSTO | POLIDEIRA | ESTOQUE
    ' Tamanho do cabeçalho left   | 7    | 157,5 | 208,5 | 259,5 | 315 | 351   | 431,5     | 551
    ' Tamanho do cabeçalho width  | 150  | 50    | 50    | 55    | 35  | 80    | 127       | 114,5
    ' Tamanho das colunas da list
    ListBox.ColumnWidths = "150;50;50;55;35;80;118,5;48"
    
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
            ListBox.list(ListBox.ListCount - 1, 4) = tamanho.espessura
            ListBox.list(ListBox.ListCount - 1, 5) = _
                                M_METODOS_GLOBAL.formatarComPontos(Format(tamanho.valorPolimento, "currency"))
            ListBox.list(ListBox.ListCount - 1, 6) = tamanho.polideira.nome
            ListBox.list(ListBox.ListCount - 1, 7) = tamanho.qtdEstoque
            ListBox.list(ListBox.ListCount - 1, 8) = tamanho.id
            
            ' Soma total de chapa da pesquisa
            qtdChapas = qtdChapas + CInt(tamanho.qtdEstoque)
            ' Libera espaço na memoria
            Set tamanho = Nothing
        Next i
    End If
    
    ' Seta qtd total
    txtEstoque.Value = qtdChapas
End Sub

' Carrega a combobox de tipo polimento com algum tipos
Private Sub carregarTiposPolimentoAlgum(cbTiposPolimento As MSForms.comboBox, lista As Collection)
    ' Variaveis do metodo
    Dim listaObjetos As Collection
    Dim tipoPolimento As objTipoPolimento
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
