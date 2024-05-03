Attribute VB_Name = "ObjectFactory"
Option Explicit

' Cria��o da inst�ncia da rsAuxiliar
Public Function factoryRsBloco(rsBloco As ADODB.Recordset) As ADODB.Recordset
    ' Verifica��o se a inst�ncia j� foi criada
    If Not rsBloco Is Nothing Then
        ' Verifica se a consulta est� aberta
        If rsBloco.State = 1 Then
            ' A conex�o est� aberta n�o faz nada
            Exit Function
        Else
            ' Abre para conex�o
            Set rsBloco = New ADODB.Recordset
        End If
    Else
        ' Abre para conex�o
        Set rsBloco = New ADODB.Recordset
    End If
    ' Retorna a inst�ncia
    Set factoryRsBloco = rsBloco
End Function

' Cria��o da inst�ncia da rsAuxiliar
Public Function factoryRsAuxiliar(rsAuxiliar As ADODB.Recordset) As ADODB.Recordset
    ' Verifica��o se a inst�ncia j� foi criada
    If Not rsAuxiliar Is Nothing Then
        ' Verifica se a consulta est� aberta
        If rsAuxiliar.State = 1 Then
            ' A conex�o est� aberta n�o faz nada
            Exit Function
        Else
            ' Abre para conex�o
            Set rsAuxiliar = New ADODB.Recordset
        End If
    Else
        ' Abre para conex�o
        Set rsAuxiliar = New ADODB.Recordset
    End If
    ' Retorna a inst�ncia
    Set factoryRsAuxiliar = rsAuxiliar
End Function

' Cria��o da inst�ncia da lista
Public Function factoryLista(variavelLista As Collection) As Collection
    ' Verifica��o se a inst�ncia j� foi criada
    If variavelLista Is Nothing Then
        Set variavelLista = New Collection
    End If
    ' Retorna a inst�ncia
    Set factoryLista = variavelLista
End Function

' Cria��o da inst�ncia de bloco
Public Function factoryBloco(variavelBloco As objBloco) As objBloco
    ' Verifica��o se a inst�ncia j� foi criada
    If variavelBloco Is Nothing Then
        Set variavelBloco = New objBloco
    End If
    ' Retorna a inst�ncia
    Set factoryBloco = variavelBloco
End Function

' Cria��o da inst�ncia de chapa
Public Function factoryChapa(variavelChapa As objChapa) As objChapa
    ' Verifica��o se a inst�ncia j� foi criada
    If variavelChapa Is Nothing Then
        Set variavelChapa = New objChapa
    End If
    ' Retorna a inst�ncia
    Set factoryChapa = variavelChapa
End Function

' Cria��o da inst�ncia de destino
Public Function factoryDestino(variavelDestino As objDestino) As objDestino
    ' Verifica��o se a inst�ncia j� foi criada
    If variavelDestino Is Nothing Then
        Set variavelDestino = New objDestino
    End If
    ' Retorna a inst�ncia
    Set factoryDestino = variavelDestino
End Function

' Cria��o da inst�ncia de motorista
Public Function factoryMotorista(variavelMotorista As objMotoista) As objMotoista
    ' Verifica��o se a inst�ncia j� foi criada
    If variavelMotorista Is Nothing Then
        Set variavelMotorista = New objMotoista
    End If
    ' Retorna a inst�ncia
    Set factoryMotorista = variavelMotorista
End Function

' Cria��o da inst�ncia de pedreira
Public Function factoryPedreira(variavelPedreira As objPedreira) As objPedreira
    ' Verifica��o se a inst�ncia j� foi criada
    If variavelPedreira Is Nothing Then
        Set variavelPedreira = New objPedreira
    End If
    ' Retorna a inst�ncia
    Set factoryPedreira = variavelPedreira
End Function

' Cria��o da inst�ncia de polideira
Public Function factoryPolideira(variavelPolideira As objPolideira) As objPolideira
    ' Verifica��o se a inst�ncia j� foi criada
    If variavelPolideira Is Nothing Then
        Set variavelPolideira = New objPolideira
    End If
    ' Retorna a inst�ncia
    Set factoryPolideira = variavelPolideira
End Function

' Cria��o da inst�ncia de serraria
Public Function factorySerraria(variavelSerraria As objSerraria) As objSerraria
    ' Verifica��o se a inst�ncia j� foi criada
    If variavelSerraria Is Nothing Then
        Set variavelSerraria = New objSerraria
    End If
    ' Retorna a inst�ncia
    Set factorySerraria = variavelSerraria
End Function

' Cria��o da inst�ncia de status
Public Function factoryStatus(variavelStatus As objStatus) As objStatus
    ' Verifica��o se a inst�ncia j� foi criada
    If variavelStatus Is Nothing Then
        Set variavelStatus = New objStatus
    End If
    ' Retorna a inst�ncia
    Set factoryStatus = variavelStatus
End Function

' Cria��o da inst�ncia de tamanho
Public Function factoryTamanho(variavelTamanho As objTamanho) As objTamanho
    ' Verifica��o se a inst�ncia j� foi criada
    If variavelTamanho Is Nothing Then
        Set variavelTamanho = New objTamanho
    End If
    ' Retorna a inst�ncia
    Set factoryTamanho = variavelTamanho
End Function

' Cria��o da inst�ncia de tamanho
Public Function factoryTipoMaterial(variavelTipoMaterial As objTipoMaterial) As objTipoMaterial
    ' Verifica��o se a inst�ncia j� foi criada
    If variavelTipoMaterial Is Nothing Then
        Set variavelTipoMaterial = New objTipoMaterial
    End If
    ' Retorna a inst�ncia
    Set factoryTipoMaterial = variavelTipoMaterial
End Function

' Cria��o da inst�ncia de tamanho
Public Function factoryTipoPolimento(variavelTipoPolimento As objTipoPolimento) As objTipoPolimento
    ' Verifica��o se a inst�ncia j� foi criada
    If variavelTipoPolimento Is Nothing Then
        Set variavelTipoPolimento = New objTipoPolimento
    End If
    ' Retorna a inst�ncia
    Set factoryTipoPolimento = variavelTipoPolimento
End Function

' Cria��o da inst�ncia de estoque blocos
Public Function factoryEstoque(variavelEstoque As objEstoque) As objEstoque
    ' Verifica��o se a inst�ncia j� foi criada
    If variavelEstoque Is Nothing Then
        Set variavelEstoque = New objEstoque
    End If
    ' Retorna a inst�ncia
    Set factoryEstoque = variavelEstoque
End Function

' Cria��o da inst�ncia de estoque de chapas
Public Function factoryEstoqueChapas(variavelEstoque As objEstoqueChapa) As objEstoque
    ' Verifica��o se a inst�ncia j� foi criada
    If variavelEstoque Is Nothing Then
        Set variavelEstoque = New objEstoque
    End If
    ' Retorna a inst�ncia
    Set factoryEstoqueChapas = variavelEstoque
End Function
