Attribute VB_Name = "ObjectFactory"
Option Explicit

' Criação da instância da rsAuxiliar
Public Function factoryRsBloco(rsBloco As ADODB.Recordset) As ADODB.Recordset
    ' Verificação se a instância já foi criada
    If Not rsBloco Is Nothing Then
        ' Verifica se a consulta está aberta
        If rsBloco.State = 1 Then
            ' A conexão está aberta não faz nada
            Exit Function
        Else
            ' Abre para conexão
            Set rsBloco = New ADODB.Recordset
        End If
    Else
        ' Abre para conexão
        Set rsBloco = New ADODB.Recordset
    End If
    ' Retorna a instância
    Set factoryRsBloco = rsBloco
End Function

' Criação da instância da rsAuxiliar
Public Function factoryRsAuxiliar(rsAuxiliar As ADODB.Recordset) As ADODB.Recordset
    ' Verificação se a instância já foi criada
    If Not rsAuxiliar Is Nothing Then
        ' Verifica se a consulta está aberta
        If rsAuxiliar.State = 1 Then
            ' A conexão está aberta não faz nada
            Exit Function
        Else
            ' Abre para conexão
            Set rsAuxiliar = New ADODB.Recordset
        End If
    Else
        ' Abre para conexão
        Set rsAuxiliar = New ADODB.Recordset
    End If
    ' Retorna a instância
    Set factoryRsAuxiliar = rsAuxiliar
End Function

' Criação da instância da lista
Public Function factoryLista(variavelLista As Collection) As Collection
    ' Verificação se a instância já foi criada
    If variavelLista Is Nothing Then
        Set variavelLista = New Collection
    End If
    ' Retorna a instância
    Set factoryLista = variavelLista
End Function

' Criação da instância de bloco
Public Function factoryBloco(variavelBloco As objBloco) As objBloco
    ' Verificação se a instância já foi criada
    If variavelBloco Is Nothing Then
        Set variavelBloco = New objBloco
    End If
    ' Retorna a instância
    Set factoryBloco = variavelBloco
End Function

' Criação da instância de chapa
Public Function factoryChapa(variavelChapa As objChapa) As objChapa
    ' Verificação se a instância já foi criada
    If variavelChapa Is Nothing Then
        Set variavelChapa = New objChapa
    End If
    ' Retorna a instância
    Set factoryChapa = variavelChapa
End Function

' Criação da instância de destino
Public Function factoryDestino(variavelDestino As objDestino) As objDestino
    ' Verificação se a instância já foi criada
    If variavelDestino Is Nothing Then
        Set variavelDestino = New objDestino
    End If
    ' Retorna a instância
    Set factoryDestino = variavelDestino
End Function

' Criação da instância de motorista
Public Function factoryMotorista(variavelMotorista As objMotoista) As objMotoista
    ' Verificação se a instância já foi criada
    If variavelMotorista Is Nothing Then
        Set variavelMotorista = New objMotoista
    End If
    ' Retorna a instância
    Set factoryMotorista = variavelMotorista
End Function

' Criação da instância de pedreira
Public Function factoryPedreira(variavelPedreira As objPedreira) As objPedreira
    ' Verificação se a instância já foi criada
    If variavelPedreira Is Nothing Then
        Set variavelPedreira = New objPedreira
    End If
    ' Retorna a instância
    Set factoryPedreira = variavelPedreira
End Function

' Criação da instância de polideira
Public Function factoryPolideira(variavelPolideira As objPolideira) As objPolideira
    ' Verificação se a instância já foi criada
    If variavelPolideira Is Nothing Then
        Set variavelPolideira = New objPolideira
    End If
    ' Retorna a instância
    Set factoryPolideira = variavelPolideira
End Function

' Criação da instância de serraria
Public Function factorySerraria(variavelSerraria As objSerraria) As objSerraria
    ' Verificação se a instância já foi criada
    If variavelSerraria Is Nothing Then
        Set variavelSerraria = New objSerraria
    End If
    ' Retorna a instância
    Set factorySerraria = variavelSerraria
End Function

' Criação da instância de status
Public Function factoryStatus(variavelStatus As objStatus) As objStatus
    ' Verificação se a instância já foi criada
    If variavelStatus Is Nothing Then
        Set variavelStatus = New objStatus
    End If
    ' Retorna a instância
    Set factoryStatus = variavelStatus
End Function

' Criação da instância de tamanho
Public Function factoryTamanho(variavelTamanho As objTamanho) As objTamanho
    ' Verificação se a instância já foi criada
    If variavelTamanho Is Nothing Then
        Set variavelTamanho = New objTamanho
    End If
    ' Retorna a instância
    Set factoryTamanho = variavelTamanho
End Function

' Criação da instância de tamanho
Public Function factoryTipoMaterial(variavelTipoMaterial As objTipoMaterial) As objTipoMaterial
    ' Verificação se a instância já foi criada
    If variavelTipoMaterial Is Nothing Then
        Set variavelTipoMaterial = New objTipoMaterial
    End If
    ' Retorna a instância
    Set factoryTipoMaterial = variavelTipoMaterial
End Function

' Criação da instância de tamanho
Public Function factoryTipoPolimento(variavelTipoPolimento As objTipoPolimento) As objTipoPolimento
    ' Verificação se a instância já foi criada
    If variavelTipoPolimento Is Nothing Then
        Set variavelTipoPolimento = New objTipoPolimento
    End If
    ' Retorna a instância
    Set factoryTipoPolimento = variavelTipoPolimento
End Function

' Criação da instância de estoque blocos
Public Function factoryEstoque(variavelEstoque As objEstoque) As objEstoque
    ' Verificação se a instância já foi criada
    If variavelEstoque Is Nothing Then
        Set variavelEstoque = New objEstoque
    End If
    ' Retorna a instância
    Set factoryEstoque = variavelEstoque
End Function

' Criação da instância de estoque de chapas
Public Function factoryEstoqueChapas(variavelEstoque As objEstoqueChapa) As objEstoque
    ' Verificação se a instância já foi criada
    If variavelEstoque Is Nothing Then
        Set variavelEstoque = New objEstoque
    End If
    ' Retorna a instância
    Set factoryEstoqueChapas = variavelEstoque
End Function
