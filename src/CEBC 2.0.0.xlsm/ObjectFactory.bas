Attribute VB_Name = "ObjectFactory"
Option Explicit

Private bloco As objBloco
Private chapa As objChapa
Private destino As objDestino
Private motorista As objMotoista
Private pedreira As objPedreira
Private polideira As objPolideira
Private serraria As objSerraria
Private status As objStatus
Private tamanho As objTamanho
Private tipoMaterial As objTipoMaterial
Private tipoPolimento As objTipoPolimento
Private lista As Collection

' Cria��o da inst�ncia da lista
Public Function factoryLista(variavelLista As Collection) As Collection
    ' Verifica��o se a inst�ncia j� foi criada
    If lista Is Nothing Then
        Set lista = New Collection
    End If
    
    ' Retorna a inst�ncia
    Set factoryLista = lista
End Function
' Cria��o da inst�ncia de bloco
Public Function factoryBloco(variavelBloco As objBloco) As objBloco
    ' Verifica��o se a inst�ncia j� foi criada
    If bloco Is Nothing Then
        Set bloco = New objBloco
    End If
    
    ' Retorna a inst�ncia
    Set factoryBloco = bloco
End Function

' Cria��o da inst�ncia de chapa
Public Function factoryChapa(variavelChapa As objChapa) As objChapa
    ' Verifica��o se a inst�ncia j� foi criada
    If chapa Is Nothing Then
        Set chapa = New objChapa
    End If
    
    ' Retorna a inst�ncia
    Set factoryChapa = chapa
End Function

' Cria��o da inst�ncia de destino
Public Function factoryDestino(variavelDestino As objDestino) As objDestino
    ' Verifica��o se a inst�ncia j� foi criada
    If destino Is Nothing Then
        Set destino = New objDestino
    End If
    
    ' Retorna a inst�ncia
    Set factoryDestino = destino
End Function

' Cria��o da inst�ncia de motorista
Public Function factoryMotorista(variavelMotorista As objMotoista) As objMotoista
    ' Verifica��o se a inst�ncia j� foi criada
    If motorista Is Nothing Then
        Set motorista = New objMotoista
    End If
    
    ' Retorna a inst�ncia
    Set factoryMotorista = motorista
End Function

' Cria��o da inst�ncia de pedreira
Public Function factoryPedreira(variavelPedreira As objPedreira) As objPedreira
    ' Verifica��o se a inst�ncia j� foi criada
    If pedreira Is Nothing Then
        Set pedreira = New objPedreira
    End If
    
    ' Retorna a inst�ncia
    Set factoryPedreira = pedreira
End Function

' Cria��o da inst�ncia de polideira
Public Function factoryPolideira(variavelPolideira As objPolideira) As objPolideira
    ' Verifica��o se a inst�ncia j� foi criada
    If polideira Is Nothing Then
        Set polideira = New objPolideira
    End If
    
    ' Retorna a inst�ncia
    Set factoryPolideira = polideira
End Function

' Cria��o da inst�ncia de serraria
Public Function factorySerraria(variavelSerraria As objSerraria) As objSerraria
    ' Verifica��o se a inst�ncia j� foi criada
    If serraria Is Nothing Then
        Set serraria = New objSerraria
    End If
    
    ' Retorna a inst�ncia
    Set factorySerraria = serraria
End Function

' Cria��o da inst�ncia de status
Public Function factoryStatus(variavelStatus As objStatus) As objStatus
    ' Verifica��o se a inst�ncia j� foi criada
    If status Is Nothing Then
        Set status = New objStatus
    End If
    
    ' Retorna a inst�ncia
    Set factoryStatus = status
End Function

' Cria��o da inst�ncia de tamanho
Public Function factoryTamanho(variavelTamanho As objTamanho) As objTamanho
    ' Verifica��o se a inst�ncia j� foi criada
    If tamanho Is Nothing Then
        Set tamanho = New objTamanho
    End If
    
    ' Retorna a inst�ncia
    Set factoryTamanho = tamanho
End Function

' Cria��o da inst�ncia de tamanho
Public Function factoryTipoMaterial(variavelTipoMaterial As objTipoMaterial) As objTipoMaterial
    ' Verifica��o se a inst�ncia j� foi criada
    If tipoMaterial Is Nothing Then
        Set tipoMaterial = New objTipoMaterial
    End If
    
    ' Retorna a inst�ncia
    Set factoryTipoMaterial = tipoMaterial
End Function

' Cria��o da inst�ncia de tamanho
Public Function factoryTipoPolimento(variavelTipoPolimento As objTipoPolimento) As objTipoPolimento
    ' Verifica��o se a inst�ncia j� foi criada
    If tipoPolimento Is Nothing Then
        Set tipoPolimento = New objTipoPolimento
    End If
    
    ' Retorna a inst�ncia
    Set factoryTipoPolimento = tipoPolimento
End Function
