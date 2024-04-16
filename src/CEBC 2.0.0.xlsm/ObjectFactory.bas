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

' Criação da instância da lista
Public Function factoryLista(variavelLista As Collection) As Collection
    ' Verificação se a instância já foi criada
    If lista Is Nothing Then
        Set lista = New Collection
    End If
    
    ' Retorna a instância
    Set factoryLista = lista
End Function
' Criação da instância de bloco
Public Function factoryBloco(variavelBloco As objBloco) As objBloco
    ' Verificação se a instância já foi criada
    If bloco Is Nothing Then
        Set bloco = New objBloco
    End If
    
    ' Retorna a instância
    Set factoryBloco = bloco
End Function

' Criação da instância de chapa
Public Function factoryChapa(variavelChapa As objChapa) As objChapa
    ' Verificação se a instância já foi criada
    If chapa Is Nothing Then
        Set chapa = New objChapa
    End If
    
    ' Retorna a instância
    Set factoryChapa = chapa
End Function

' Criação da instância de destino
Public Function factoryDestino(variavelDestino As objDestino) As objDestino
    ' Verificação se a instância já foi criada
    If destino Is Nothing Then
        Set destino = New objDestino
    End If
    
    ' Retorna a instância
    Set factoryDestino = destino
End Function

' Criação da instância de motorista
Public Function factoryMotorista(variavelMotorista As objMotoista) As objMotoista
    ' Verificação se a instância já foi criada
    If motorista Is Nothing Then
        Set motorista = New objMotoista
    End If
    
    ' Retorna a instância
    Set factoryMotorista = motorista
End Function

' Criação da instância de pedreira
Public Function factoryPedreira(variavelPedreira As objPedreira) As objPedreira
    ' Verificação se a instância já foi criada
    If pedreira Is Nothing Then
        Set pedreira = New objPedreira
    End If
    
    ' Retorna a instância
    Set factoryPedreira = pedreira
End Function

' Criação da instância de polideira
Public Function factoryPolideira(variavelPolideira As objPolideira) As objPolideira
    ' Verificação se a instância já foi criada
    If polideira Is Nothing Then
        Set polideira = New objPolideira
    End If
    
    ' Retorna a instância
    Set factoryPolideira = polideira
End Function

' Criação da instância de serraria
Public Function factorySerraria(variavelSerraria As objSerraria) As objSerraria
    ' Verificação se a instância já foi criada
    If serraria Is Nothing Then
        Set serraria = New objSerraria
    End If
    
    ' Retorna a instância
    Set factorySerraria = serraria
End Function

' Criação da instância de status
Public Function factoryStatus(variavelStatus As objStatus) As objStatus
    ' Verificação se a instância já foi criada
    If status Is Nothing Then
        Set status = New objStatus
    End If
    
    ' Retorna a instância
    Set factoryStatus = status
End Function

' Criação da instância de tamanho
Public Function factoryTamanho(variavelTamanho As objTamanho) As objTamanho
    ' Verificação se a instância já foi criada
    If tamanho Is Nothing Then
        Set tamanho = New objTamanho
    End If
    
    ' Retorna a instância
    Set factoryTamanho = tamanho
End Function

' Criação da instância de tamanho
Public Function factoryTipoMaterial(variavelTipoMaterial As objTipoMaterial) As objTipoMaterial
    ' Verificação se a instância já foi criada
    If tipoMaterial Is Nothing Then
        Set tipoMaterial = New objTipoMaterial
    End If
    
    ' Retorna a instância
    Set factoryTipoMaterial = tipoMaterial
End Function

' Criação da instância de tamanho
Public Function factoryTipoPolimento(variavelTipoPolimento As objTipoPolimento) As objTipoPolimento
    ' Verificação se a instância já foi criada
    If tipoPolimento Is Nothing Then
        Set tipoPolimento = New objTipoPolimento
    End If
    
    ' Retorna a instância
    Set factoryTipoPolimento = tipoPolimento
End Function
