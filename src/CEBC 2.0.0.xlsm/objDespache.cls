VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objDespache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As String
Public dataDespache As String
Public qtd As String
Public motorista As objMotorista
Public destino As objDestino
Public chapa As objChapa
Public tamanho As objTamanho

' Construtor
Public Sub carregarDespache(id As String, dataDespache As String, qtd As String, motorista As objMotorista, destino As objDestino, _
                            chapa As objChapa, tamanho As objTamanho)

    ' Atribuições
    Me.id = id
    Me.dataDespache = dataDespache
    Me.qtd = qtd
    Set Me.motorista = motorista
    Set Me.destino = destino
    Set Me.chapa = chapa
    Set Me.tamanho = tamanho
End Sub

' Gets e sets
Public Function getMotorista() As objMotorista
    Set getMotorista = Me.motorista
End Function

Public Function setMotorista(motorista As objMotorista)
    Set Me.motorista = motorista
End Function

Public Function getDestino() As objDestino
    Set getDestino = Me.destino
End Function

Public Function setDestino(destino As objDestino)
    Set Me.destino = destino
End Function

Public Function getChapa() As objChapa
    Set getChapa = Me.chapa
End Function

Public Function setChapa(chapa As objChapa)
    Set Me.chapa = chapa
End Function

Public Function getTamanho() As Collection
    Set getTamanho = Me.tamanho
End Function

Public Function setTamanho(tamanho As objTamanho)
    Set Me.tamanho = tamanho
End Function


