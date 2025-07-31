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

Public id As Long
Public dataDespache As String
Public despachado As String
Public qtdChapas As Integer
Public motorista As objMotorista
Public destino As objDestino
Public listaChapas As Collection

' Construtor para cadastro
Public Sub carregarDespacheCadastro(dataDespache As String, despachado As String, qtdChapas As Integer, _
                                    motorista As objMotorista, destino As objDestino, listaChapas As Collection)

    ' Atribuições
    Me.dataDespache = dataDespache
    Me.despachado = despachado
    Me.qtdChapas = qtdChapas
    Set Me.motorista = motorista
    Set Me.destino = destino
    Set Me.listaChapas = listaChapas
End Sub

' Construtor para manipulação
Public Sub carregarDespacheManipulacao(id As Long, dataDespache As String, qtdChapas As Integer, _
                                    motorista As objMotorista, destino As objDestino, listaChapas As Collection)

    ' Atribuições
    Me.id = id
    Me.dataDespache = dataDespache
    Me.despachado = despachado
    Me.qtdChapas = qtdChapas
    Set Me.motorista = motorista
    Set Me.destino = destino
    Set Me.listaChapas = listaChapas
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

Public Function getListaChapas() As Collection
    Set getListaChapas = Me.listaChapas
End Function

Public Function setListaChapas(listaChapas As Collection)
    Set Me.listaChapas = listaChapas
End Function


