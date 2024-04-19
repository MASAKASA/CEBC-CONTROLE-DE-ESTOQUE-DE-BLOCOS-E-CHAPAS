Attribute VB_Name = "serviceBloco"
Option Explicit

Dim bloco As objBloco

' Cadastra e edita objeto
Function cadastrarEEditar(bloco As objBloco)
    Call daoBloco.cadastrarEEditar(bloco)
    
End Function

' Exclui objeto
Function excluir()

End Function

' Pesquisa objeto por id
Function pesquisarPorId(id As String) As objBloco
    Set bloco = daoBloco.pesquisarPorId
    
    Set pesquisarPorId = bloco
End Function

' Pesquisa objeto por nome
Function pesquisarPorNome()

End Function

' Pesquisa objeto
Function listarBlocosFilter()
    'Call daoBloco.listarBlocosFilter
End Function
