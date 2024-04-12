Attribute VB_Name = "serviceBloco"
Option Explicit

' Cadastra e edita objeto
Function cadastrarEEditar()

End Function

' Exclui objeto
Function excluir()

End Function

' Pesquisa objeto por id
Function pesquisarPorId(id As String) As objBloco
    Call daoBloco.pesquisarPorId
    
    'pesquisarPorId
End Function

' Pesquisa objeto por nome
Function pesquisarPorNome()

End Function

' Pesquisa objeto
Function listarBlocosFilter()
    Call daoBloco.listarBlocosFilter
End Function
