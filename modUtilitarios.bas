Attribute VB_Name = "modUtilitarios"
Option Explicit

' Funções pequenas usadas pelos outros módulos.

Public Sub IrParaCadastro()
    With ThisWorkbook.Worksheets("Cadastro")
        .Activate
        .Range("C4").Select
    End With
End Sub

Public Sub IrParaDashboard()
    ThisWorkbook.Worksheets("Dashboard").Activate
End Sub

Public Function ContarRegistros() As Long
    ContarRegistros = ThisWorkbook.Worksheets("BaseDados").ListObjects("TabelaVendas").ListRows.Count
End Function

Public Function FormatarMoeda(ByVal valor As Double) As String
    FormatarMoeda = Format(valor, "R$ #,##0.00")
End Function
