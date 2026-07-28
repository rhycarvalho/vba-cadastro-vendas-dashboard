Attribute VB_Name = "modDashboard"
Option Explicit

' Reordena a base por data e força o recálculo do dashboard
' (KPIs e gráficos já são fórmula, então normalmente nem precisaria,
' mas garante que tudo fica coerente depois de vários cadastros).

Public Sub AtualizarDashboard()

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("BaseDados").ListObjects("TabelaVendas")

    OrdenarPorData tbl
    Application.CalculateFullRebuild

    ThisWorkbook.Worksheets("Dashboard").Activate
    MsgBox "Dashboard atualizado.", vbInformation, "Dashboard de Vendas"

End Sub

Private Sub OrdenarPorData(ByVal tbl As ListObject)

    If tbl.ListRows.Count = 0 Then Exit Sub

    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=tbl.ListColumns("Data").Range, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

End Sub
