Sub Czyszczenie()
    Cells.Style = "Normal"
End Sub

Sub Tabela()
    Selection.Columns.AutoFit
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .Interior.Color = 255
        .Font.ThemeColor = xlThemeColorDark1
    End With
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

