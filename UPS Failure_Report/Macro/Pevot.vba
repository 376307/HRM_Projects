Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("Sheet1").Select
    Cells.Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet1!R1C1:R1048576C7", Version:=6).CreatePivotTable TableDestination:= _
        "Sheet2!R3C1", TableName:="PivotTable1", DefaultVersion:=6
    Sheets("Sheet2").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("FZM")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("FZM")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("REGION")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Status"), "Count of Status", xlCount
    ActiveWindow.SmallScroll Down:=-39
    Columns("A:B").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("1:2").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Region"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Count of UPS Failure"
    Range("B5").Select
    Columns("B:B").EntireColumn.AutoFit
    Range("A1:B1").Select
    Selection.Font.Bold = True
    Range("B1").Select
    Sheets("Sheet1").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C4").Select
    Application.CutCopyMode = False
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "Region"
    ActiveWindow.SmallScroll Down:=-18
    Range("B1").Select
    Sheets("Sheet1").Select
    ActiveWindow.SmallScroll Down:=-15
    Range("F2").Select
    ActiveWindow.SmallScroll Down:=-9
    ActiveWorkbook.Save
End Sub
