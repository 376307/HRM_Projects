Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("BRANCH OPENING SUMMARY|FZM WISE").Select
    Columns("A:A").EntireColumn.AutoFit
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("B4").Select
    Rows("1:1").RowHeight = 33
    Columns("B:B").ColumnWidth = 14.14
    Range("A1:F1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("C:C").ColumnWidth = 21.14
    Columns("D:D").ColumnWidth = 17.43
    Columns("E:E").ColumnWidth = 19.71
    Columns("E:E").ColumnWidth = 17.29
    Columns("F:F").ColumnWidth = 12.57
    Range("A1:F10").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:F1").Select
    Selection.Font.Bold = True
    Range("A10:F10").Select
    Selection.Font.Bold = True
    Range("B1").Select
    Sheets("BRANCH EMPLOYEE PUNCHING STATUS").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1:F1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows("1:1").RowHeight = 30.75
    Columns("B:B").ColumnWidth = 12.57
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("C:C").ColumnWidth = 14.29
    Columns("D:D").ColumnWidth = 14
    Columns("E:E").ColumnWidth = 33.14
    Columns("F:F").ColumnWidth = 12.57
    Columns("F:F").ColumnWidth = 13.43
    Columns("F:F").ColumnWidth = 14.57
    Selection.Font.Bold = True
    Range("A10").Select
    ActiveCell.FormulaR1C1 = "Grand Total"
    Range("A10:F10").Select
    Selection.Font.Bold = True
    Range("E10").Select
    Sheets("REGION REPORT").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Rows("1:1").Select
    Selection.RowHeight = 31.5
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("C:C").ColumnWidth = 15.86
    Columns("C:C").ColumnWidth = 17.29
    Columns("D:D").ColumnWidth = 14.57
    Columns("E:E").ColumnWidth = 14.57
    Columns("F:F").ColumnWidth = 28.29
    Columns("F:F").ColumnWidth = 24.71
    Columns("F:F").ColumnWidth = 23.57
    Range("A1:F1").Select
    Selection.Font.Bold = True
    Range("E13").Select
    Sheets("NOT OPEN ASPER SHIFT").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Sheets("NOT_OPEN_BRANCH").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Sheets("PUNCHING STATUS REPORT").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Sheets("Punching Report").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("B1").Select

    Sheets("NOT OPEN ASPER SHIFT").Select
    Columns("I:J").Select
    Selection.NumberFormat = "[$-x-systime]h.mm.ss AM/PM"
    Range("L3").Select
    ActiveWindow.SmallScroll Down:=-9
    Sheets("NOT_OPEN_BRANCH").Select
    ActiveWindow.SmallScroll Down:=-39
    Columns("I:J").Select
    Selection.NumberFormat = "[$-x-systime]h.mm.ss AM/PM"
    Range("K2").Select
    Sheets("PUNCHING STATUS REPORT").Select
    ActiveWindow.SmallScroll Down:=-45
    Columns("I:J").Select
    Selection.NumberFormat = "[$-x-systime]h.mm.ss AM/PM"
    Range("I2").Select
    Sheets("Punching Report").Select
    ActiveWindow.SmallScroll Down:=-18
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    Columns("L:L").Select
    Selection.NumberFormat = "[$-x-systime]h.mm.ss AM/PM"
    Range("L2").Select
    ActiveWindow.SmallScroll Down:=-9

    ActiveWorkbook.Save
End Sub
