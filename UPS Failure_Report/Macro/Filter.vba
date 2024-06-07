Sub Macro1()
'
' Macro1 Macro
'

'
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Cells.Select
    Selection.ColumnWidth = 20.57
    Cells.EntireColumn.AutoFit
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=" ", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1)), TrailingMinusNumbers:=True
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollColumn = 2
    Columns("C:J").Select
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.ScrollColumn = 1
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Cells.Select
    Selection.ColumnWidth = 26.71
    Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Columns("C:F").Select
    Selection.ClearContents
    Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=".", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Columns("C:F").Select
    Selection.ClearContents
    Columns("B:B").EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="@", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Selection.ColumnWidth = 16.43
    Selection.ColumnWidth = 20.86
    Range("C1").Select
    ActiveWindow.SmallScroll Down:=-48
    Columns("C:G").Select
    Selection.ClearContents
    Range("C1").Select
    ActiveWindow.SmallScroll Down:=-9
    Columns("B:B").Select
    Selection.Replace What:="br0", Replacement:="br", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("B2").Select
    ActiveWorkbook.Save
End Sub
