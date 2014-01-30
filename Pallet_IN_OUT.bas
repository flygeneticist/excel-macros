Attribute VB_Name = "Pallet_IN_OUT"
Sub pallet_in_out()
' pallet_in_out Macro
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>RC[-2],""Mismatch"","""")"
    lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    Range("R2").AutoFill Destination:=Range("R2:R" & lastrow), Type:=xlFillDefault
    Rows("1:1").Select
    Selection.AutoFilter
End Sub

