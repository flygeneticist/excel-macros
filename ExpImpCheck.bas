Attribute VB_Name = "ExpImpCheck"
Sub ExpImpCheck()
Attribute ExpImpCheck.VB_ProcData.VB_Invoke_Func = " \n14"
    lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[2],RC[16],RC[3])"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "ID"
    Range("B3").Select
    Selection.AutoFill Destination:=Range("B3:B" & lastrow)

    ' Creates a PivotTable report from the table on Sheet1 by using the PivotTableWizard
    ' method with the PivotFields method to specify the fields in the PivotTable.
    Dim objTable As PivotTable, objField As PivotField
    ' Select the sheet and first cell of the table that contains the data.
    ActiveWorkbook.Worksheets(1).Select
    Range("A2").Select
    ' Create the PivotTable object based on the Employee data on Sheet1.
    Set objTable = Worksheets(1).PivotTableWizard
    ' Specify row and column fields.
    Set objField = objTable.PivotFields("ID")
    objField.Orientation = xlRowField
    ' Specify a data field with its summary function and format.
    Set objField = objTable.PivotFields("EXPORT QTY")
    objField.Orientation = xlDataField
    objField.Function = xlSum

    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Imp Qty"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],Sheet1!C[-1]:C[8],10,0)"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "Imp>Exp?"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]>=RC[-2]"
    Range("C3:D3").Select
    Range("D3").Activate
    Selection.AutoFill Destination:=Range("C3:D" & lastrow)
    Range("C3:D" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
End Sub
