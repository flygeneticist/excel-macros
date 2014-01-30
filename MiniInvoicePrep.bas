Attribute VB_Name = "MiniInvoicePrep"
Sub Mini_InvoicePrep()
    Dim lastrow
    Dim main_wb
    main_wb = ActiveWorkbook.Name
    Application.ScreenUpdating = True
    
' insert autofilters into report's first row
    Rows("1:1").Select
    Selection.AutoFilter
    lastrow = Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
    
' Inserts new columns A and pulls down blank fixing formula
    Columns("A:A").Select
    Selection.Delete
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("A:A").Select
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 2), Array(2, 9), Array(3, 2)), TrailingMinusNumbers:=True
   Range("A1").Select
   ActiveCell.FormulaR1C1 = "Art"
   Range("A1").Select
   ActiveCell.FormulaR1C1 = "Sup"
   
' place formulas into new columns to pull qty and price into correct places
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Qty"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(LEFT(RC[-1],1)<>""A""),R[1]C[-1],"""")"
    Range("E2").AutoFill Destination:=Range("E2:E" & lastrow)
    
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Price"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(LEFT(RC[-1],1)<>""A""),R[1]C[-1],"""")"
    Range("G2").AutoFill Destination:=Range("G2:G" & lastrow)
    
    Columns("E:G").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("D").Delete
    Columns("E").Delete
    Columns("F:G").Delete
    
'filters and removes the rows that contain the subtotaled article counts
    With ws
    'filter for blanks and remove rows
        ActiveSheet.Range("A1:J" & lastrow).AutoFilter Field:=4, Criteria1:="0"
        
        LR = Cells(Rows.Count, 1).End(xlUp).Row
        If LR > 1 Then
            Range("A2:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        ActiveSheet.ShowAllData
        
    'filter for Customs stat number rows and remove
        ActiveSheet.Range("A1:J" & lastrow).AutoFilter Field:=1, Criteria1:="=Cost*"
        LR = Cells(Rows.Count, 1).End(xlUp).Row
        If LR > 1 Then
            Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        ActiveSheet.ShowAllData
    End With
    lastrow = Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
    
' finalize formulas and delete the source column/s
    Columns("A:F").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 
' fix  "," issue with the price and qty
    Columns("D:E").Select
    Selection.Replace What:=",", Replacement:=".", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("E").Select
    Selection.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    
' text formatting/sizing, auto filters, and auto size columns
    Rows(1).Select
    Selection.Font.Bold = True
    Selection.AutoFilter
    Selection.AutoFilter
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 8
    End With
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    
    Application.Goto Reference:=Range("A1"), Scroll:=True
    Application.ScreenUpdating = True

End Sub

