Attribute VB_Name = "CarPrepMacro"
Sub CARPrep()
    lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
'inserts the uniqueID title and formula into a newly created column in the spreadsheet
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "UniqueID"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[4],RC[3],RC[1],RC[2])"
    Range("A2").AutoFill Destination:=Range("A2:A" & lastrow), Type:=xlFillDefault
    
'selects the UniqueID column and copy/pastes special to lock in values
    CopyPasteColumns ("A:A")

'inserts the tCalc columns into the spreadsheet
    Columns("AC:AC").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AC:AC").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AC:AC").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
'inserts the tValue title and formula into a newly created column in the spreadsheet
    Range("AC1").Select
    ActiveCell.FormulaR1C1 = "tValue"
    Range("AC2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[4]=""x"",RC[10]+RC[16],IF(AND(RC[7]="""",RC[6]=""""),RC[11],IF(AND(RC[7]<>"""",RC[6]<>""""),RC[5]+RC[11]+RC[17]+RC[23],"""")))"
'inserts the tDuty title and formula into a newly created column in the spreadsheet
    Range("AD1").Select
    ActiveCell.FormulaR1C1 = "tDuty"
    Range("AD2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[3]=""x"",RC[5],IF(AND(RC[6]=""""),RC[11],IF(AND(RC[5]<>"""",RC[6]<>""""),RC[5]+RC[11]+RC[17]+RC[23],"""")))"
'inserts the tRate title and formula into a newly created column in the spreadsheet
    Range("AE1").Select
    ActiveCell.FormulaR1C1 = "tRate"
    Range("AE2").Select
    ActiveCell.FormulaR1C1 = "=ROUND(RC[-1]/RC[-2],4)"
    
'fills down all of the newly created theoretical cells and pastes in values
    Range("AC2:AE2").AutoFill Destination:=Range("AC2:AE" & lastrow), Type:=xlFillDefault
    CopyPasteColumns ("AC:AE")
    
'insert autofilters into report's first row
    Rows("1:1").Select
    Selection.AutoFilter

'filters and removes the rows that contain the blanks for DISCUS numbers
    With ws
        ActiveSheet.Range("A1:CA" & lastrow).AutoFilter Field:=4, Criteria1:="="
        LR = Cells(Rows.Count, 1).End(xlUp).Row
        If LR > 1 Then
            Range("A2:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
    End With

'Recalculate the last row count
    ActiveSheet.ShowAllData
    lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row

'sorts the active sheet by the 4R/CAR initial sort criteria of:
'Entry#,Exp#,Art#,4R_ExpQty,4R_ExpDuty
    Cells.Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "E2:E" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "D2:D" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "B2:B" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "I2:I" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "J2:J" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1:BC" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
       
    Application.ScreenUpdating = True

End Sub

Sub CopyPasteColumns(Column)
    Columns(Column).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

