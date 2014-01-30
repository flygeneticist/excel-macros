Attribute VB_Name = "Clean2RMacro"
Sub Clean2R()
Attribute Clean2R.VB_ProcData.VB_Invoke_Func = " \n14"
'
' clean2R Macro
    lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row

    Columns("E:E").Select
    Selection.Replace What:="--", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "PORT"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "IMP DATE"
    
    Range("D2").Select
    Selection.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(SEARCH(712,RC[-1])),712,IF(ISNUMBER(SEARCH(1101,RC[-1])),1101,IF(ISNUMBER(SEARCH(2402,RC[-1])),2402,IF(ISNUMBER(SEARCH(2704,RC[-1])),2704,IF(ISNUMBER(SEARCH(3001,RC[-1])),3001,IF(ISNUMBER(SEARCH(3004,RC[-1])),3004,IF(ISNUMBER(SEARCH(4601,RC[-1])),4601,"""")))))))"
    
    Range("E2").Select
    Selection.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(SEARCH(712,RC[-2])),RIGHT(RC[-2],LEN(RC[-2])-LEN(712)),IF(ISNUMBER(SEARCH(1101,RC[-2])),RIGHT(RC[-2],LEN(RC[-2])-LEN(1101)),IF(ISNUMBER(SEARCH(2402,RC[-2])),RIGHT(RC[-2],LEN(RC[-2])-LEN(2402)),IF(ISNUMBER(SEARCH(2704,RC[-2])),RIGHT(RC[-2],LEN(RC[-2])-LEN(2704)),IF(ISNUMBER(SEARCH(3001,RC[-2])),RIGHT(RC[-2],LEN(RC[-2])-LEN(3001)),IF(ISNUMBER(SEARCH(3004,RC[-2])),RIGHT(RC[-2],LEN(RC[-2])-LEN(3004)),IF(ISNUMBER(SEARCH(4601,RC[-2])),RIGHT(RC[-2],LEN(RC[-2])-LEN(4601)),"""")))))))"
    
    Range("D2:E2").AutoFill Destination:=Range("D2:E" & lastrow), Type:=xlFillDefault
    Columns("D:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
End Sub
