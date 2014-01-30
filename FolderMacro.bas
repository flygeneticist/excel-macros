Attribute VB_Name = "FolderMacro"
Sub AllFiles()
    Dim folderPath As String
    Dim filename As String
    Dim wb As Workbook
    
    folderPath = "\\DSUS061-NT0001\KEKEL1$\Desktop\NAFTA Sup Rpts\Raw Sup Rpts\"
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath + "\"
    
    filename = Dir(folderPath & "*.xls")
    Do While filename <> ""
      Application.ScreenUpdating = False
        Set wb = Workbooks.Open(folderPath & filename)
         
        'Call a subroutine here to operate on the just-opened workbook
        lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
        Columns("B:B").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("B5").Select
        ActiveCell.FormulaR1C1 = "Sup No"
        Range("B6").Select
        ActiveCell.FormulaR1C1 = "=REPLACE(R2C1,1,14,"""")"
        Range("B6").Select
        Selection.AutoFill Destination:=Range("B6:B" & lastrow)
        Range("B6:B" & lastrow).Select
        Columns("B:B").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        Rows("1:4").Select
        Selection.Delete Shift:=xlUp
        Cells.Select
        ActiveSheet.Range("$A1:BV" & lastrow).RemoveDuplicates Columns:=Array(1, 2, 6), _
            Header:=xlYes
        Columns("G:CC").Select
        Selection.Delete Shift:=xlToLeft
        Columns("C:C").ColumnWidth = 8.25
        Columns("C:C").EntireColumn.AutoFit
        SupName = [B2]
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs filename:=folderPath & SupName & ".csv", FileFormat:=xlCSV
        ActiveWorkbook.Close
    filename = Dir
    Loop
  Application.ScreenUpdating = True
End Sub

