Attribute VB_Name = "ADS_Prep"
Sub ADS_Prep()
    Dim folderPath As String
    Dim ADSPath As String
    Dim FinishedPath As String
    Dim filename As String
    Dim wb As Workbook
    Dim ADSwb As Workbook
    Dim ArtList
    
    folderPath = "RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$"
    ADSPath = "RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$"
    FinishedPath = "\RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$"
    filename = Dir(folderPath & "*.xlsx")
    Do While filename <> ""
      Application.ScreenUpdating = False
        Set wb = Workbooks.Open(folderPath & filename)

        'calculate the last row count
        lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
        SupName = ActiveSheet.Range("D2").Value
        ' prep the family names for inspection
        Range("R1").Select
        ActiveCell.FormulaR1C1 = "FamilyName"
        Range("R2").Select
        ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-4],"" "",RC[-3],"" "",RC[-2])"
        On Error Resume Next
        Selection.AutoFill Destination:=Range("R2:R" & lastrow)
        Range("R2:R" & lastrow).Select
        Columns("R:R").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
        ' Range Variables
        Dim rngRangeOfFamilies As Range
        Dim StartingRow As Long
        Dim EndingRow As Long
        Dim strLastFamilyName As String
        Dim LineCount As Long

        StartingRow = 2
        Set rngRangeOfFamilies = Range(ActiveSheet.Cells(StartingRow, 18), ActiveSheet.Cells((lastrow + 1), 18))
        
        For Each Cell In rngRangeOfFamilies
            Cell.Select
            strLastFamilyName = Cell.Offset(-1, 0).Value
            ArtNo = Cell.Offset(0, -15).Text
            If (Cell.Value <> strLastFamilyName) And (strLastFamilyName <> "FamilyName") Then
                EndingRow = Cell.Row - 1
                FamilyName = ActiveSheet.Cells(EndingRow, 18).Value
                ' MAKE NEW ADS SHEET HERE
                Set ADSwb = Workbooks.Open(ADSPath)
                ADSwb.Sheets(1).Range("B1").Select
                ActiveCell.FormulaR1C1 = SupName
                ADSwb.Sheets(1).Range("B7").Select
                ActiveCell.FormulaR1C1 = ArtList
                ADSwb.Sheets(1).Range("B8").Select
                ActiveCell.FormulaR1C1 = FamilyName
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs filename:=FinishedPath & SupName & "_" & FamilyName & "_ADS.xlsx", FileFormat:=51
                ActiveWorkbook.Close
                
                'Clean up
                StartingRow = Cell.Row
                ArtList = ArtNo & ", "
            Else
                ArtList = ArtList & ArtNo & ", "
            End If
        Next Cell
        
        Application.DisplayAlerts = False
        wb.Close
        
    filename = Dir
    Loop
  Application.ScreenUpdating = True
End Sub
