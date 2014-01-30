VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewClaimVettingForm 
   Caption         =   "New Drawback Claim Vetting Program"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6045
   OleObjectBlob   =   "NewClaimVettingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewClaimVettingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Defines the public/global variables that are accessable to all subs herein
Dim lastrow
Dim fPath(1 To 4)
Dim fName(1 To 4)
Dim fRootPath(1 To 4)
Dim fRange(1 To 8)
Dim wb(1 To 4) As Workbook

Private Sub CommandButton1_Click()
    Dim fpath1
    fpath1 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=1, Title:="Open File", MultiSelect:=False)
    If fpath1 <> False Then
        TextBox1.ForeColor = &H80000017
        TextBox1.Value = fpath1
        fPath(1) = fpath1
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub

Private Sub CommandButton2_Click()
    Dim fpath2
    fpath2 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=1, Title:="Open File", MultiSelect:=False)
    If fpath2 <> False Then
        TextBox2.ForeColor = &H80000017
        TextBox2.Value = fpath2
        fPath(2) = fpath2
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub
Private Sub CommandButton3_Click()
    Dim fpath3
    fpath3 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=1, Title:="Open File", MultiSelect:=False)
    If fpath3 <> False Then
        TextBox3.ForeColor = &H80000017
        TextBox3.Value = fpath3
        fPath(3) = fpath3
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub
Private Sub CommandButton4_Click()
    Dim fpath4
    fpath4 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=1, Title:="Open File", MultiSelect:=False)
    If fpath4 <> False Then
        TextBox4.ForeColor = &H80000017
        TextBox4.Value = fpath4
        fPath(4) = fpath4
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub

Private Sub CancelButton_Click()
    End
End Sub

Private Sub StartButton_Click()
' executes a final file check code before running the NEW DRB PROCESS
' checks the strings given by the user are valid, opens target wb, and
' saves wb name to fName array along with workbook open function to the
' wb array for easy calling to switch windows.
    Dim x As Integer
    Dim c As Integer
    x = 1
    c = 1
    While x < 5
        If fPath(x) <> False Then
            Application.Workbooks.Open (fPath(x))
            Set wb(x) = Workbooks.Open(fPath(x))
            fName(x) = ActiveWorkbook.Name
            fRootPath(x) = ActiveWorkbook.Path
            fRootPath(x) = fRootPath(x) & "\"
            ' set previously nameed sheets to raw sht_num to avoid a re-naming error
            sht_num = 1
            For Each sht In ActiveWorkbook.Worksheets
                sht.Name = sht_num
                sht_num = sht_num + 1
            Next sht
            ' now rename all sheets correctly
            sht_num = 1
            For Each sht In ActiveWorkbook.Worksheets
                sht.Name = "Sheet" & sht_num
                sht_num = sht_num + 1
            Next sht
            If x = 1 Then
                fRange(c) = "'" & fRootPath(x) & "[" & fName(x) & "]Sheet2'!"
                c = c + 1
                fRange(c) = "'" & fRootPath(x) & "[" & fName(x) & "]Sheet4'!"
                c = c + 1
                x = x + 1
            Else
                fRange(c) = "'" & fRootPath(x) & "[" & fName(x) & "]Sheet1'!"
                c = c + 1
                fRange(c) = "'" & fRootPath(x) & "[" & fName(x) & "]Sheet2'!"
                c = c + 1
                x = x + 1
            End If
        ' if a file string is not set, disply error message and end program
        Else
            MsgBox "You did not specify a valid location or file for all of the required reports. The application will now terminate."
            End
        End If
    Wend

' executes the main code for the NEW DRB CLAIM VETTING process
    VettingNewClaim ' apply checks to the claim from external reports
    LayoutTweeks ' standardized the look/feel of the claim

    ' Close the external data source files without saving changes
    wb(2).Close savechanges:=False ' CAR wb
    wb(3).Close savechanges:=False '  Pallet wb
    wb(4).Close savechanges:=False ' BoL wb
    Windows("Vetting DRB New Claim Template.xlsx").Close savechanges:=False ' temp formula placeholder file
    
    ' restores the screen updating to the application object
    Application.ScreenUpdating = True
    
' Close the main sub, ending the program
End Sub

Sub VettingNewClaim()
    Application.ScreenUpdating = True
    ' create the correct var links to each of the workbooks from the initial setup in 'Start_button'
    Dim OHL_Range_Sht2 As String
    Dim OHL_Range_Sht4 As String
    Dim CAR_Range_Sht1 As String
    Dim Pallet_Range_Sht1 As String
    Dim Pallet_Range_Sht2 As String
    Dim BoL_Range_Sht1 As String
    
    OHL_Range_Sht2 = fRange(1)
    OHL_Range_Sht4 = fRange(2)
    CAR_Range_Sht1 = fRange(3)
    Pallet_Range_Sht1 = fRange(5)
    Pallet_Range_Sht2 = fRange(6)
    BoL_Range_Sht1 = fRange(7)
    
    ' setup date, month, year variables to ajust formulas that are dependent
    Dim todaydate As Date ' open variable in memory
    todaydate = Date ' set variable equal to system date
    dateFormula = "DATE(" & VBA.Year(todaydate) & "," & VBA.Month(todaydate) & "," & VBA.day(todaydate) & ")" ' format variable to be formula readable
    
    ' prep the CAR for working into the DRB claim
    SwitchWindows (2)
    CARPrep
    
    ' --------------------------------------------------------------------------
    'PREP THE 4R
    ' --------------------------------------------------------------------------
    SwitchWindows (1)
    ActiveWorkbook.Sheets(4).Activate
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("G:G").ColumnWidth = 15
    Workbooks.Open filename:="\\DSUS061-NT0001\KEKEL1$\Desktop\DRB Stuff\Vetting DRB New Claim Template.xlsx", UpdateLinks:=0
    Range("Q11:AH12").Select
    Selection.Copy
    
    SwitchWindows (1)
    Range("Q11").Select
    ActiveSheet.Paste
    Range("Q12").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-6]<>"""",VLOOKUP(RC[-16]," & Pallet_Range_Sht1 & "C1:C3,3,0),"""")"
    Range("A11").Select
    Windows("Vetting DRB New Claim Template.xlsx").Activate
    Range("A11:B12").Select
    Selection.Copy
    
    SwitchWindows (1)
    ActiveSheet.Paste
    Range("V12").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>"""",VLOOKUP(RC[-21]," & CAR_Range_Sht1 & "C1:C23,23,0),"""")"
    Range("X12").Select
    ActiveCell.FormulaR1C1 = "=IF(RC23<>"""",VLOOKUP(RC1," & CAR_Range_Sht1 & "C1:C31,29,0),"""")"
    Range("X12").Select
    Selection.AutoFill Destination:=Range("X12:Z12"), Type:=xlFillDefault
    Range("Y12").Select
    ActiveCell.FormulaR1C1 = "=IF(RC23<>"""",VLOOKUP(RC[-24]," & CAR_Range_Sht1 & "C1:C31,30,0),"""")"
    Range("Z12").Select
    ActiveCell.FormulaR1C1 = "=IF(RC23<>"""",VLOOKUP(RC[-25]," & CAR_Range_Sht1 & "C1:C31,31,0),"""")"
    Range("AA12").Select
    ActiveCell.FormulaR1C1 = "=IF(RC26<>"""",LEFT(VLOOKUP(RC[-26]," & Pallet_Range_Sht2 & "C1:C14,13,0),3),"""")"
    Range("AB12").Select
    ActiveCell.FormulaR1C1 = "=IF(LEFT(RC[-1],3)=""439"",RIGHT(VLOOKUP(RC[-27]," & Pallet_Range_Sht2 & "C1:C14,14,0),8),"""")"
    Range("AC12").Select
    ActiveCell.FormulaR1C1 = "=IF(LEFT(RC[-2],3)=""439"",VLOOKUP(RC28," & BoL_Range_Sht1 & "C9:C14,6,0),"""")"
    Range("AD12").Select
    ActiveCell.FormulaR1C1 = "=IF(LEFT(RC[-3],3)=""439"",VLOOKUP(RC28," & BoL_Range_Sht1 & "C9:C19,10,0),"""")"
    
    ' Pastes over the pre-set formulas
    Range("A12:B12").Select
    lastrow = ActiveSheet.Range("C" & Rows.Count).End(xlUp).Row
    Selection.AutoFill Destination:=Range("A12:B" & lastrow)
    Range("A12:B" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ' modify date sensitive formulas
    Range("U12").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-4]<>"""",IF((" & dateFormula & " -Date(Year(RC[-15]),Month(RC[-15]),Day(RC[-15])))<=1088,""Inside"",IF(AND((" & dateFormula & " -Date(Year(RC[-15]),Month(RC[-15]),Day(RC[-15])))<=1095, (" & dateFormula & " -Date(Year(RC[-15]),Month(RC[-15]),Day(RC[-15])))>1088),""Check!"",""Drop!"")),"""")"
    Range("W12").Select
    ActiveCell.Formula = "=IFERROR(IF(RC[-6]<>"""",IF(RC[-1]=0,""Check!"",IF((" & dateFormula & " -Date(LEFT(RC[-1],4),MID(RC[-1],5,2),RIGHT(RC[-1],2)))>=180,""Okay"",""Check!"")),""""),""Check!"")"
    
    ' fill down all the formulas to the last row
    Range("Q12:AD12").Select
    Selection.AutoFill Destination:=Range("Q12:AD" & lastrow)
       
    ' fix the formula values into place
    Range("Q12:AD" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ' correct dates into readable formats
    Columns("V:V").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.TextToColumns Destination:=Range("V1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 5), TrailingMinusNumbers:=True
    Columns("V:V").EntireColumn.AutoFit
      
        
    ' --------------------------------------------------------------------------
    'PREP THE 2R
    ' --------------------------------------------------------------------------
    ActiveWorkbook.Sheets(2).Activate
    lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    ' setup the unique IDs for linking reports
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Windows("Vetting DRB New Claim Template.xlsx").Activate
    Range("A4:A5").Select
    Selection.Copy
    SwitchWindows (1)
    Range("A11").Select
    ActiveSheet.Paste
    Range("A12").Select
    lastrow = ActiveSheet.Range("C" & Rows.Count).End(xlUp).Row
    Selection.AutoFill Destination:=Range("A12:A" & lastrow)
    
    ' remove article numbers' leading zeros
    Columns("G:G").Select
    Selection.TextToColumns Destination:=Range("G1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("A:A").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ' paste in the formulas for 2R
    Windows("Vetting DRB New Claim Template.xlsx").Activate
    Range("U4:AH5").Select
    Selection.Copy
    SwitchWindows (1)
    Range("U11").Select
    ActiveSheet.Paste
    Range("X12").Select
    Range("X12").FormulaR1C1 = "=IF(RC[-1]=""FALSE"",VLOOKUP(RC[-23],Sheet4!C2:C26,24,0),"""")"
    Range("AC12").Select
    Range("AC12").FormulaR1C1 = "=IF(RC[-1]<>"""",VLOOKUP(RC[-28],Sheet4!C2:C26,23,0)/RC[-20],"""")"
    Range("AF12").Select
    Range("AF12").FormulaR1C1 = "=IF(RC[-4]<>"""",VLOOKUP(RC[-31],Sheet4!C2:C26,25,0),"""")"
    
    ' Add new checks to re-calc the 99% based off CAR duty as a check on OHL rolling up duty for similar arts
    Range("AI11").Select
    Range("AI11").FormulaR1C1 = "CAR/IKEA 99%"
    Range("AI12").Select
    Range("AI12").FormulaR1C1 = "=IF(RC[-14]<>"""",TRUNC((VLOOKUP(RC[-34],Sheet4!C[-33]:C[-9],24,0)/RC[-26]*RC[-25]*0.99),2),"""")"
    
    Range("AJ11").Select
    Range("AJ11").FormulaR1C1 = "Match?"
    Range("AJ12").Select
    Range("AJ12").FormulaR1C1 = "=IF(RC[-1]<>"""",IF(ABS(TRUNC(RC[-1]-RC[-10],2))<=0.03,""TRUE"",""FALSE""),"""")"
    
    Range("AK11").Select
    Range("AK11").FormulaR1C1 = "Diff"
    Range("AK12").Select
    Range("AK12").FormulaR1C1 = "=IF(AND(RC[-1]<>"""",RC[-1]<>""TRUE""),ABS(TRUNC(RC[-2]-RC[-22],2)),"""")"
     
    ' pull down and copy paste all values once calcs are finished
    Range("U12:AK12").Select
    Selection.AutoFill Destination:=Range("U12:AK" & lastrow)
    Range("U12:AK" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub SwitchWindows(ArrayNumber) ' Sub-routine will switch windows, get lastrow var, and turn off screen updating if on
    wb(ArrayNumber).Activate
    Application.ScreenUpdating = False
    lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
End Sub

Sub CopyPasteColumns(Column)
    Columns(Column).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Public Sub CARPrep()
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

Sub LayoutTweeks()
' modify the visual layout of the claim to make starting the vetting process easier
    ' working over sheet 2 from the last sub end point
    Cells.Select
    Application.CutCopyMode = False
    With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Columns("H:H").Select
    Selection.ColumnWidth = 8.29
    Columns("I:I").Select
    Selection.ColumnWidth = 6.86
    Columns("J:J").ColumnWidth = 6.57
    Columns("L:L").ColumnWidth = 7.14
    Columns("N:N").ColumnWidth = 6.71
    Columns("F:F").ColumnWidth = 28.43
    Columns("E:E").ColumnWidth = 10.86
    Columns("D:D").ColumnWidth = 7
    Columns("C:C").ColumnWidth = 6

    Range("AI11:AK11").Select
    Range("AK11").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    
    Columns("AI:AK").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("AI:AI,AK:AK").Select
    Range("AK1").Activate
    Selection.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    
    Rows("11:11").Select
    Range("S11").Activate
    Selection.AutoFilter
    Range("A11").Select
    Selection.Font.Bold = True
    
    Columns("AC:AC").Select
    Selection.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    Columns("AC:AC").EntireColumn.AutoFit
    Columns("AF:AF").Select
    Selection.NumberFormat = "0.00%"
    
    
    ' modify sheet 4 in a similar fashion
    Sheets("Sheet4").Select
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Columns("D:D").ColumnWidth = 9
    Columns("E:E").ColumnWidth = 12.5
    Columns("G:G").ColumnWidth = 7
    Columns("H:H").EntireColumn.AutoFit
    Columns("I:I").EntireColumn.AutoFit
    Columns("K:K").ColumnWidth = 9
    Columns("L:L").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit

    Columns("Q:AD").EntireColumn.AutoFit
    Columns("AA:AA").Select
    Selection.ColumnWidth = 9
    Columns("AC:AC").ColumnWidth = 13
    Rows("11:11").Select
    Range("M11").Activate
    Selection.AutoFilter

    Columns("Y:Y").Select
    Selection.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    Columns("Z:Z").Select
    Selection.NumberFormat = "0.00%"
    Columns("Y:Z").EntireColumn.AutoFit
    Range("A11:B11").Select
    Selection.Font.Bold = True
    
    ' for sheets 1 and 3, largely unused in the vetting process, apply simple font standardization changes
    Sheets("Sheet1").Select
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Sheets("Sheet3").Select
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Sheets("Sheet2").Select
End Sub

