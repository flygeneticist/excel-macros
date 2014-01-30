VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScheduleMakerForm 
   Caption         =   "UserForm3"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5970
   OleObjectBlob   =   "ScheduleMakerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ScheduleMakerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Defines the public/global variables that are accessable to all subs herein
Dim lastrow
Dim fPath(1 To 4)
Dim fName(1 To 4)
Dim fRootPath(1 To 4)
Dim fRange(1 To 4)
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
    fpath4 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=2, Title:="Open File", MultiSelect:=False)
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
    x = 1
    While x < 5
        If fPath(x) <> False Then
            Application.Workbooks.Open (fPath(x))
            Set wb(x) = Workbooks.Open(fPath(x))
            fName(x) = ActiveWorkbook.Name
            fRootPath(x) = ActiveWorkbook.Path
            fRootPath(x) = fRootPath(x) & "\"
            fRange(x) = "'" & fRootPath(x) & "[" & fName(x) & "]Sheet1'!"
            x = x + 1
        ' if a file string is not set, disply error message and end program
        Else
            MsgBox "You did not specify a valid location or file for all of the required reports. The application will now terminate."
            End
        End If
    Wend

' executes the main code for the NEW DRB PROCESS
    ScheduleMaker
    Application.ScreenUpdating = True
    
' Close the program/sub
End Sub

Sub ScheduleMaker()
'run the prep processes for the new drb
    R1prep
    R3Prep
    WARPasting
    ScheduleFormatting
    
' CLOSE THE B3, CA SETS, AND XTHETA FILES
    wb(1).Close savechanges:=False ' 1R wb
    wb(2).Close savechanges:=False ' 3R wb
    wb(3).Close savechanges:=False ' WAR wb

    MsgBox "Your Schedule Report has been completed!" ' Show completion message box

End Sub

Sub R1prep()
' Rprep Macro
    SwitchWindows (2)
    Columns("A:A").Select
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("E:E").Select
    Selection.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
End Sub

Sub R3Prep()
' RPrep Macro
    SwitchWindows (3)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[2],RC[7],RC[8],RC[9])"
    Range("A2").AutoFill Destination:=Range("A2:A" & lastrow)
    Columns("H:H").Select
    Selection.TextToColumns Destination:=Range("H1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
End Sub

Sub WARPasting()
' WARPastes Macro

' copy and paste WAR row #s
    SwitchWindows (1)
    Range("B3:B" & lastrow).Select
    Selection.Copy
    
    SwitchWindows (4)
    Sheets("Designation Summary (2R)").Select
    Range("A22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Sheets("Designation Sheet (4R)").Select
    
    SwitchWindows (1)
    Selection.Copy
    SwitchWindows (4)
    Range("A22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
' --------------------------------------------------------------------------------
'             2R BUILDING
' --------------------------------------------------------------------------------
' Copy over the new 2R
    SwitchWindows (1)
    Range("C3:P" & lastrow).Select
    Selection.Copy
    
    SwitchWindows (4)
    Sheets("Designation Summary (2R)").Select
    Range("P22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
' copy over the new 4R
    SwitchWindows (1)
    Range("Q3").Select
    Range("Q3:Z" & lastrow).Select
    Selection.Copy
    
    SwitchWindows (4)
    Sheets("Designation Sheet (4R)").Select
    Range("L22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
' copy over the old 2R
    SwitchWindows (1)
    Range("BB3:BO" & lastrow).Select
    Selection.Copy
    SwitchWindows (4)
    Sheets("Designation Summary (2R)").Select
    Range("B22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    ' format paintbrush
    Range("B22:O" & lastrow).Select
    Selection.Copy
    Range("P22").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'apply conditional formatting
    Range("P22:W" & lastrow).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=B22<>P22"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("Y22:AA" & lastrow).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=K22<>Y22"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("AB22:AB" & lastrow).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=TRUNC(ABS(N22-AB22),4)<>0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
' --------------------------------------------------------------------------------
'             4R BUILDING
' --------------------------------------------------------------------------------
' copy over the old 4R
    SwitchWindows (1)
    Range("BP3:BY" & lastrow).Select
    Selection.Copy
    SwitchWindows (4)
    Sheets("Designation Sheet (4R)").Select
    Range("B22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    ' format paintbrush
    Range("B22:K" & lastrow).Select
    Selection.Copy
    Range("L22").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    ' apply conditional formatting
    Range("L22:O" & lastrow).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=B22<>L22"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("Q22:T" & lastrow).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=G22<>Q22"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ' copy over the final comments and code into the 2R and 4R
    SwitchWindows (1)
    Range("CS3:CT" & lastrow).Select
    Selection.Copy
    SwitchWindows (4)
    Sheets("Designation Summary (2R)").Select
    Range("AF22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Designation Sheet (4R)").Select
    Range("X22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub

Sub ScheduleFormatting()
    SwitchWindows (4)
' --------------------------------------------------------------------------------
'             2R FORMATTING
' --------------------------------------------------------------------------------
' differences and final columns formatting 2R
    Sheets("Designation Summary (2R)").Select
    Range("AD22").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-15]"
    Range("AE22").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]"
    Range("AD22:AE22").Select
    Selection.Copy
    Range("AD22:AE" & lastrow).Select
    Application.CutCopyMode = False
    Selection.FillDown
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

' formatting box around 2R comments column
    Range("AF22:AF" & lastrow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

' formatting box around the 2R code column
    Range("AG22:AG" & lastrow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

' formatting around the 2R code diff column
    Range("AH22:AH" & lastrow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

' --------------------------------------------------------------------------------
'             4R FORMATTING
' --------------------------------------------------------------------------------
' differences and final columns formatting 4R
    Sheets("Designation Sheet (4R)").Select
    Range("V22").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-11]"
    Range("W22").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]"
    Range("V22:W" & lastrow).Select
    Application.CutCopyMode = False
    Selection.FillDown
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

' format a box around the 4R comments
    Range("X22:X" & lastrow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
' format box around the 4R code
    Range("Y22:Y" & lastrow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("Z22").Select
    
' format box around the 4R code diff column
    Range("Z22:Z10868").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

' --------------------------------------------------------------------------------
'             3R BUILDING
' --------------------------------------------------------------------------------
    ' copy article and supplier numbers
    Sheets("Designation Sheet (4R)").Select
    Range("L22:M" & lastrow).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Chronological Summary (3R)").Select
    Range("G10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Designation Sheet (4R)").Select
    Range("B22:C" & lastrow).Select
    Selection.Copy
    Sheets("Chronological Summary (3R)").Select
    Range("Q10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
     
    'copy DISCUS number
    Sheets("Designation Sheet (4R)").Select
    Range("N22:N" & lastrow).Select
    Selection.Copy
    Sheets("Chronological Summary (3R)").Select
    Range("L10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Designation Sheet (4R)").Select
    Range("D22:D" & lastrow).Select
    Selection.Copy
    Sheets("Chronological Summary (3R)").Select
    Range("B10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ' copy export date
    Sheets("Designation Sheet (4R)").Select
    Range("O22:O" & lastrow).Select
    Selection.Copy
    Sheets("Chronological Summary (3R)").Select
    Range("K10").Select
    Selection.PasteSpecial
    Sheets("Designation Sheet (4R)").Select
    Range("E22:E" & lastrow).Select
    Selection.Copy
    Sheets("Chronological Summary (3R)").Select
    Range("A10").Select
    Selection.PasteSpecial

    ' copy export qty and uom
    Sheets("Designation Sheet (4R)").Select
    Range("P22:Q" & lastrow).Select
    Selection.Copy
    Sheets("Chronological Summary (3R)").Select
    Range("S10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Designation Sheet (4R)").Select
    Range("F22:G" & lastrow).Select
    Selection.Copy
    Sheets("Chronological Summary (3R)").Select
    Range("I10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    ' differences and final columns formatting 3R
    Sheets("Chronological Summary (3R)").Select
    lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    Range("U10").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-12]"
    Range("V10").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]"
    Range("U10:V10").Select
    Selection.Copy
    Range("U10:V" & lastrow).Select
    Application.CutCopyMode = False
    Selection.FillDown
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    ' pull in misc info from original 3R via a vlookup
    R3_Range = fRange(3)
    ' create uniqueID for lookups
    Range("W10").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-21],RC[-16],RC[-15],RC[-14])"
    Range("W10:W" & lastrow).Select
    Selection.FillDown
    ' Inserts the 3R vlookup formulas
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[20]," & R3_Range & "C1:C7,4,0)"
    Range("D10").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[19]," & R3_Range & "C1:C7,5,0)"
    Range("E10").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[18]," & R3_Range & "C1:C7,6,0)"
    Range("F10").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[17]," & R3_Range & "C1:C7,7,0)"
    Range("C10:F10").Select
    Selection.Copy
    Range("C10:F" & lastrow).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("M10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ' removes the uniqueIDs
    Columns("W:W").Select
    Selection.ClearContents
    
    ' format painbrush
    Range("A10:J" & lastrow).Select
    Selection.Copy
    Range("K10").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    ' conditional formatting
    Range("K10:O" & lastrow).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=A10<>K10"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Sheets("Adeps (1R)").Select
    
' --------------------------------------------------------------------------------
'             1R BUILDING
' --------------------------------------------------------------------------------
    ' copy over the 1R into the WAR
    SwitchWindows (2)
    Range("A2:E" & lastrow).Select
    Selection.Copy
    SwitchWindows (4)
    Sheets("Adeps (1R)").Select
    Range("A22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("F22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("L22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J22").Select
    
    lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row

    ' remove liq dates and old subtotals to be replaced
    Range("J22:J" & lastrow).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("O22:P" & lastrow).Select
    Selection.ClearContents
    
    ' insert differences between various 1Rs
    Range("K22").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-6]"
    Range("K22:K" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    Range("Q22").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-7]"
    Range("R22").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]"
    Range("S22:S" & lastrow).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    ' format the columns
    Range("S22:S" & lastrow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone


    Range("Q22:R" & lastrow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone


    Range("A22:P" & lastrow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
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

Sub FormulaPulldown(Cell As String, Formula As String)
    Range(Cell).Select
    ActiveCell.FormulaR1C1 = Formula
    Range(Cell).AutoFill Destination:=Range("Cell:Cell" & lastrow)
End Sub

Sub InsertFormula(Cell As String, Formula As String)
    Range(Cell).Select
    ActiveCell.FormulaR1C1 = Formula
End Sub
