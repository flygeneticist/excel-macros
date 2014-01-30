VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "IKEA DRB Prep Software"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Defines the public/global variables that are accessable to all subs herein
Dim lastrow
Dim fPath(1 To 8)
Dim fName(1 To 8)
Dim fRootPath(1 To 8)
Dim fRange(1 To 8)
Dim fRange2(1 To 2)
Dim Wb(1 To 8) As Workbook

Private Sub CommandButton1_Click()
    Dim fpath5
    fpath5 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=1, Title:="Open File", MultiSelect:=False)
    If fpath5 <> False Then
        TextBox1.ForeColor = &H80000017
        TextBox1.Value = fpath5
        fPath(5) = fpath5
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub

Private Sub CommandButton2_Click()
    Dim fpath6
    fpath6 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=1, Title:="Open File", MultiSelect:=False)
    If fpath6 <> False Then
        TextBox2.ForeColor = &H80000017
        TextBox2.Value = fpath6
        fPath(6) = fpath6
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub
Private Sub CommandButton3_Click()
    Dim fpath7
    fpath7 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=1, Title:="Open File", MultiSelect:=False)
    If fpath7 <> False Then
        TextBox3.ForeColor = &H80000017
        TextBox3.Value = fpath7
        fPath(7) = fpath7
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub
Private Sub CommandButton4_Click()
    Dim fpath3
    fpath3 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=2, Title:="Open File", MultiSelect:=False)
    If fpath3 <> False Then
        TextBox4.ForeColor = &H80000017
        TextBox4.Value = fpath3
        fPath(3) = fpath3
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub
Private Sub CommandButton5_Click()
    Dim fpath4
    fpath4 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=2, Title:="Open File", MultiSelect:=False)
    If fpath4 <> False Then
        TextBox5.ForeColor = &H80000017
        TextBox5.Value = fpath4
        fPath(4) = fpath4
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub
Private Sub CommandButton6_Click()
    Dim fpath8
    fpath8 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=1, Title:="Open File", MultiSelect:=False)
    If fpath8 <> False Then
        TextBox6.ForeColor = &H80000017
        TextBox6.Value = fpath8
        fPath(8) = fpath8
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub

Private Sub CommandButton7_Click()
    Dim fpath1
    fpath1 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=1, Title:="Open File", MultiSelect:=False)
    If fpath1 <> False Then
        TextBox7.ForeColor = &H80000017
        TextBox7.Value = fpath1
        fPath(1) = fpath1
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub

Private Sub CommandButton8_Click()
    Dim fpath2
    fpath2 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=1, Title:="Open File", MultiSelect:=False)
    If fpath2 <> False Then
        TextBox8.ForeColor = &H80000017
        TextBox8.Value = fpath2
        fPath(2) = fpath2
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub

Private Sub CommandButton9_Click()
    Dim fpath3
    fpath3 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=2, Title:="Open File", MultiSelect:=False)
    If fpath3 <> False Then
        TextBox9.ForeColor = &H80000017
        TextBox9.Value = fpath3
        fPath(3) = fpath3
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub

Private Sub CommandButton10_Click()
    Dim fpath4
    fpath4 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=2, Title:="Open File", MultiSelect:=False)
    If fpath4 <> False Then
        TextBox10.ForeColor = &H80000017
        TextBox10.Value = fpath4
        fPath(4) = fpath4
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub

Private Sub CommandButton11_Click()
    Dim fpath1
    fpath1 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=1, Title:="Open File", MultiSelect:=False)
    If fpath1 <> False Then
        TextBox11.ForeColor = &H80000017
        TextBox11.Value = fpath1
        fPath(1) = fpath1
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub

Private Sub CommandButton12_Click()
    Dim fpath2
    fpath2 = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=1, Title:="Open File", MultiSelect:=False)
    If fpath2 <> False Then
        TextBox12.ForeColor = &H80000017
        TextBox12.Value = fpath2
        fPath(2) = fpath2
    Else
        MsgBox ("Please select a valid file.")
    End If
End Sub

Private Sub CancelButtonOld_Click()
    End
End Sub

Private Sub CancelButtonNew_Click()
    End
End Sub

Private Sub StartButtonNew_Click()
' hides the file input box
    UserForm1.Hide
' start the progress bar box
    UserForm2.Show vbModeless ' this way it will not make macro wait for a close to continue running
    UpdateProgress (0.05) '  progress @ 0%
    
' executes a final file check code before running the NEW DRB PROCESS
' checks the strings given by the user are valid, opens target wb, and
' saves wb name to fName array along with workbook open function to the
' wb array for easy calling to switch windows.

    Dim x As Integer
    x = 1
    While x < 5
        If fPath(x) <> False Then
            Application.Workbooks.Open (fPath(x))
            Set Wb(x) = Workbooks.Open(fPath(x))
            ActiveSheet.Name = "Sheet1"
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
    UpdateProgress (0.05)  ' progress @ 5%
    DRBProcessNew
    Application.ScreenUpdating = True

' executes sub-macro to save newly created files
    SaveNewFiles

' Close the program/sub
End Sub

Sub DRBProcessNew()
'run the prep processes for the new drb
    XthetaPrep
    PalletPrepNew
' CLOSE THE B3, CA SETS, AND XTHETA FILES
    Wb(1).Close savechanges:=False ' B3 Data wb
    Wb(2).Close savechanges:=False ' CA Sets wb
    Wb(3).Close savechanges:=False ' Xtheta wb
    
End Sub

Private Sub StartButtonOld_Click()
'hides the file input box
    UserForm1.Hide
'start the progress bar box
    UserForm2.lblProgressbar.Width = 0 * UserForm2.LblBackground.Width 'start progress @ 0%
    UserForm2.Show vbModeless ' this way it won't make macro wait for a close to continue running

'executes a final file check code before running the OLD DRB PROCESS
'checks the strings given by the user are valid, opens target wb, and
'saves wb name to fName array along with workbook open function to the
'wb array for easy calling to switch windows.

    Dim x As Integer
    x = 1
    While x < 9
        If fPath(x) <> False Then
            Application.Workbooks.Open (fPath(x))
            Set Wb(x) = Workbooks.Open(fPath(x))
            ActiveSheet.Name = "Sheet1"
            fName(x) = ActiveWorkbook.Name
            fRootPath(x) = ActiveWorkbook.Path
            fRootPath(x) = fRootPath(x) & "\"
            fRange(x) = "'" & fRootPath(x) & "[" & fName(x) & "]Sheet1'!"
            If x = 4 Then
                fRange2(1) = "'" & fRootPath(x) & "[" & fName(x) & "]Sheet2'!"
            Else
            End If
            x = x + 1
        'if a file string is not set, disply error message and end program
        Else
            MsgBox "You did not specify a valid location or file for all of the required reports. The application will now terminate."
            End
        End If
    Wend
    end_time1 = Now() ' finished opening files and stroing names

'executes the main code for the OLD DRB PROCESS
    UpdateProgress (0.05)  ' progress @ 5%
    DRBProcessOld
    Application.ScreenUpdating = True
    SaveNewFiles
    
'Close the program/sub
End Sub

Sub DRBProcessOld()
'run each of the prep processes
    XthetaPrep
    UpdateProgress (0.2)  ' progress @ 25%
    Prep4R
    Prep2R
    CARPrep
    PalletPrep
    ' CLOSE THE B3 AND CA SETS FILES
    Wb(1).Close savechanges:=False ' B3 Data wb
    Wb(2).Close savechanges:=False ' CA Sets wb
    
    XthetaPrepAdjust
    UpdateProgress (0.45)   ' progress @ 50%

'build the WAR
    BuildWAR
    UpdateProgress (0.65)  ' progress @ 70%
    
    WAR_Vlookups
    UpdateProgress (0.75)  ' progress @ 80%
    
    MiscWARFormulas
    UpdateProgress (0.8)   ' progress @ 85%
    
    'Varify_WAR sub will go here
    'UserForm2.lblProgressbar.Width = 0.85 * UserForm2.LblBackground.Width  ' progress @ 90%
    
End Sub

Sub Prep4R()

'activate the 4R report
    SwitchWindows (6)
    
'runs text-to-columns on the article numbers to remove leading zeros and custom formats them back in
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Selection.NumberFormat = "00000000"
    
'inserts the uniqueID title and formula into a newly created column in the spreadsheet
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "UniqueID"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[7],RC[3],RC[1],RC[2])"
    
'find and replace the "-" in the entry number
    Columns("H:H").Select
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Range("A2").AutoFill Destination:=Range("A2:A" & lastrow), Type:=xlFillDefault
    
'selects the UniqueID column and copy/pastes special to lock in vlaues
    CopyPasteColumns ("A:A")

'filters and removes the rows that contain the subtotaled article counts
    With ws
        ActiveSheet.Range("A1:Q" & lastrow).autofilter Field:=10, Criteria1:="="
        LR = Cells(Rows.Count, 1).End(xlUp).Row
        If LR > 1 Then
            Range("A2:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
    End With

'clears all filters and sorts the active sheet by the 4R/CAR initial sort criteria of:
'Entry#,Exp#,Art#,Sup#,ExpQty,ExpDuty
    ActiveSheet.ShowAllData
    ActiveSheet.UsedRange.Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "H2:H" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "D2:D" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "B2:B" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "F2:F" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "K2:K" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1:L" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Application.ScreenUpdating = True

End Sub

Sub CARPrep()
'activate the CAR report
    SwitchWindows (7)
       
'runs text-to-columns on the article numbers to remove leading zeros and custom formats them back in
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Selection.NumberFormat = "00000000"
    
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
    
'insert the CAR multi-line 7501 set vlookup column far down the spreadsheet
'these line counts will help to drive the values and duties grabbed to avoids
'as many manual lookups to vet the set data.
    Range("CV1").Select
    ActiveCell.FormulaR1C1 = "CAR Line Qty"
    Range("CV2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-98],'\\DSUS061-NT0002.ikea.com\Common\Compliance\DISCUS\Reports\WORKING DRB CLAIMS\[CAR Multi-7501 Line Set Corrections.xls]Sheet1'!C1:C2,2,0),"""")"
    Range("CV2").AutoFill Destination:=Range("CV2:CV" & lastrow), Type:=xlFillDefault
' copy and paste values filled into new columns
    CopyPasteColumns ("CV:CV")

' adjusts the CAR values and duty paid to reflect the correct number of lines based on the the above vlookup from the
' DB of line counts verified through manual lookups for a given article on the 7501
    For Counter = 2 To lastrow
        If (Worksheets(1).Cells(Counter, 100).Value = """") Then ' do nothing. not a recorded set issue yet.
        ElseIf (Worksheets(1).Cells(Counter, 100).Value = 2) Then
            Worksheets(1).Cells(Counter, 46).Value = 0
            Worksheets(1).Cells(Counter, 47).Value = 0
            Worksheets(1).Cells(Counter, 52).Value = 0
            Worksheets(1).Cells(Counter, 53).Value = 0
        ElseIf (Worksheets(1).Cells(Counter, 100).Value = 3) Then
            Worksheets(1).Cells(Counter, 52).Value = 0
            Worksheets(1).Cells(Counter, 53).Value = 0
        ElseIf (Worksheets(1).Cells(Counter, 100).Value >= 4) Then ' do nothing. use norm calc.
        End If
    Next Counter
            
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
    Selection.autofilter

'filters and removes the rows that contain the blanks for DISCUS numbers
    With ws
        ActiveSheet.Range("A1:CA" & lastrow).autofilter Field:=4, Criteria1:="="
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

Sub Prep2R()

'activate the 2R report
    SwitchWindows (5)
   
'runs text-to-columns on the article numbers to remove leading zeros and custom formats them back in
    Columns("F:F").Select
    On Error Resume Next
        Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
    On Error GoTo 0
    
    Selection.NumberFormat = "00000000"

'find and replace the "-" in the entry number
    Columns("A:A").Select
    On Error Resume Next
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    On Error GoTo 0
    
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
'inserts the uniqueID title and formula into a newly created column in the spreadsheet
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "UniqueID"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[1],RC[6],RC[7])"
    Range("A2").AutoFill Destination:=Range("A2:A" & lastrow), Type:=xlFillDefault

'selects the UniqueID column and copy/pastes special to lock in vlaues
    CopyPasteColumns ("A:A")
    
'filters and removes the rows that contain the subtotaled article counts
    With ws
    'filter for summary rows and remove
        ActiveSheet.Range("A1:Q" & lastrow).autofilter Field:=11, Criteria1:="="
        LR = Cells(Rows.Count, 1).End(xlUp).Row
        If LR > 1 Then
            Range("A2:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        ActiveSheet.ShowAllData
'filter for blanks and remove rows
        ActiveSheet.Range("A1:Q" & lastrow).autofilter Field:=8, Criteria1:="="
        LR = Cells(Rows.Count, 1).End(xlUp).Row
        If LR > 1 Then
            Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If

    End With
    
    ActiveSheet.ShowAllData
    Application.ScreenUpdating = True

End Sub

Sub PalletPrep()
'-----------------------------------------------------------------------------------------------
'activate Pallet workbook
    SwitchWindows (4)
        
'insert the uniqueID title and formula into a newly created column in the spreadsheet (DISCUS/ART/SUP#)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "UniqueID - OrigEntry#"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[2],RC[4],RC[5])"
'fill down all of the newly created UniqueID columns and copy/pastes special to lock in vlaues
    Range("A2").AutoFill Destination:=Range("A2:A" & lastrow), Type:=xlFillDefault
    CopyPasteColumns ("A:A")
    
'-----------------------------------------------------------------------------------------------
'This part will populate the Pallet with the Actual DRB info based off of corrected DISCUS data
    
'inserts the "Actual DRB Claim Qty" title and formula into a new column in the spreadsheet
    Range("AF1").Select
    ActiveCell.FormulaR1C1 = "Actual Claim Qty"
    Range("AF2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]<>"""",RC[-7],RC[-10])"
'inserts the "Actual DRB Claim Qty" title and formula into a new column in the spreadsheet
    Range("AG1").Select
    ActiveCell.FormulaR1C1 = "Actual Dispatch Qty"
    Range("AG2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]<>"""",RC[-7],RC[-10])"
'inserts the "Actual DRB Claim Qty" title and formula into a new column in the spreadsheet
    Range("AH1").Select
    ActiveCell.FormulaR1C1 = "Actual Unload Qty"
    Range("AH2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]<>"""",RC[-7],RC[-10])"
'Fill down formulas and copy paste values
    Range("AF2:AH2").AutoFill Destination:=Range("AF2:AH" & lastrow), Type:=xlFillDefault
    CopyPasteColumns ("AF:AH")

'-----------------------------------------------------------------------------------------------
'This part will correct and  NAFTA Supplier "N" DRB status flags to blank
    Dim w
    Dim CellA
    Dim CellB

    w = 2
    While w <= lastrow
        CellA = Sheets(1).Cells(w, 30)
        CellB = Sheets(1).Cells(w, 31)
        
        If CellB = "NAFTA Supplier" Then
            With Sheets(1).Cells(w, 30)
                .Select
                .Activate
                Selection = ""
            End With
        End If
        w = w + 1
    Wend
'-----------------------------------------------------------------------------------------------
'This part will populate the Pallet with the Xtheta data
    Dim Xtheta_Range As String
    Xtheta_Range = fRange(3)
'inserts the Xtheta "DRB Claim Qty" title and formula into a new column in the spreadsheet
    Range("AI1").Select
    ActiveCell.FormulaR1C1 = "Xtheta DRB Claim Qty"
    Range("AI2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-5]<>""N"",VLOOKUP(RC[-34]," & Xtheta_Range & "C1:C11,8,0),"""")"
'inserts the Xtheta Invoice Qty" title and formula into a new column in the spreadsheet
    Range("AJ1").Select
    ActiveCell.FormulaR1C1 = "Xtheta Invoice Qty"
    Range("AJ2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-6]<>""N"",VLOOKUP(RC[-35]," & Xtheta_Range & "C1:C11,10,0),"""")"
'inserts the "Xtheta Entry Qty" title and formula into a new column in the spreadsheet
    Range("AK1").Select
    ActiveCell.FormulaR1C1 = "Xtheta Entry Qty"
    Range("AK2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]<>""N"",VLOOKUP(RC[-36]," & Xtheta_Range & "C1:C11,7,0),"""")"
    'inserts the "Xtheta Unload Qty" title and formula into a new column in the spreadsheet
    Range("AL1").Select
    ActiveCell.FormulaR1C1 = "Xtheta Unload Qty"
    Range("AL2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-8]<>""N"",VLOOKUP(RC[-37]," & Xtheta_Range & "C1:C11,11,0),"""")"
'Fill down formulas and copy paste values
    Range("AI2:AL2").AutoFill Destination:=Range("AI2:AL" & lastrow)
' Copy paste new values
    CopyPasteColumns ("AI:AL")
'-----------------------------------------------------------------------------------------------
'This part will populate the Pallet with the Canadian Sets data
    Dim Sets_Range As String
    Sets_Range = fRange(2)
'inserts the Xtheta "Canadian Sets" title and formula into a new column in the spreadsheet
    Range("AM1").Select
    ActiveCell.FormulaR1C1 = "Canadian Sets"
    Range("AM2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-9]<>""N"",IF((IFERROR(VLOOKUP(RC[-38]," & Sets_Range & "C1:C1,1,0),""N""))<>""N"",""Y"",""N""),"""")"
'Fill down formulas and copy paste values
    Range("AM2").AutoFill Destination:=Range("AM2:AM" & lastrow)
' Copy paste new values
    CopyPasteColumns ("AM:AM")
        
'-----------------------------------------------------------------------------------------------
'This part will populate the Pallet with the B3 Check data
    Dim B3_Range As String
    B3_Range = fRange(1)
'inserts the Xtheta "B3 Check" title and formula into a new column in the spreadsheet
    Range("AN1").Select
    ActiveCell.FormulaR1C1 = "B3 Check"
    Range("AN2").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-3]=0),IFERROR(VLOOKUP(CONCATENATE(RC[-37],RC[-35])," & B3_Range & "C1:C2,2,0),""""),"""")"
'Fill down formulas and copy paste values
    Range("AN2").AutoFill Destination:=Range("AN2:AN" & lastrow)
' Copy paste new values
    CopyPasteColumns ("AN:AN")
        
'-----------------------------------------------------------------------------------------------
'This part will populate the Pallet with the Check Flags and Final Send Flag

'insert the Xtheta "Lesser/Equal" title and formula into a new column in the spreadsheet
    Range("AO1").Select
    ActiveCell.FormulaR1C1 = "Lesser/Equal"
    Range("AO2").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-9]<=RC[-7],RC[-9]<=RC[-7]),""Lesser/Equal"","""")"
'insert the Xtheta "DRB vs B3" title and formula into a new column in the spreadsheet
    Range("AP1").Select
    ActiveCell.FormulaR1C1 = "DRB vs B3"
    Range("AP2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-2]="""",IF(RC[-10]<=RC[-5],""Correct"",""""),IF(RC[-10]<=RC[-2],""Correct"",""""))"
'insert the Xtheta "Final Send Flag" title and formula into a new column in the spreadsheet
    Range("AQ1").Select
    ActiveCell.FormulaR1C1 = "Final Send Flag"
    Range("AQ2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-11]=0,""DO NOT SEND"",IF(OR(AND(RC[-2]=""Lesser/Equal"",RC[-30]=""439STO"",RC[-13]=""""),AND(RC[-2]=""Lesser/Equal"",RC[-1]=""Correct"",RC[-13]="""")),""SEND"",""DO NOT SEND""))"
'Fill down formulas and copy paste values
    Range("AO2:AQ2").AutoFill Destination:=Range("AO2:AQ" & lastrow)
' Copy paste new values
    CopyPasteColumns ("AO:AQ")
'-----------------------------------------------------------------------------------------------
'This part will pivot the pallet report ActualClaimQty based on the SuperUniqueID and FinalSendFlag for the WAR
    
    ' insert the uniqueID title and formula into a newly created column in the spreadsheet (DISCUS/ART/SUP#)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Super UniqueID"
    
'-----------------------------------------------------------------------------------------------
' *** CHANGE THIS SECTION WITH EACH YEAR!! ***
' Create new column to pull in correct entry numbers based on In/Out Report
    Columns("T:T").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "In/Out Corrected Entry"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]=RC[-2],RC[-1],IF(RC[-1]<>RC[-2],VLOOKUP(RC[-1],'\\DSUS061-NT0002.ikea.com\Common\Compliance\DISCUS\Reports\WORKING DRB CLAIMS\[Entry Number Report.2005.xlsx]Entry Number Report.2005'!C5:C7,3,0),""""))"
'Fill down formulas and copy paste values
    Range("T2").AutoFill Destination:=Range("T2:T" & lastrow)
' Copy paste new values
    CopyPasteColumns ("T")
'-----------------------------------------------------------------------------------------------

' Popuate the Super Unique ID column using the new entry number data
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[19],RC[1])"
    ' fill down all of the newly created UniqueID columns and copy/pastes special to lock in vlaues
    Range("A2").AutoFill Destination:=Range("A2:A" & lastrow), Type:=xlFillDefault
' Copy paste new values
    CopyPasteColumns ("A:A")
    
    ' Creates a PivotTable report from the table on Sheet1 by using the PivotTableWizard
    ' method with the PivotFields method to specify the fields in the PivotTable.
    Dim objTable As PivotTable, objField As PivotField
    
    ' Select the sheet and first cell of the table that contains the data.
    ActiveWorkbook.Worksheets(1).Select
    Range("A1").Select
    
    ' Create the PivotTable object based on the Employee data on Sheet1.
    Set objTable = Worksheets(1).PivotTableWizard
    
    ' Specify row and column fields.
    Set objField = objTable.PivotFields("Super UniqueID")
    objField.Orientation = xlRowField
    Set objField = objTable.PivotFields("Final Send Flag")
    objField.Orientation = xlColumnField
    
    ' Specify a data field with its summary function and format.
    Set objField = objTable.PivotFields("Actual Claim Qty")
    objField.Orientation = xlDataField
    objField.Function = xlSum

'-----------------------------------------------------------------------------------------------
'This part will tidy up the spreadsheet for human use
'insert autofilters into report's first row
    
    ActiveWorkbook.Worksheets(2).Activate
    Rows("1:1").Select
    Selection.autofilter
    
    Cells.Select
    Range("AA1").Activate
    Cells.EntireColumn.AutoFit
    
    Rows("1:1").Select
    Range("AA1").Activate
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    Range("AH1:AJ1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Range("AK1:AN1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
        Range("AS1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Application.ScreenUpdating = True
    
End Sub
Sub PalletPrepNew()
'-----------------------------------------------------------------------------------------------
'activate Pallet workbook
    SwitchWindows (4)
        
'insert the uniqueID title and formula into a newly created column in the spreadsheet (DISCUS/ART/SUP#)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "UniqueID - OrigEntry#"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[2],RC[4],RC[5])"
'fill down all of the newly created UniqueID columns and copy/pastes special to lock in vlaues
    Range("A2").AutoFill Destination:=Range("A2:A" & lastrow), Type:=xlFillDefault
    CopyPasteColumns ("A:A")
    
'-----------------------------------------------------------------------------------------------
'This part will populate the Pallet with the Actual DRB info based off of corrected DISCUS data
    
'inserts the "Actual DRB Claim Qty" title and formula into a new column in the spreadsheet
    Range("AF1").Select
    ActiveCell.FormulaR1C1 = "Actual Claim Qty"
    Range("AF2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]<>"""",RC[-7],RC[-10])"
'inserts the "Actual DRB Claim Qty" title and formula into a new column in the spreadsheet
    Range("AG1").Select
    ActiveCell.FormulaR1C1 = "Actual Dispatch Qty"
    Range("AG2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]<>"""",RC[-7],RC[-10])"
'inserts the "Actual DRB Claim Qty" title and formula into a new column in the spreadsheet
    Range("AH1").Select
    ActiveCell.FormulaR1C1 = "Actual Unload Qty"
    Range("AH2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]<>"""",RC[-7],RC[-10])"
'Fill down formulas and copy paste values
    Range("AF2:AH2").AutoFill Destination:=Range("AF2:AH" & lastrow), Type:=xlFillDefault
    CopyPasteColumns ("AF:AH")

'-----------------------------------------------------------------------------------------------
'This part will correct and  NAFTA Supplier "N" DRB status flags to blank
    Dim w
    Dim CellA
    Dim CellB

    w = 2
    While w <= lastrow
        CellA = Sheets(1).Cells(w, 30)
        CellB = Sheets(1).Cells(w, 31)
        
        If CellB = "NAFTA Supplier" Then
            With Sheets(1).Cells(w, 30)
                .Select
                .Activate
                Selection = ""
            End With
        End If
        w = w + 1
    Wend
'-----------------------------------------------------------------------------------------------
'This part will populate the Pallet with the Xtheta data
    Dim Xtheta_Range As String
    Xtheta_Range = fRange(3)
'inserts the Xtheta "DRB Claim Qty" title and formula into a new column in the spreadsheet
    Range("AI1").Select
    ActiveCell.FormulaR1C1 = "Xtheta DRB Claim Qty"
    Range("AI2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-5]<>""N"",VLOOKUP(RC[-34]," & Xtheta_Range & "C1:C11,8,0),"""")"
'inserts the Xtheta Invoice Qty" title and formula into a new column in the spreadsheet
    Range("AJ1").Select
    ActiveCell.FormulaR1C1 = "Xtheta Invoice Qty"
    Range("AJ2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-6]<>""N"",VLOOKUP(RC[-35]," & Xtheta_Range & "C1:C11,10,0),"""")"
'inserts the "Xtheta Entry Qty" title and formula into a new column in the spreadsheet
    Range("AK1").Select
    ActiveCell.FormulaR1C1 = "Xtheta Entry Qty"
    Range("AK2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]<>""N"",VLOOKUP(RC[-36]," & Xtheta_Range & "C1:C11,7,0),"""")"
    'inserts the "Xtheta Unload Qty" title and formula into a new column in the spreadsheet
    Range("AL1").Select
    ActiveCell.FormulaR1C1 = "Xtheta Unload Qty"
    Range("AL2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-8]<>""N"",VLOOKUP(RC[-37]," & Xtheta_Range & "C1:C11,11,0),"""")"
'Fill down formulas and copy paste values
    Range("AI2:AL2").AutoFill Destination:=Range("AI2:AL" & lastrow)
' Copy paste new values
    CopyPasteColumns ("AI:AL")
'-----------------------------------------------------------------------------------------------
'This part will populate the Pallet with the Canadian Sets data
    Dim Sets_Range As String
    Sets_Range = fRange(2)
'inserts the Xtheta "Canadian Sets" title and formula into a new column in the spreadsheet
    Range("AM1").Select
    ActiveCell.FormulaR1C1 = "Canadian Sets"
    Range("AM2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-9]<>""N"",IF((IFERROR(VLOOKUP(RC[-38]," & Sets_Range & "C1:C1,1,0),""N""))<>""N"",""Y"",""N""),"""")"
'Fill down formulas and copy paste values
    Range("AM2").AutoFill Destination:=Range("AM2:AM" & lastrow)
' Copy paste new values
    CopyPasteColumns ("AM:AM")
        
'-----------------------------------------------------------------------------------------------
'This part will populate the Pallet with the B3 Check data
    Dim B3_Range As String
    B3_Range = fRange(1)
'inserts the Xtheta "B3 Check" title and formula into a new column in the spreadsheet
    Range("AN1").Select
    ActiveCell.FormulaR1C1 = "B3 Check"
    Range("AN2").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-3]=0),IFERROR(VLOOKUP(RC[-39]," & B3_Range & "C1:C2,2,0),""""),"""")"
'Fill down formulas and copy paste values
    Range("AN2").AutoFill Destination:=Range("AN2:AN" & lastrow)
' Copy paste new values
    CopyPasteColumns ("AN:AN")
        
'-----------------------------------------------------------------------------------------------
'This part will populate the Pallet with the Check Flags and Final Send Flag

'insert the Xtheta "Lesser/Equal" title and formula into a new column in the spreadsheet
    Range("AO1").Select
    ActiveCell.FormulaR1C1 = "Lesser/Equal"
    Range("AO2").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-9]<=RC[-7],RC[-9]<=RC[-7]),""Lesser/Equal"","""")"
'insert the Xtheta "DRB vs B3" title and formula into a new column in the spreadsheet
    Range("AP1").Select
    ActiveCell.FormulaR1C1 = "DRB vs B3"
    Range("AP2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-2]="""",IF(RC[-10]<=RC[-5],""Correct"",""""),IF(RC[-10]<=RC[-2],""Correct"",""""))"
'insert the Xtheta "Final Send Flag" title and formula into a new column in the spreadsheet
    Range("AQ1").Select
    ActiveCell.FormulaR1C1 = "Final Send Flag"
    Range("AQ2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-11]=0,""DO NOT SEND"",IF(OR(AND(RC[-2]=""Lesser/Equal"",RC[-30]=""439STO"",RC[-13]=""""),AND(RC[-2]=""Lesser/Equal"",RC[-1]=""Correct"",RC[-13]="""")),""SEND"",""DO NOT SEND""))"
'Fill down formulas and copy paste values
    Range("AO2:AQ2").AutoFill Destination:=Range("AO2:AQ" & lastrow)
' Copy paste new values
    CopyPasteColumns ("AO:AQ")
'-----------------------------------------------------------------------------------------------
'This part will pivot the pallet report ActualClaimQty based on the SuperUniqueID and FinalSendFlag for the WAR
    
    ' insert the uniqueID title and formula into a newly created column in the spreadsheet (DISCUS/ART/SUP#)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Super UniqueID"
    
'-----------------------------------------------------------------------------------------------
' *** CHANGE THIS SECTION WITH EACH YEAR!! ***
' Create new column to pull in correct entry numbers based on In/Out Report
    Columns("T:T").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "Corrected Entry No"
'-----------------------------------------------------------------------------------------------

' Popuate the Super Unique ID column using the new entry number data
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[18],RC[1])"
    ' fill down all of the newly created UniqueID columns and copy/pastes special to lock in vlaues
    Range("A2").AutoFill Destination:=Range("A2:A" & lastrow), Type:=xlFillDefault
' Copy paste new values
    CopyPasteColumns ("A:A")
    
    ' Creates a PivotTable report from the table on Sheet1 by using the PivotTableWizard
    ' method with the PivotFields method to specify the fields in the PivotTable.
    Dim objTable As PivotTable, objField As PivotField
    
    ' Select the sheet and first cell of the table that contains the data.
    ActiveWorkbook.Worksheets(1).Select
    Range("A1").Select
    
    ' Create the PivotTable object based on the Employee data on Sheet1.
    Set objTable = Worksheets(1).PivotTableWizard
    
    ' Specify row and column fields.
    Set objField = objTable.PivotFields("Super UniqueID")
    objField.Orientation = xlRowField
    Set objField = objTable.PivotFields("Final Send Flag")
    objField.Orientation = xlColumnField
    
    ' Specify a data field with its summary function and format.
    Set objField = objTable.PivotFields("Actual Claim Qty")
    objField.Orientation = xlDataField
    objField.Function = xlSum

'-----------------------------------------------------------------------------------------------
'This part will tidy up the spreadsheet for human use
'insert autofilters into report's first row
    
    ActiveWorkbook.Worksheets(2).Activate
    Rows("1:1").Select
    Selection.autofilter
    
    Cells.Select
    Range("AA1").Activate
    Cells.EntireColumn.AutoFit
    
    Rows("1:1").Select
    Range("AA1").Activate
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    Range("AH1:AJ1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Range("AK1:AN1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
        Range("AS1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Sub XthetaPrep()

'activate the xtheta report
    SwitchWindows (3)
   
'runs custom formats for leading zeros
    Columns("D:D").Select
    Selection.NumberFormat = "00000000"

'inserts the uniqueID title and formula into a newly created column in the spreadsheet
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "UniqueID"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[2],RC[4],RC[5])"
    Range("A2").AutoFill Destination:=Range("A2:A" & lastrow), Type:=xlFillDefault

' Copy paste new values
    CopyPasteColumns ("A:A")
    Application.ScreenUpdating = True

End Sub
Sub XthetaPrepAdjust()

'activate the xtheta report
    SwitchWindows (3)

'inserts the uniqueID title and formula into a newly created column in the spreadsheet
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[13],RC[2],RC[4],RC[5])"
    Range("A2").AutoFill Destination:=Range("A2:A" & lastrow), Type:=xlFillDefault
    
' Copy paste new values
    CopyPasteColumns ("A:A")

    Application.ScreenUpdating = True

End Sub

Sub BuildWAR()

    Dim CAR_Range As String
    CAR_Range = fRange(7)
'-----------------------------------------------------------------------------------------------
'This part will pull in all 4R data into the WAR report

'Select all 4R data to bring over
    Wb(6).Activate
    lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    Range("B2:K" & lastrow).Select
    Selection.Copy

'Paste 4R data into the WAR in the old section
    Wb(8).Activate
    Range("Q3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Paste 4R data into the WAR in the new section
    Range("BP3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Dim LastRowWAR
    LastRowWAR = ActiveSheet.Range("Q" & Rows.Count).End(xlUp).Row
    
'-----------------------------------------------------------------------------------------------
'This part will pull in all CAR data into the WAR report

'CAR Report main details entered into WAR
    SwitchWindows (7)
    Range("B2:Q" & lastrow).Select
    Selection.Copy
    
'Paste CAR data into the WAR in the main CAR section (ie. original CAR data)
    SwitchWindows (8)
    Range("AD3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Pull down the 2R/4R/CAR line match check flag
    Range("AT3:AW3").Select
    Selection.AutoFill Destination:=Range("AT3:AW" & LastRowWAR)
    
'CAR Report 'Errors' entered into WAR
    SwitchWindows (7)
    Range("S2:S" & lastrow).Select
    Selection.Copy
    SwitchWindows (8)
    Range("CA3").Select
    ActiveSheet.Paste
    
'CAR Report 'Protests' entered into WAR
    SwitchWindows (7)
    Range("AA2:AA" & lastrow).Select
    Selection.Copy
    SwitchWindows (8)
    Range("BZ3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'CAR Report 'theoretical values' entered into WAR
    SwitchWindows (7)
    Range("AC2:AE" & lastrow).Select
    Selection.Copy
    SwitchWindows (8)
    Range("CD3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'CAR Report 'HTS1' entered into WAR
    SwitchWindows (7)
    Range("AF2:AF" & lastrow).Select
    Selection.Copy
    SwitchWindows (8)
    Range("CB3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'CAR Report 'HTS2' entered into WAR
    SwitchWindows (7)
    Range("AL2:AL" & lastrow).Select
    Selection.Copy
    SwitchWindows (8)
    Range("CC3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'-----------------------------------------------------------------------------------------------
'This part will pull in all 2R data into the WAR report
    
'With the newly entered 4R/CAR data entered, sort the entire WAR by the following criteria:
'Entry #, Article #, Export Qty, 4R Duty @ 99%
    SwitchWindows (8)
    ActiveSheet.Range("A2:CN" & LastRowWAR).Select
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add Key:= _
        Range("W3:W" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add Key:= _
        Range("Q3:Q" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add Key:= _
        Range("U3:U" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add Key:= _
        Range("Z3:Z" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets(1).Sort
        .SetRange Range("A2:CN" & LastRowWAR)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Sort all the 2R data by the same above criteria as the WAR
    SwitchWindows (5)
    ActiveSheet.Range("A1:O" & lastrow).Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "B2:B" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "G2:G" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "J2:J" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "O2:O" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1:O" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Select and copy the now sorted 2R Data
    Range("B2:O" & lastrow).Select
    Selection.Copy

'Paste 2R data into the WAR in the old section
    SwitchWindows (8)
    Range("C3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Paste 2R data into the WAR in the new section
    Range("BB3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'-----------------------------------------------------------------------------------------------
'This part will populate the WAR with the row counts and uniqueID
    
'Fill down the two formulas and then copy paste to lock in values
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[32],RC[31],RC[29],RC[30])"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "=ROW(RC[-1])-2"
    Range("A3:B3").AutoFill Destination:=Range("A3:B" & LastRowWAR)
    Range("A3:B" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

'Enter the vlookup info for the duty values from the CAR report to pull into WAR
    Range("CG3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-84]," & CAR_Range & "C1:C53,35,0)"
    Range("CH3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-85]," & CAR_Range & "C1:C53,41,0)"
    Range("CI3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-86]," & CAR_Range & "C1:C53,47,0)"
    Range("CJ3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-87]," & CAR_Range & "C1:C53,53,0)"
    Range("CG3:CH3").Select

'Fill down the formulas and then copy paste to lock in values
    Range("CG3:CJ3").AutoFill Destination:=Range("CG3:CJ" & LastRowWAR)
' Copy paste new values
    CopyPasteColumns ("CG:CJ")
 
End Sub

Sub WAR_Vlookups()
'-----------------------------------------------------------------------------------------------
'This part will pull in all Xtheta and Pallet Report data into the WAR report using a vlookup.
'Activate WAR workbook
    SwitchWindows (8)
    
    Dim Xtheta_Range As String
    Dim Pallet_Range As String
    Xtheta_Range = fRange(3)
    Pallet_Range = fRange2(1)
    
'Inserts the Xtheta vlookup formula
    Range("AX3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-49]," & Xtheta_Range & "C1:C8,8,0),""No Hit"")"
    
'Inserts the Pallet vlookup formula
    Range("AY3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-50]," & Pallet_Range & "C1:C3,3,0),""No Hit"")"

'Insert the Export Qty Change check formulas
    Range("AZ3").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<RC[10],RC[-1],RC[10])"
    Range("BA3").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<RC[9],""Chng"","""")"

'Fill down the formulas and then copy paste to lock in values
    Range("AX3:BA3").Select
    Range("AX3:BA3").AutoFill Destination:=Range("AX3:BA" & lastrow)
' Copy paste new values
    CopyPasteColumns ("AX:BA")

'insert the CAR multi-line 7501 set vlookup column at end of spreadsheet
'these line counts will help to flag adjustments and to vet the set data.
    Range("DP3").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-112],'\\DSUS061-NT0002.ikea.com\Common\Compliance\DISCUS\Reports\WORKING DRB CLAIMS\[CAR Multi-7501 Line Set Corrections.xls]Sheet1'!C1:C2,2,0),"""")"
    Range("DP3:DP3").AutoFill Destination:=Range("DP3:DP" & lastrow)
' Copy paste new values
    CopyPasteColumns ("DP:DP")

' pull down formula to check if lines in claim are also found on a supplemental claim
    Range("DT3:DT3").AutoFill Destination:=Range("DT3:DT" & lastrow)
' Copy paste new values
    CopyPasteColumns ("DT:DT")
    
    Application.ScreenUpdating = True

End Sub
Sub MiscWARFormulas()
   
'----------------------------------------------------------------------------------------------
'This part will insert all misc forumlas, flags, and calculations into the WAR report.
    SwitchWindows (8)
    
' formula to calculate the correct amount of duty due @ 99% for all set lines
' excludes non multi tariff lines, XVV lines, and
    Range("CK3").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-8]<>"""",RC[-8]<>""V""),TRUNC(RC[-4]/RC[-79]*RC[-78]*0.99,2)+TRUNC(RC[-3]/RC[-79]*RC[-78]*0.99,2)+TRUNC(RC[-2]/RC[-79]*RC[-78]*0.99,2)+TRUNC(RC[-1]/RC[-79]*RC[-78]*0.99,2),"""")"
    Range("CK3:CK3").AutoFill Destination:=Range("CK3:CK" & lastrow)
    
' formula to populate the SQ comment field if Export Qty was changed
    Range("CP3").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-3]="""",RC[-41]=""Chng""),'Comment Key'!R6C2,"""")"
    Range("CP3:CP3").AutoFill Destination:=Range("CP3:CP" & lastrow)
' Copy paste new values
    CopyPasteColumns ("CP:CP")
        
' formula to pull in correct 7501 line number from CAR into 2R
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "=RC[37]"
    Range("D3:D3").AutoFill Destination:=Range("D3:D" & lastrow)

' formula to pull in correct import qty from CAR into 2R
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "=RC[32]"
    Range("J3:J3").AutoFill Destination:=Range("J3:J" & lastrow)

' formula to pull in correct port number from CAR into 2R
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[29]<>"""",RC[29],RC[51])"
    Range("E3:E3").AutoFill Destination:=Range("E3:E" & lastrow)
' Copy paste new values
    CopyPasteColumns ("E:E")
    
' formula to pull in correct export qty from corrected expt qty column
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "=RC[41]"
    Range("K3:K3").AutoFill Destination:=Range("K3:K" & lastrow)

' formula to pull in correct export qty from 2R to 4R expt qty column
    Range("U3").Select
    ActiveCell.FormulaR1C1 = "=RC[-10]"
    Range("U3:U3").AutoFill Destination:=Range("U3:U" & lastrow)
    
' formula to pull in correct total duty from CAR column into 2R
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "=RC[69]"
    Range("N3:N3").AutoFill Destination:=Range("N3:N" & lastrow)

' formula to pull in correct export qty from 2R to 4R total duty column
    Range("Y3").Select
    ActiveCell.FormulaR1C1 = "=RC[-11]"
    Range("Y3:Y3").AutoFill Destination:=Range("Y3:Y" & lastrow)
    
' formula to pull in correct UDV from CAR imported columns
    Range("M3").Select
    ActiveCell.FormulaR1C1 = "=IF(TRUNC(ABS(TRUNC(RC[69]/RC[-3],2)-RC[51]),2)<=0.01,RC[51],TRUNC(RC[69]/RC[-3],2))"
    Range("M3:M3").AutoFill Destination:=Range("M3:M" & lastrow)

' CAR duty rate
    Range("O3").Select
    ActiveCell.FormulaR1C1 = "=RC[69]"
    Range("O3:O3").AutoFill Destination:=Range("O3:O" & lastrow)
    
' formula to pull in correct 99% duty paid
    Range("P3").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[73]<>"""",RC[73],TRUNC(RC[-2]/RC[-6]*RC[-5]*.99,2))"
    Range("P3:P3").AutoFill Destination:=Range("P3:P" & lastrow)

' formula to pull in correct 99% duty paid from 2R to 4R column
    Range("Z3").Select
    ActiveCell.FormulaR1C1 = "=RC[-10]"
    Range("Z3:Z3").AutoFill Destination:=Range("Z3:Z" & lastrow)

' pull down the difference formula
    Range("AC3:AC3").AutoFill Destination:=Range("AC3:AC" & lastrow)

' pull down all comments columns
    Range("CM3:CO3").AutoFill Destination:=Range("CM3:CO" & lastrow)
    Range("CQ3:DB3").AutoFill Destination:=Range("CQ3:DB" & lastrow)
        
'Compare export and import dates. If export < imp, flag "N"
    Range("DQ3").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-101]>RC[-97],""Y"",""N"")"
    Range("DQ3:DQ3").AutoFill Destination:=Range("DQ3:DQ" & lastrow)
' Copy paste new values
    CopyPasteColumns ("DQ:DQ")
    
' pull down formula to check if export date is at 3yr time bar limit
    Range("DS3:DS3").AutoFill Destination:=Range("DS3:DS" & lastrow)

End Sub

Sub VerifyWAR()

    Dim CAR_Range As String
    Dim Xtheta_Range As String
    Dim Pallet_Range As String
   
    Application.ScreenUpdating = False
    
' ##### DEVELOP THIS PART OUT #####
'-----------------------------------------------------------------------------------------------
'This sub will search the various flags and checks previously run in the sub BuildWAR() for
'any errors, inconsistancies, or misalighnments in the lines of data. If all the data is
'correct the WAR will be saved for complaince-based checks and edits done by hand. If there
'are errors found, the macro will terminate with a message box alerting the user of the type
'and location of the error.
    SwitchWindows (8)

    Xtheta_Range = fRange(3)
    Pallet_Range = fRange(4)
    CAR_Range = fRange(7)

'Check the Pallet values returned for "N/A" and re-do vlookup based off the corrected UniqueID in Pallet
'    With ws
'        ActiveSheet.Range("A2:CR" & lastrow).autofilter Field:=51, Criteria1:="=No Hit"
'        LR = Cells(Rows.Count - 2, 1).End(xlUp).Row
'        If LR > 1 Then
'            Range("AY3").SpecialCells(xlCellTypeVisible).Select
'            ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-50]," & Pallet_Range & "C2:C33,32,0)"
'        End If
'    End With
'Activates the WAR report screen and defines the lastrow variable
        
'-----------------------------------------------------------------------------------------------
'This part will search the 2R/4R/CAR line duty and qty check flags to find any "N" and alert
'the user if found.
    Range("B2:K" & lastrow).Select
    Selection.Copy
    
    For lCount = 1 To lastrow
        Set rCell = Columns("1:2").Find(What:="N", After:=rCell, _
        LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False)
    Next lCount

'-----------------------------------------------------------------------------------------------
'This part will check the Xtheta values returned for "N/A" and alert user if found.
    With ws
        ActiveSheet.Range("A2:CR" & lastrow).autofilter Field:=51, Criteria1:="No Hit"
        LR = Cells(Rows.Count, 1).End(xlUp).Row
        If LR > 1 Then
            Range("AY2").SpecialCells(xlCellTypeVisible).Select
            ActiveCell.FormulaR1C1 = _
                "=VLOOKUP(RC[-50]," & Xtheta_Range & "C2:C33,32,0)"
            Range("AY2").SpecialCells(xlCellTypeVisible).AutoFill Destination:=Range("AY2:AY" & LR)
        End If
    End With

End Sub

Sub SaveNewFiles()

' ##### DEVELOP THIS PART OUT #####
' After opening, sorting, and preparing the various reports needed to work the DRB claim, save the worked file
' under its own name.

    UpdateProgress (0.97) ' progress @ 100%
    MsgBox "You report has been completed!" ' Show completion message box
    UserForm2.Hide ' hide the progress bar

End Sub

Sub SwitchWindows(ArrayNumber) ' Sub-routine will switch windows, get lastrow var, and turn off screen updating if on
    Wb(ArrayNumber).Activate
    Application.ScreenUpdating = False
    lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
End Sub

Sub CopyPasteColumns(Column)
    Columns(Column).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub FormulaPulldown(Cell, Formula)
    Range(Cell).Select
    ActiveCell.FormulaR1C1 = Formula
    Range(Cell).AutoFill Destination:=Range("Cell:Cell" & lastrow)
End Sub

Sub InsertFormula(Cell As String, Formula As String)
    Range(Cell).Select
    ActiveCell.FormulaR1C1 = Formula
End Sub

Sub UpdateProgress(Percentage)
Application.ScreenUpdating = True
UserForm2.lblProgressbar.Width = Percentage * UserForm2.LblBackground.Width ' progress @ some persentage
Application.ScreenUpdating = False
End Sub


