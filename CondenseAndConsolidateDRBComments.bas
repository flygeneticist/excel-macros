Attribute VB_Name = "Module21"
Sub ConsolidateAndCondenseErrorComments()

Application.ScreenUpdating = False
'Sort/Match column A values, merge all other cells into row format
Dim LR As Long, i As Long

'Sort data
    LR = Range("A" & Rows.Count).End(xlUp).Row
        
'Group matching entry numbers
    For i = LR To 2 Step -1
        If Cells(i, "A").Value = Cells(i - 1, "A").Value Then
            Range(Cells(i, "B"), Cells(i, Columns.Count).End(xlToLeft)).Copy _
                Cells(i - 1, Columns.Count).End(xlToLeft).Offset(0, 1)
            Rows(i).EntireRow.Delete (xlShiftUp)
        End If
    Next i
    
'make new column to hold the combined lines' comments
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Comment 1R"
    Range("B2").Select

'run the process to combine the comments for all lines in an entry into one cell
Dim icell As Long, lastrow As Long, lastcol As Long, iConc As Long
Dim myValue As String
 
'get length of newly condensed data
lastrow = Range("A" & Rows.Count).End(xlUp).Row

'iterate through all the entries, consolidating comments into column B
For icell = 2 To lastrow
    lastcol = Cells(icell, Columns.Count).End(xlToLeft).Column
    myValue = Cells(icell, 2).Value
        For iConc = 2 To lastcol
            myValue = myValue & Cells(icell, iConc).Value & "   "
        Next iConc
    Range("B" & icell).Value = myValue
    
Next icell

Application.ScreenUpdating = True

End Sub
