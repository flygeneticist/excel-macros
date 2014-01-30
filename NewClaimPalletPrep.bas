Attribute VB_Name = "NewClaimPrep"
Sub NewClaimPrep()

lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row

Columns("AH:AR").Select
Range("AR1").Activate
Selection.Delete Shift:=xlToLeft

Columns("A:B").Select
Range("B1").Activate
Selection.Delete Shift:=xlToLeft

'filters and removes the rows that contain the subtotaled article counts
    With ws
        ActiveSheet.Range("A1:AG" & lastrow).AutoFilter Field:=33, Criteria1:="SEND"
        LR = Cells(Rows.Count, 1).End(xlUp).Row
        If LR > 1 Then
            Range("A2:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
    End With

Range("AF2").Select

End Sub

