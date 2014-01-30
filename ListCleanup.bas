Attribute VB_Name = "ListCleanup"
Sub ListCleanup()

lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row

    With ActiveSheet
        ActiveSheet.Range("A2:AV" & lastrow).AutoFilter Field:=40, Criteria1:="<>Send"
        LR = Cells(Rows.Count, 1).End(xlUp).Row
        If LR > 2 Then
            Range("A2:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        ActiveSheet.ShowAllData
    End With

End Sub
