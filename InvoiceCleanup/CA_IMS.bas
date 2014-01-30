Attribute VB_Name = "CA_IMS"
Public Sub CA_IMS_InvoicePrep()
    Application.ScreenUpdating = False
    
    Dim lastrow
    Dim main_wb
    main_wb = ActiveWorkbook.Name
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial
    lastrow = Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
    
    ' inserts a row for the title descriptions
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    'filters and removes the rows that contain the subtotaled article counts
    With ws
    'filter for goods value rows and remove
        lastrow = Worksheets(1).Range("E" & Rows.Count).End(xlUp).Row
        ActiveSheet.Range("A1:G" & lastrow).AutoFilter Field:=5, Criteria1:="=Goods Value"
        LR = Cells(Rows.Count, 5).End(xlUp).Row
        If LR > 1 Then
            Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
    ActiveSheet.ShowAllData
    
    'filter for blank rows and remove
        ActiveSheet.Range("A1:G" & lastrow).AutoFilter Field:=3, Criteria1:="="
        LR = Cells(Rows.Count, 1).End(xlUp).Row
        If LR > 1 Then
            Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
    ActiveSheet.ShowAllData
        
    'filter for blanks and remove rows
        Range("H2").Select
        ActiveCell.FormulaR1C1 = "=LEN(RC[-7])"
        Range("H2").AutoFill Destination:=Range("H2:H" & lastrow)
        Cells.Select
        Selection.AutoFilter
        Selection.AutoFilter
        ActiveSheet.Range("A1:H" & lastrow).AutoFilter Field:=1, Criteria1:="<>"
        ActiveSheet.Range("A1:H" & lastrow).AutoFilter Field:=8, Criteria1:=">3"
        LR = Cells(Rows.Count, 8).End(xlUp).Row
        If LR > 1 Then
            Range("A2:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        ActiveSheet.ShowAllData
        Columns("H:H").Delete
        
    End With
    lastrow = Worksheets(1).Range("C" & Rows.Count).End(xlUp).Row
    
    ' format the weights and CoO into usable formats with find and replacements
    Columns("C:C").Select
    Selection.Replace What:="Country of origin", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Net weight kg", Replacement:="", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" Customs stat No.", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    ' remove the spaces from each of the non-description fields
    ActiveSheet.Range("A1:H" & lastrow).AutoFilter Field:=1, Criteria1:="="
    Selection.Replace What:=" ", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    ActiveSheet.ShowAllData ' reveal all data after manipulations are completed
        
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Columns("D:F").Select
    Selection.NumberFormat = "General"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=LEN(RC[-2])"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=LEN(RC[-2])"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(R[1]C[-1]>5,R[1]C[-2]=0),CONCATENATE(RC[-3],"" "",R[1]C[-3]),RC[-3])"
    Range("D2:F2").AutoFill Destination:=Range("D2:F" & lastrow)
    Columns("D:F").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    ' re-take length of cells based on new descriptions for next round of deletions
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=LEN(RC[1])"
    Columns("E:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
     With ws
    'filter for extra description rows and remove
        lastrow = Worksheets(1).Range("D" & Rows.Count).End(xlUp).Row
        ActiveSheet.Range("A1:J" & lastrow).AutoFilter Field:=4, Criteria1:="0"
        ActiveSheet.Range("A1:J" & lastrow).AutoFilter Field:=5, Criteria1:=">9"
        LR = Cells(Rows.Count, 5).End(xlUp).Row
        If LR > 1 Then
            Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        ActiveSheet.ShowAllData
    End With
    Columns("C:E").Delete
    
    ' insert new columns for the weights, CoO, HTS#
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    ' formulas to pull the data into the main article line
    Columns("D:F").Select
    Selection.NumberFormat = "General"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=IF(LEN(RC[-2])<>0,R[1]C[-1],"""")"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-1]<>"""",LEN(RC[-3])<>0),R[2]C[-2],"""")"
    Range("D2:E2").Select
    Range("D2:E2").AutoFill Destination:=Range("D2:E" & lastrow)
    Columns("D:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
    With ws
    'filter for blank rows and remove
        ActiveSheet.Range("A1:J" & lastrow).AutoFilter Field:=4, Criteria1:="="
        LR = Cells(Rows.Count, 3).End(xlUp).Row
        If LR > 1 Then
            Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        ActiveSheet.ShowAllData
    End With
    
    ' generate titles for columns
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Art No"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Invoice Description"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "C/O"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Dely Qty"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "UoM"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "HTS #"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "PR Qty"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Net Price"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Net Weight"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Total Amount"
    
    ' execute LOOKUPS into the invoice from the Access Database
    ' in foreign wb
    Dim SourceDataWB
    Workbooks.Open ("RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$")
    SourceDataWB = ActiveWorkbook.Name
    lastrow = Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[1],IF(RC[2]=""--"","""",RC[2]))"
    Selection.AutoFill Destination:=Range("B2:B" & lastrow), Type:=xlFillDefault

    ' in home workbook
    Workbooks(main_wb).Activate
    lastrow = Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
    ' create unique ID to link back to the data_export wb
    Columns("L:L").Select
    Selection.NumberFormat = "General"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""IM"",RC[-10],RC[-7])"
    Selection.AutoFill Destination:=Range("L2:L" & lastrow), Type:=xlFillDefault
    Range("L2:L" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ' pull in HTS_no
    Columns("F:F").Select
    Selection.NumberFormat = "General"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[6],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C8,7,0)"
    Selection.AutoFill Destination:=Range("F2:F" & lastrow), Type:=xlFillDefault
    Range("F2:F" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ' pull in cust_descrip, other_descrip, vender_name, add/cvd(case,rate,date), add info(date), notes
    Columns("M:Z").Select
    Selection.NumberFormat = "General"
    ' customs description
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Cust_Descrip"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,5,0)"
    ' other description
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Other_Descrip"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,37,0)"
    ' vendor name
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Vendor_Name"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,6,0)"
    ' CA ruling no/date
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Ruling_No"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,8,0)"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Ruling_Date"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,9,0)"
    ' CA Notes + notes_date
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Notes"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-6],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,10,0)"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "Notes_Date"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-7],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,11,0)"
    ' CA Sigma
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "SIMA"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,17,0)"
    ' CA ADD Info (date,case_no,rate)
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "ADD_Date"
    Range("U2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-9],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,18,0)"
    Range("V1").Select
    ActiveCell.FormulaR1C1 = "ADD_Case_No"
    Range("V2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-10],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,19,0)"
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "ADD_Rate"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-11],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,20,0)"
    ' CA CVD Info (date,case_no,rate)
    Range("X1").Select
    ActiveCell.FormulaR1C1 = "CVD_Date"
    Range("X2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-12],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,21,0)"
    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "CVD_Case_No"
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-13],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,22,0)"
    Range("Z1").Select
    ActiveCell.FormulaR1C1 = "CVD_Rate"
    Range("Z2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-14],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,23,0)"

    ' perform vlookups for all invoice rows and then remove formulas via copy paste special
    Range("M2:Z2").Select
    Selection.AutoFill Destination:=Range("M2:Z" & lastrow), Type:=xlFillDefault
    Columns("M:Z").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ' Pull in the Access article description, Classified By, and Classified Date into the spreadsheet
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Access Article Description"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[3],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,4,0)"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Classified By"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[2],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,31,0)"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Classified On"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],'RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$'!C2:C40,32,0)"
    Range("K2:M2").Select
    Selection.AutoFill Destination:=Range("K2:M" & lastrow), Type:=xlFillDefault
    Columns("K:M").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("N:N").Delete
    
    ' close data_export wb
    Workbooks(SourceDataWB).Close savechanges:=False
        
' Tidy up and format all cells to be uniform
    ' data formating issues with dates corrected
    DateFixer (13)
    DateFixer (18)
    DateFixer (20)
    DateFixer (22)
    DateFixer (25)
    
    ' move column K with the other descriptions and center columns A thru J for easier reading
    Columns("K:K").Select
    Selection.Cut
    Columns("O:O").Select
    Selection.Insert Shift:=xlToRight
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=""IM""&RC[1]"
    Selection.AutoFill Destination:=Range("A2:A" & lastrow), Type:=xlFillDefault
    Columns("A:A").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Delete
    
    Columns("A:J").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' text formatting/sizing, auto filters, and auto size columns
    Rows(1).Select
    Selection.Font.Bold = True
    Selection.AutoFilter
    Selection.AutoFilter
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 8
    End With
    Columns("A:J").AutoFit
    Range("A1").Select
    
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Application.ScreenUpdating = True

End Sub

Sub DateFixer(column As Integer)
'This part will correct the "0" for the Date columns to be blank and then format remaining values into European Date Formats
    With ActiveWorksheet
        Dim w
        Dim CellA
        w = 2
        lastrow = Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
        While w <= lastrow
            CellA = Sheets(1).Cells(w, column)
            On Error GoTo NextW
            If CellA = 0 Then
                With Sheets(1).Cells(w, column)
                    .Select
                    .Activate
                    Selection = ""
                End With
            End If
NextW:
Resume Next
            w = w + 1
        Wend
    
        Sheets(1).Columns(column).Select
        Selection.NumberFormat = "mm/dd/yyyy;@"
    End With
End Sub

