Attribute VB_Name = "CA_ICOMP_DC"
Public Sub CA_ICOMP_DC_InvoicePrep()
    Dim lastrow
    Dim main_wb
    main_wb = ActiveWorkbook.Name
    lastrow = Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
    
    Application.ScreenUpdating = False
    
' rearrange columns in the spreadsheet to relfect old macro arrangements
    ' Move 'net weight'
    Columns("F:F").Select
    Selection.Cut
    Columns("I:I").Select
    ActiveSheet.Paste
    ' move 'amount'
    Columns("G:G").Select
    Selection.Cut
    Columns("J:J").Select
    ActiveSheet.Paste
    ' move 'pr qty'
    Columns("D:D").Select
    Selection.Cut
    Columns("G:G").Select
    ActiveSheet.Paste
    
 ' inserts a row for the title descriptions
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
'insert autofilters into report's first row
    Rows("1:1").Select
    Selection.AutoFilter
    lastrow = Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
    
' Inserts new columns A and pulls down blank fixing formula
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=IF(IsBLANK(RC[1]),R[-1]C[0],RC[1])"
    Range("A2").AutoFill Destination:=Range("A2:A" & lastrow)
    Columns("A").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    
're-insert autofilters into report's first row
    Rows("1:1").Select
    Selection.AutoFilter
    
'filters and removes the rows that contain the subtotaled article counts
    With ws
    'filter for blanks and remove rows
        ActiveSheet.Range("A1:J" & lastrow).AutoFilter Field:=1, Criteria1:="="
        
        LR = Cells(Rows.Count, 1).End(xlUp).Row
        If LR > 1 Then
            Range("A2:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        ActiveSheet.ShowAllData
        
    'filter for Customs stat number rows and remove
        ActiveSheet.Range("A1:J" & lastrow).AutoFilter Field:=1, Criteria1:="=Customs*"
        LR = Cells(Rows.Count, 1).End(xlUp).Row
        If LR > 1 Then
            Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        ActiveSheet.ShowAllData
    End With
    lastrow = Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
    
' Lock in the correct item no into Col A
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[1],6)"
    Range("A2").AutoFill Destination:=Range("A2:A" & lastrow)
    Columns("A").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
' place formulas into columns C and F to pull descriptions and C/O into correct places
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(LEFT(RC[-1],1)<>""A""),R[1]C[-1],"""")"
    Range("C2").AutoFill Destination:=Range("C2:C" & lastrow)
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(LEFT(RC[-1],1)<>""0""),R[1]C[-1],"""")"
    Range("F2").AutoFill Destination:=Range("F2:F" & lastrow)
' place formulas into columns G to pull "Net Price" into correct place
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(LEFT(RC[-1],1)<>""0""),R[1]C[-1],"""")"
    Range("H2").AutoFill Destination:=Range("H2:H" & lastrow)
    Columns("H:H").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("G:G").Delete
' place formulas into columns K to pull "Amount" into correct place
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(LEFT(RC[-1],1)<>""0""),R[1]C[-1],"""")"
    Range("L2").AutoFill Destination:=Range("L2:L" & lastrow)
    Columns("L:L").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("K:K").Delete
    
' finalize formulas and delete the source column/s
    Columns("A:H").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    
' filters for column A where the art no is blank or '0'
    're-insert autofilters into report's first row
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A1:J" & lastrow).AutoFilter Field:=7, Criteria1:="="
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    If LR > 1 Then
        Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End If
    ActiveSheet.ShowAllData

    ActiveSheet.Range("A1:J" & lastrow).AutoFilter Field:=3, Criteria1:="="
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    If LR > 1 Then
        Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End If
    ActiveSheet.ShowAllData

    ActiveSheet.Range("A1:J" & lastrow).AutoFilter Field:=6, Criteria1:="0"
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    If LR > 1 Then
        Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End If
    ActiveSheet.ShowAllData

' moves CoO column into a more logical spot and inserts a new columns HTS #
    Columns("E:E").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Cut
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

' Run Text-to-Columns on the PCS and the attached QTYs (ONLY IF NEEDED!!)
    Range("E2").Select
    If "PCS" <> ActiveCell Then
        Columns("D:D").Select
        Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    End If
    lastrow = Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
    
' formating the unit price and total prices into humanly readable values
' generate titles for columns
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "C/O"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Dely Qty"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "UoM"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "HTS #"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "PR Qty"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Net Price"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Net Weight"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Total Amount"
    
' inserts new column to house the modified article numbers
    lastrow = Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Art No"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""IC"",RC[1])"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A" & lastrow), Type:=xlFillDefault
    Range("A2:A" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft

' re-calculate the total amount with 4 decimal points of precision
'   Range("J2").Select
'   ActiveCell.FormulaR1C1 = "="
    
' execute LOOKUPS
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
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-11],RC[-9])"
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
    Range("M2:Z2").AutoFill Destination:=Range("M2:Z" & lastrow), Type:=xlFillDefault
    Range("M2:Z" & lastrow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ' Pull in the Access article description, Classified By, and Classified Date into the spreadsheet
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
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
    
    ' move column K with the other descriptions and center columns I and J for easier reading
    Columns("K:K").Select
    Selection.Cut
    Columns("N:N").Select
    Selection.Insert Shift:=xlToRight
    Columns("A:L").Select
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
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Application.ScreenUpdating = True

End Sub

Sub DateFixer(column As Integer)
'This part will correct the "0" for the Date columns to be blank
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

