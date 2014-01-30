Attribute VB_Name = "SplitWB"
Option Explicit
Sub ParseSupplierList()
'Based on selected column, data is filtered to individual workbooks
'workbooks are named for the split value
Dim LR As Long, Itm As Long, MyCount As Long, vCol As Long
Dim ws As Worksheet, MyArr As Variant, vTitles As String, SvPath As String

'Sheet with data in it
   Set ws = Sheets(1)

'Path to save files into (remember the final '\'!)
    SvPath = "\\DSUS061-NT0002.ikea.com\Common\Compliance\Customs Compliance NA\NAFTA\ADS Prep Files\NAFTA_Sup_ADS_Art\"

'Range where titles are across top of data, as string, data MUST
'have titles in this row, edit to suit your titles locale
    vTitles = "A1:AN1"
   
'Choose column to evaluate from, column A = 1, B = 2, etc.
   vCol = Application.InputBox("What column to split data by? " & vbLf _
        & vbLf & "(A=1, B=2, C=3, etc)", "Which column?", 1, Type:=1)
   If vCol = 0 Then Exit Sub

'Spot bottom row of data
   LR = ws.Cells(ws.Rows.Count, vCol).End(xlUp).Row

'Speed up macro execution
   Application.ScreenUpdating = False

'Get a temporary list of unique values from column A
    ws.Columns(vCol).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ws.Range("EE1"), Unique:=True

'Sort the temporary list
    ws.Columns("EE:EE").Sort Key1:=ws.Range("EE2"), Order1:=xlAscending, Header:=xlYes, _
       OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

'Put list into an array for looping (values cannot be the result of formulas, must be constants)
    MyArr = Application.WorksheetFunction.Transpose(ws.Range("EE2:EE" & Rows.Count).SpecialCells(xlCellTypeConstants))

'clear temporary worksheet list
    ws.Range("EE:EE").Clear

'Turn on the autofilter, one column only is all that is needed
    ws.Range(vTitles).AutoFilter

'Loop through list one value at a time
    For Itm = 1 To UBound(MyArr)
        ws.Range(vTitles).AutoFilter Field:=vCol, Criteria1:=MyArr(Itm)
        
        ws.Range("A1:A" & LR).EntireRow.Copy
        Workbooks.Add
        Range("A1").PasteSpecial xlPasteAll
        Cells.Columns.AutoFit
        MyCount = MyCount + Range("A" & Rows.Count).End(xlUp).Row - 1
        
        'collapse the supplier list of articles to an art/sup level getting rid of split CA/US flows
        Dim lastrow As Integer
        lastrow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
        Cells.Select
        ActiveSheet.Range("$A1:AC" & lastrow).RemoveDuplicates Columns:=Array(1), _
            Header:=xlYes
        
        'saves new wb as an xlsx named after the split value and closes afterward
        ActiveWorkbook.SaveAs SvPath & MyArr(Itm), 51
        ActiveWorkbook.Close False
        
        ws.Range(vTitles).AutoFilter Field:=vCol
    Next Itm

'Cleanup
    ws.AutoFilterMode = False
    MsgBox "Rows with data: " & (LR - 1) & vbLf & "Rows copied to other sheets: " & MyCount & vbLf & "Hope they match!!"
    Application.ScreenUpdating = True
End Sub


