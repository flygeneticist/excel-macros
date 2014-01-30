VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TextileEmailsForm 
   Caption         =   "UserForm1"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5220
   OleObjectBlob   =   "TextileEmailsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TextileEmailsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim UserName As String
  Dim UserChoice As String
  
Private Sub FollowupEmail_Click()
    TextileEmailsForm.Hide
    UserChoice = "2"
    Call FollowupPrep
End Sub

Private Sub NewEmail_Click()
    TextileEmailsForm.Hide
    UserChoice = "1"
    Call FollowupPrep
End Sub

Sub FollowupPrep()
    Application.ScreenUpdating = False
    ' Copy all cells to a new sheet for working
    On Error Resume Next
    ActiveSheet.ShowAllData
    Cells.Select
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.PasteSpecial
    
    ' PUBLIC VAR OF THE USERNAME THE COMPUTER MACRO IS INSTALLED ON
    UserName = (Environ$("Username"))
    
    'insert autofilters into report's first row
    lastrow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    Rows("2:" & lastrow).Select
    Selection.AutoFilter
    
    rngRangeOfSups = Range(ActiveSheet.Cells(3, 7), ActiveSheet.Cells(lastrow, 7))
    
    If UserChoice = "1" Then
        'filters and removes the rows that do not contain the blanks for Request Dates
        With ActiveSheet
            lastrow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
            ActiveSheet.Range("A3:J" & lastrow).AutoFilter Field:=7, Criteria1:="<>"
            LR = Cells(Rows.Count, 1).End(xlUp).Row
            If LR > 2 Then
                Range("A3:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            End If
            ActiveSheet.ShowAllData
            
            lastrow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
            ActiveSheet.Range("A3:J" & lastrow).AutoFilter Field:=8, Criteria1:="<>"
            LR = Cells(Rows.Count, 1).End(xlUp).Row
            If LR > 2 Then
                Range("A3:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            End If
            ActiveSheet.ShowAllData
            
            lastrow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
            ActiveSheet.Range("A3:J" & lastrow).AutoFilter Field:=9, Criteria1:="<>"
            LR = Cells(Rows.Count, 1).End(xlUp).Row
            If LR > 2 Then
                Range("A3:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            End If
        End With
        NewRequestSupplierEmails
    
    ElseIf UserChoice = "2" Then
        'filters and removes the rows that do not contain the blanks for Received Dates
        With ActiveSheet
            lastrow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
            .Range("A3:J" & lastrow).AutoFilter Field:=7, Criteria1:="="
            LR = Cells(Rows.Count, 1).End(xlUp).Row
            If LR > 2 Then
                Range("A3:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            End If
            ActiveSheet.ShowAllData
            
            lastrow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
            .Range("A3:J" & lastrow).AutoFilter Field:=9, Criteria1:="="
            LR = Cells(Rows.Count, 1).End(xlUp).Row
            If LR > 2 Then
                Range("A3:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            End If
            ActiveSheet.ShowAllData
            
            Dim TodayDate As Date
            TodayDate = Date
            lastrow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
            .Range("A3:J" & lastrow).AutoFilter Field:=9, Criteria1:=">" & TodayDate
            LR = Cells(Rows.Count, 1).End(xlUp).Row
            If LR > 2 Then
                Range("A3:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            End If
            ActiveSheet.ShowAllData
            
            lastrow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
            .Range("A3:J" & lastrow).AutoFilter Field:=10, Criteria1:="<>" & UserName
            LR = Cells(Rows.Count, 1).End(xlUp).Row
            If LR > 2 Then
                Range("A3:A" & LR).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            End If
            ActiveSheet.ShowAllData
        End With
        FollowupRequestSupplierEmails
    Else
        MsgBox ("Not a valid choice! Macro will now terminate!")
        Application.DisplayAlerts = False
        ActiveWorkbook.Sheets(4).Delete
        Application.DisplayAlerts = True
        ActiveWorkbook.Sheets(1).Activate
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets(4).Delete
    Application.DisplayAlerts = True
    ActiveWorkbook.Sheets(1).Activate
    FrontPageUpdates
End Sub

Public Sub NewRequestSupplierEmails()
    'Recalculate the last row count
    On Error Resume Next
    ActiveSheet.ShowAllData
    lastrow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    
    ' Workbook And Worksheet Variables
    Dim SupWrkBk As Workbook
    Dim SupWrkSht As Worksheet
    
    ' Range Variables
    Dim rngRangeToSort As Range
    Dim rngRangeOfSups As Range
    Dim DataToSend As Range
    Dim c As Range
    
    ' Other Variables
    Dim strSingleSupplierWorkbookPath As String
    Dim StartingRow As Long
    Dim EndingRow As Long
    Dim strLastSupplierName As String
    Dim LineCount As Long
    
    ' Set The Workbook and Worksheet Variables
    Set SupWrkBk = ActiveWorkbook
    Set SupWrkSht = SupWrkBk.ActiveSheet
    
    'Clear Filters
    On Error Resume Next
    ActiveSheet.ShowAllData
    lastrow = Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
    
    
    ' Sort By SupNo And By ArtNo
    ActiveSheet.UsedRange.Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "D2:D" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "B2:B" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A2:K" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Locate The Last Data Line
    LineCount = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    StartingRow = 3
    HeadersRange = RangetoHTML(Range(SupWrkSht.Cells(2, 2), SupWrkSht.Cells(2, 6)))
    ' Now That The Worksheet Is Sorted, Write Out New Email For Each Unique Supplier Number
    Set rngRangeOfSups = Range(SupWrkSht.Cells(StartingRow, 4), SupWrkSht.Cells((LineCount + 1), 4))
    
    Dim OutApp As Object
    Dim OutMail As Object
    
    For Each Cell In rngRangeOfSups
        strLastSupplierName = Cell.Offset(-1, 0).Value
        If (Cell.Value <> strLastSupplierName) And (strLastSupplierName <> "Supplier") Then
            SupRow = Cell.Row - 1
            SupName = ActiveSheet.Cells(SupRow, 4).Value
            ' setup supplier contact id and name for email use
            ActiveSheet.Cells(SupRow, 11).Select
            ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-7],SupContacts!C1:C2,2,0)"
            SupContactID = ActiveSheet.Cells(SupRow, 11).Value
            ActiveSheet.Cells(SupRow, 12).Select
            ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],SupContacts!C1:C3,3,0)"
            SupContactName = ActiveSheet.Cells(SupRow, 12).Value
            EndingRow = Cell.Row
            'generate email
            Set DataToSend = Range(SupWrkSht.Cells(StartingRow, 2), SupWrkSht.Cells(EndingRow - 1, 6))
            Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.CreateItem(0)
            With OutMail
                .To = SupContactID
                .CC = ""
                .BCC = ""
                .Subject = SupName & " - Textile Confirmation Request"
                .HTMLBody = "<HTML><BODY><div><p>Dear " + SupContactName + ",</p><p>Please fill out the attached form with the full company name and address for the raw material (where the fabric was woven) and the manufacturer (where the good was assembled or finished) along with the actual processes for the following  articles/supplier.</p><p>Please note that this form is required to be filled out and  valid for 1 calendar year. We will only request the information when the form has expired at the end of the year.</p><p>***Please note that any time during the year the information provided on the form has changed you MUST provide a new form.***</p></div><div>" + HeadersRange + RangetoHTML(DataToSend) + "</div><div><p>Thank you for your cooperation.</p><p>Best regards,</p><p>Kevin Keller<br>Compliance Coordinator<br>IKEA Distribution Services Inc.<br>Phone: (609)261-1208 x2244</p></div></BODY></HTML>"
                .Attachments.Add ("\\DSUS061-FS0001.ikea.com\Common_A\Compliance\Customs Compliance NA\Textile MID\Textile Declaration Template\2014_Textile Declaration Form.xlsx")
                '.Display ' display the e-mail message.
                .Save ' saves the message for later editing
                '.Send ' sends the e-mail message (puts it in the Outbox)
            End With
            StartingRow = Cell.Row
        End If
    Next Cell
End Sub

Public Sub FollowupRequestSupplierEmails()
    'Recalculate the last row count
    'ActiveSheet.ShowAllData
    lastrow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    
    ' Workbook And Worksheet Variables
    Dim SupWrkBk As Workbook
    Dim SupWrkSht As Worksheet
    
    ' Range Variables
    Dim rngRangeToSort As Range
    Dim rngRangeOfSups As Range
    Dim DataToSend As Range
    Dim c As Range
    
    ' Other Variables
    Dim strSingleSupplierWorkbookPath As String
    Dim StartingRow As Long
    Dim EndingRow As Long
    Dim strLastSupplierName As String
    Dim LineCount As Long
    
    ' Set The Workbook and Worksheet Variables
    Set SupWrkBk = ActiveWorkbook
    Set SupWrkSht = SupWrkBk.ActiveSheet
    
    'Clear Filters
    On Error Resume Next
        ActiveSheet.ShowAllData
    lastrow = Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row
    
    ' Sort By SupNo And By ArtNo
    ActiveSheet.UsedRange.Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "D2:D" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "B2:B" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A2:K" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Locate The Last Data Line
    LineCount = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    StartingRow = 3
    HeadersRange = RangetoHTML(Range(SupWrkSht.Cells(2, 2), SupWrkSht.Cells(2, 6)))
    ' Now That The Worksheet Is Sorted, Write Out New Email For Each Unique Supplier Number
    Set rngRangeOfSups = Range(SupWrkSht.Cells(StartingRow, 4), SupWrkSht.Cells(LineCount + 1, 4))
    
    ' Set formatting of vlookup columns to general to allow formulas to flow through
    Columns("J:M").Select
    Selection.NumberFormat = "General"
    
    Dim OutApp As Object
    Dim OutMail As Object
        
    For Each Cell In rngRangeOfSups
        strLastSupplierName = Cell.Offset(-1, 0).Value
        If (Cell.Value <> strLastSupplierName) And (strLastSupplierName <> "Supplier") Then
            SupRow = Cell.Row - 1
            SupName = ActiveSheet.Cells(SupRow, 4).Value
            ' setup supplier contact id and name for email use
            ActiveSheet.Cells(SupRow, 11).Select
            ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-7],SupContacts!C1:C2,2,0)"
            SupContactID = ActiveSheet.Cells(SupRow, 11).Value
            ActiveSheet.Cells(SupRow, 12).Select
            ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],SupContacts!C1:C3,3,0)"
            SupContactName = ActiveSheet.Cells(SupRow, 12).Value
            EndingRow = Cell.Row
            SentDate = ActiveSheet.Cells(SupRow, 7).Text
            'generate email
            Set DataToSend = Range(SupWrkSht.Cells(StartingRow, 2), SupWrkSht.Cells(EndingRow - 1, 6))
            Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.CreateItem(0)
            With OutMail
                .To = SupContactID
                .CC = ""
                .BCC = ""
                .Subject = SupName & " - Textile Confirmation Request"
                .HTMLBody = "<HTML><BODY><div><p>Dear " & SupContactName & ",</p><p>I have still not received a completed textiles form for the articles originally sent out to you on " & SentDate & ". It is imperative that we have these forms completed to comply with US Customs' regulations. I have re-attached the blank form to this email. Please see the requested articles shown below:</p></div><div>" + HeadersRange + RangetoHTML(DataToSend) + "</div><div><p>If you will not be producing these articles, or not shipping them to the US, please let me know so that I can stop following up, as a textiles form would not be needed in those cases. However, if these articles will be shipped to the US, please send me the completed form ASAP. Thank you!</p><p>Best regards,</p><p>Kevin Keller<br>Compliance Coordinator<br>IKEA Distribution Services Inc.<br>Phone: (609)261-1208 x2244</p></div></BODY></HTML>"
                .Attachments.Add ("\\DSUS061-FS0001.ikea.com\Common_A\Compliance\Customs Compliance NA\Textile MID\Textile Declaration Template\2014_Textile Declaration Form.xlsx")
                '.Display ' display the e-mail message.
                .Save ' saves the message for later editing
                '.Send ' sends the e-mail message (puts it in the Outbox)
            End With
            StartingRow = Cell.Row
        End If
    Next Cell
End Sub

Sub FrontPageUpdates()
    Dim TodayDate As Date
    Dim FollowupDate
    TodayDate = Date
    FollowupDate = TodayDate + 7
    
    ActiveWorkbook.Sheets(1).Activate
    
    If UserChoice = "1" Then
        'filters for the rows that do not contain the blanks for Request Dates
        With ActiveSheet
            lastrow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
            ActiveSheet.Range("A3:J" & lastrow).AutoFilter Field:=7, Criteria1:="="
            ActiveSheet.Range("A3:J" & lastrow).AutoFilter Field:=8, Criteria1:="="
            ActiveSheet.Range("A3:J" & lastrow).AutoFilter Field:=9, Criteria1:="="
            LR = Cells(Rows.Count, 1).End(xlUp).Row
            If LR > 2 Then
                For Each x In ActiveSheet.AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible)
                    If x <> "Art No" Then
                        ' set currect cell values to store for checking against new currect value
                        Row = x.Row
                        ' mark row with orig request date (current date)
                        ActiveWorkbook.Sheets(1).Range("G" & Row).Activate
                        ActiveCell.Value = TodayDate
                        ' mark row with followup date (current date + 7)
                        ActiveWorkbook.Sheets(1).Range("I" & Row).Activate
                        ActiveCell.Value = FollowupDate
                        ' mark row with user initials
                        ActiveWorkbook.Sheets(1).Range("J" & Row).Activate
                        ActiveCell.Value = UserName
                    End If
                Next x
            End If
            ActiveSheet.ShowAllData
        End With
        
    ElseIf UserChoice = "2" Then
        'filters for the rows that need updating for followups
        With ActiveSheet
            lastrow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
            .Range("A3:J" & lastrow).AutoFilter Field:=7, Criteria1:="<>"
            .Range("A3:J" & lastrow).AutoFilter Field:=8, Criteria1:="="
            .Range("A3:J" & lastrow).AutoFilter Field:=9, Criteria1:="<=" & TodayDate
            .Range("A3:J" & lastrow).AutoFilter Field:=10, Criteria1:="=" & UserName
            LR = Cells(Rows.Count, 1).End(xlUp).Row
            If LR > 2 Then
                For Each x In ActiveSheet.AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible)
                    If x <> "Art No" Then
                        ' set currect cell values to store for checking against new currect value
                        Row = x.Row
                        ' mark row with followup date (current date + 7)
                        ActiveWorkbook.Sheets(1).Range("I" & Row).Activate
                        ActiveCell.Value = FollowupDate
                    End If
                Next x
            End If
            ActiveSheet.ShowAllData
        End With
    End If
    Exit Sub
End Sub

Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2013
    Dim FSO As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set ts = FSO.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set FSO = Nothing
    Set TempWB = Nothing
End Function

