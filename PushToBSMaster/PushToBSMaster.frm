VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PushToBSMaster 
   Caption         =   "BS Master Spreadsheet Containers"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5055
   OleObjectBlob   =   "PushToBSMaster.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PushToBSMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------
' Title: IKEA Customs BS Specialist Push To BS Master Macro
' Created by: Kevin Keller
' Created on: January 21, 2014
' Last modified on: January 24, 2014
' Version: 0.5
'-----------------------------------------------------------------------------------------------

' set all public variables for program
Dim password As String
Dim SpecialistInitials
Dim wb(1 To 2) As Workbook
Dim fpath(1 To 2)
Dim lastrow
Dim LR
Dim currdate As Date
Dim MasterPushed As Integer
Dim MasterRow As Integer

Private Sub UserForm_Initialize()
    ' grab the user's ID from the OS system environment
    SpecialistInitials = (Environ$("Username"))
    ' setup vars inital state(s)
    password = "bs007"
    ' setup current date var
    currdate = Date
    ' PATH TO THE MASTER SPREADSHEET. BE SURE TO EDIT IF MASTER IS MOVED ELSEWHERE!
    fpath(1) = "\\DSUS061-FS0001\KEKEL1$\Desktop\BS MasterFY13.xlsx"
End Sub

Private Sub RangeButton_Click()
    RangeToPush = Application.GetOpenFilename(FileFilter:="Excel Files(*.xls*),*.xls*, CSV Files(*.csv),*.csv", FilterIndex:=1, Title:="Open File", MultiSelect:=False)
    If (RangeToPush <> False) Or (RangeToPush <> Null) Then
        ' setup the range to pull logic
    Else
        ' prompt user with an error message
        MsgBox ("Please select a valid file.")
    End If
End Sub

Private Sub CancelButton_Click()
    End
End Sub

Private Sub StartButton_Click()
    Main
End Sub

Sub Main()
    ' THIS IS THE MAIN LOOP FOR THE MACRO. IT IS ACTIVATED BY CLEARING THE USERFORM'S LOGIC CHECKS.
    Application.ScreenUpdating = True ' turns off screen updating to save time
    
    ' setup Master pulled to 0 to track lines pushed to BS Master
    MasterPushed = 0
        
    ' setup BS responcible's spreadsheet
    fpath(2) = ActiveWorkbook.Path
    Set wb(2) = ActiveWorkbook

    ' open and prep BS Master for macro
    TestFileOpened (fpath(1))
    Set wb(1) = Workbooks.Open(fpath(1))
    MasterSetup
    
    ' run the push macro
    Push_Execution
    
    ' switch back to Master sheet, lock up, and save changes
    SwitchWindows (1)
    LockSheet (password)
    wb(1).Close SaveChanges:=True  ' close and save the Master after pulling reference nums and updating with initials
    
    Application.ScreenUpdating = True ' turns on screen updating again
     
     ' closes the GUI window
    PushToBSMaster.Hide
    
    ' shows the user a quick summary of the pull
    MsgBox ("Your push has been completed!" & vbCr & "" & vbCr & "You have pushed " & MasterPushed & " lines to the BS Master.")
End Sub

Sub MasterSetup()
     ' switch to BS Master
    SwitchWindows (1)
    
    ' unlock the spreadsheet
    ActiveWorkbook.Unprotect password
    
    With ActiveWorkbook
        Worksheets(1).Activate
        ' unlock the password protected sheet
        Worksheets(1).Unprotect password
    End With
    
    
    'clear any filters
    On Error Resume Next
        ActiveSheet.ShowAllData
    On Error GoTo 0
    
    ' filters out any rows with data in them already
    MasterRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
End Sub

Sub Push_Execution()
    ' switch to BS Responsible's spreadsheet
    SwitchWindows (2)
    lastrow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    
    ' filter the BS Responsible's spreadsheet on the status column to bring up rows marked with an X/x
    ActiveSheet.Range("A1:AC" & lastrow).AutoFilter Field:=1, Criteria1:="<>"
    ActiveSheet.Range("A1:AC" & lastrow).AutoFilter Field:=2, Criteria1:="<>"
    ActiveSheet.Range("A1:AC" & lastrow).AutoFilter Field:=15, Criteria1:="X", Criteria2:="x"
    
    ' set temp lastrows for count of visible rows only (using the RefNums remaining as an index)
    LR = (ActiveSheet.AutoFilter.Range.Columns(15).SpecialCells(xlCellTypeVisible).Cells.Count) - 1 ' subtract 1 for the header row
    If LR > 1 Then ' if there are no rows marked do nothing
        ' otherwise, for each row marked
        For Each x In ActiveSheet.AutoFilter.Range.Columns(15).SpecialCells(xlCellTypeVisible)
            If x.Value <> "Status" Then
                ' add one to the current BS Master LastRow tracking
                MasterRow = MasterRow + 1
                
                ' add one to the current ref nums pulled
                MasterPushed = MasterPushed + 1
                
                'copy valid row data row to specialist spreadsheet
                Range(Cells(x.Row, 1), Cells(x.Row, 14)).Copy
                SwitchWindows (1)
                Range(Cells(MasterRow, 1), Cells(MasterRow, 14)).Select
                Selection.PasteSpecial
                ' mark row with current push date
                Range("O" & MasterRow).Select
                ActiveCell.Value = currdate
                ' mark row with specialists initials
                Range("P" & MasterRow).Select
                ActiveCell.Value = SpecialistInitials
                
                ' mark row in specialist sheet as pushed
                SwitchWindows (2)
                Range("O" & x.Row).Select
                ActiveCell.Value = "PUSHED"
            End If
        Next x
    End If
    
    'clear any filters
    On Error Resume Next
        ActiveSheet.ShowAllData
    On Error GoTo 0
End Sub

Sub SwitchWindows(ArrayNumber) ' Sub-routine will switch windows and turn off screen updating if on
    wb(ArrayNumber).Activate
    ActiveWorkbook.Worksheets(1).Activate
End Sub

Sub TestFileOpened(fpath As String)
    ' Test to see if the file is open.
    If IsFileOpen(fpath) Then
        ' Display a message stating the file in use.
        MsgBox "File already in use by another user." & vbCr & "Please try again in a few minutes!"
        End
    Else
        ' Open the file in Microsoft Excel.
        Workbooks.Open fpath
    End If
End Sub

Function IsFileOpen(filename As String)
' This function checks to see if a file is open or not. If the file is
' already open, it returns True. If the file is not open, it returns
' False. Otherwise, a run-time error occurs because there is
' some other problem accessing the file.
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open filename For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum
        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False
        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            IsFileOpen = True
        ' Another error occurred.
        Case Else
            Error errnum
    End Select
End Function

Function LockSheet(passwrd As String)
    'clear any filters
    On Error Resume Next
        ActiveSheet.ShowAllData
    On Error GoTo 0
    
    ' restore password protection to BS Master sheet
    ActiveSheet.Protect passwrd, True, True, True
End Function
