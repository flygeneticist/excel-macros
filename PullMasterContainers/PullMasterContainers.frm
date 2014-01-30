VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PullMasterContainers 
   Caption         =   "BS Master Spreadsheet Containers"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5670
   OleObjectBlob   =   "PullMasterContainers.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PullMasterContainers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------------------------------------------------------------------
' Title: IKEA Customs Specialist Container Assignment Software
' Created by: Kevin Keller
' Created on: April 03, 2013
' Last modified on: January 21, 2014
' Current Version: 1.6
'-----------------------------------------------------------------------------------------------

' set all public variables for program
Dim passwrd As String
Dim SpecialistInitials
Dim ContainersToPull As Integer
Dim MasterPulled As Integer
Dim SpecialistSpread
Dim wb(1 To 2) As Workbook
Dim lastrow As Integer
Dim MasterRow As Integer
Dim fpath(1 To 2)
Dim LR
Dim currdate As Date

Private Sub UserForm_Initialize()
    ' grab the user's ID from the OS system environment
    SpecialistInitials = (Environ$("Username"))
    ' setup vars inital state(s)
    passwrd = "RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$"
    ' setup current date var
    currdate = Date
    ' PATH TO THE MASTER SPREADSHEET. BE SURE TO EDIT IF MASTER IS MOVED ELSEWHERE!
    fpath(1) = "RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$"
End Sub

Private Sub ContainersToPullField_Change()
    If Not IsNumeric(ContainersToPullField.Value) Then  ' containers entered must be a number (no letters/symbols)
        MsgBox "This is a mandatory field." & vbCr & "Please enter a real, positive number only.", vbOKOnly, "Required Field"
    ElseIf ContainersToPullField.Value <= 0 Then  ' checks that entered container is a positive number
        MsgBox "This is a mandatory field." & vbCr & "Please enter a real, positive number only.", vbOKOnly, "Required Field"
    Else  ' passes all checks, set var == to user value
        ContainersToPull = ContainersToPullField.Value
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

    ' setup the specialist spreadsheet for the pull
    fpath(2) = ActiveWorkbook.Path
    Set wb(2) = ActiveWorkbook
    
    ' open and prep BS Master for macro
    TestFileOpened (fpath(1))
    Set wb(1) = Workbooks.Open(fpath(1))
    MasterSetup
        
    ' setup Master pulled to 0 to track lines pushed to BS Master
    MasterPulled = 0
    
    ' run main macro to assign containers to a specialist spreadsheet
    DistributeMasterContainers
    
    ' close the Master after pulling containers andupdating with initials
    wb(1).Close SaveChanges:=True
        
    Application.ScreenUpdating = True ' turns on screen updating again
        
    PullMasterContainers.Hide ' close the GUI window
    
    ' display a summary of the pull to the user
    MsgBox ("Container assignments are complete!" & vbCr & "" & vbCr & "You requested " & ContainersToPull & " containers." & vbCr & "There were " & LR & " containers available to pull." & vbCr & "You have recieved " & MasterPulled & " containers.")
End Sub

Sub DistributeMasterContainers()
    ' SETUP SPECIALIST SHEET FOR PULLING
    SwitchWindows (2) ' switch to Specialist Sheet
    'clear any filters
    On Error Resume Next
        ActiveSheet.ShowAllData
    On Error GoTo 0
    'setup last row tracking var with initial settings in the Specialists sheet
    MasterRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    SwitchWindows (1) ' switch to BS Master
    ' sort master by ETA and then by vessel name then by container number to make sure all like
    ' records are grouped together and that the most urgent containers are being pulled first.
    ActiveSheet.UsedRange.Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "I2:I" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "K2:K" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "E2:E" & lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1:Z" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' filters out any records with initials or FTZ
    ActiveSheet.Range("A1:AC" & lastrow).AutoFilter Field:=17, Criteria1:="="
    ActiveSheet.Range("A1:AC" & lastrow).AutoFilter Field:=18, Criteria1:="="
    ActiveSheet.Range("A1:AC" & lastrow).AutoFilter Field:=19, Criteria1:="<>*FTZ*"
     
    ' set temp lastrows for count of visible rows only
    LR = (ActiveSheet.AutoFilter.Range.Columns(5).SpecialCells(xlCellTypeVisible).Cells.Count) - 1
    If LR > 1 Then
        For Each x In ActiveSheet.AutoFilter.Range.Columns(5).SpecialCells(xlCellTypeVisible)
            If (x.Value <> "Container") Then
                If (MasterPulled < ContainersToPull) Then
                    ' add one to the current Specialist's spreadsheet LastRow tracking
                    MasterRow = MasterRow + 1
                    
                    ' mark row with current push date
                    Range("Q" & x.Row).Value = currdate
                    
                    ' mark row with specialists initials
                    Range("R" & x.Row).Value = SpecialistInitials
                    
                    'copy valid row data row to specialist spreadsheet
                    Range(Cells(x.Row, 1), Cells(x.Row, 19)).Copy
                    SwitchWindows (2)
                    Range(Cells(MasterRow, 1), Cells(MasterRow, 19)).Select
                    Selection.PasteSpecial
                    SwitchWindows (1)
                    
                    ' add one to the current containers pulled
                    MasterPulled = MasterPulled + 1
                Else
                    Exit For
                End If
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
    Application.ScreenUpdating = False
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
    errnum = Err           ' Save the error number that occurred.
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

Sub MasterSetup()
     ' switch to BS Master
    SwitchWindows (1)
    
    lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    With ActiveWorkbook
        Worksheets(1).Activate
        ' unlock the password protected sheet
        Worksheets(1).Unprotect passwrd
    End With
    
    'clear any filters
    On Error Resume Next
        ActiveSheet.ShowAllData
    On Error GoTo 0

End Sub

Function LockSheet(passwrd As String)
    'clear any filters
    On Error Resume Next
        ActiveSheet.ShowAllData
    On Error GoTo 0
    
    ' restore password protection to BS Master sheet
    ActiveSheet.Protect passwrd, True, True, True
End Function

