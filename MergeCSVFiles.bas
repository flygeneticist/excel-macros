Attribute VB_Name = "MergeCSVFiles"
' Start Code

Declare Function OpenProcess Lib "kernel32" _
                             (ByVal dwDesiredAccess As Long, _
                              ByVal bInheritHandle As Long, _
                              ByVal dwProcessId As Long) As Long

Declare Function GetExitCodeProcess Lib "kernel32" _
                                    (ByVal hProcess As Long, _
                                     lpExitCode As Long) As Long

Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ACTIVE = &H103

    Public Sub ShellAndWait(ByVal PathName As String, Optional WindowState)
    Dim hProg As Long
    Dim hProcess As Long, ExitCode As Long
    'fill in the missing parameter and execute the program
    If IsMissing(WindowState) Then WindowState = 1
    hProg = Shell(PathName, WindowState)
    'hProg is a "process ID under Win32. To get the process handle:
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, hProg)
    Do
        'populate Exitcode variable
        GetExitCodeProcess hProcess, ExitCode
        DoEvents
    Loop While ExitCode = STILL_ACTIVE
End Sub

Sub Merge_CSV_Files()
    Dim BatFileName As String
    Dim TXTFileName As String
    Dim XLSFileName As String
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim DefPath As String
    Dim wb As Workbook
    Dim oApp As Object
    Dim oFolder
    Dim foldername

    'Create two temporary file names
    BatFileName = Environ("Temp") & _
            "\CollectCSVData" & Format(Now, "dd-mm-yy-h-mm-ss") & ".bat"
    TXTFileName = Environ("Temp") & _
            "\AllCSV" & Format(Now, "dd-mm-yy-h-mm-ss") & ".txt"

    'Folder where you want to save the Excel file
    DefPath = Application.DefaultFilePath
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If

    'Set the extension and file format
    If Val(Application.Version) < 12 Then
        'You use Excel 97-2003
        FileExtStr = ".xls": FileFormatNum = -4143
    Else
        'You use Excel 2007
        FileExtStr = ".xlsx": FileFormatNum = 51
        'If you want to save as xls(97-2003 format) in 2007 use
        'FileExtStr = ".xls": FileFormatNum = 56
    End If

    'Name of the Excel file with a date/time stamp
    XLSFileName = DefPath & "MasterCSV " & _
    Format(Now, "dd-mmm-yyyy h-mm-ss") & FileExtStr

    'Browse to the folder with CSV files
    Set oApp = CreateObject("Shell.Application")
    Set oFolder = oApp.BrowseForFolder(0, "Select folder with CSV files", 512)
    If Not oFolder Is Nothing Then
        foldername = oFolder.Self.Path
        If Right(foldername, 1) <> "\" Then
            foldername = foldername & "\"
        End If

        'Create the bat file
        Open BatFileName For Output As #1
        Print #1, "Copy " & Chr(34) & foldername & "*.csv" _
        & Chr(34) & " " & TXTFileName
        Close #1

        'Run the Bat file to collect all data from the CSV files into a TXT file
        ShellAndWait BatFileName, 0
        If Dir(TXTFileName) = "" Then
            MsgBox "There are no csv files in this folder"
            Kill BatFileName
            Exit Sub
        End If

        'Open the TXT file in Excel
        Application.ScreenUpdating = False
        Workbooks.OpenText filename:=TXTFileName, Origin:=xlWindows, StartRow _
        :=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=True, _
        Space:=False, Other:=False

        'Save text file as a Excel file
        Set wb = ActiveWorkbook
        Application.DisplayAlerts = False
        wb.SaveAs filename:=XLSFileName, FileFormat:=FileFormatNum
        Application.DisplayAlerts = True

        wb.Close savechanges:=False
        MsgBox "You find the Excel file here: " & vbNewLine & XLSFileName

        'Delete the bat and text file you temporary used
        Kill BatFileName
        Kill TXTFileName

        Application.ScreenUpdating = True
    End If
End Sub

' End code

