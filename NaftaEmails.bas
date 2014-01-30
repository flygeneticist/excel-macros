Attribute VB_Name = "NaftaEmails"
Sub NaftaEmails()
    Application.ScreenUpdating = False
    
    Dim ADS_dir As String
    Dim strFile As String
    ' Workbook And Worksheet Variables
    Dim SupWrkBk As Workbook
    Dim SupWrkSht As Worksheet
    Dim rngRangeOfSups As Range
    
    ' Setup vars to call correct worksheet objects
    LineCount = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    StartingRow = 3
    Set SupWrkBk = ActiveWorkbook
    Set SupWrkSht = SupWrkBk.ActiveSheet
    Set rngRangeOfSups = Range(SupWrkSht.Cells(StartingRow, 2), SupWrkSht.Cells(LineCount, 2))
    ADS_dir = "RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$"
    
    ' Now that the variables are defined, write out a new email for each supplier
    ' setup Outlook App object hooks to build emails
    Dim OutApp As Object
    Dim OutMail As Object
    Dim OutAcct As Object
        
    For Each Sup In rngRangeOfSups
        ' setup supplier contact id(s) for email use
        SupRow = Sup.Row
        SupName = ActiveSheet.Cells(SupRow, 2).Value
        SupContactID_1 = ActiveSheet.Cells(SupRow, 5).Value
        SupContactID_2 = ActiveSheet.Cells(SupRow, 9).Value
        'generate email
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
        With OutMail
            ' Add new reply-to recipient to make it send back to the NAFA shared box
            .ReplyRecipients.Add "the email...."
            .To = SupContactID_1
            .CC = SupContactID_2
            .Subject = SupName & " - 2014 ADS"
            ' build HTML encoded string to embed into the email
            HTMLContent = "<HTML><BODY><p>Dear Suppliers,</p>.......<p></p></BODY></HTML>"
            
            .HTMLBody = HTMLContent
            .Attachments.Add "RFEGSWBE$TS%EYBE$^^$@#$SDZVDFGNTY#W$"
            strFile = Dir(ADS_dir & SupName & "*.xlsx")
            Do While Len(strFile) > 0
                .Attachments.Add ADS_dir & strFile
                strFile = Dir
            Loop
            '.Display ' display the e-mail message (on screen only. must be saved, sent, or closed)
            '.Save ' saves the message for later editing (puts it in the Drafts)
            .Send ' sends the e-mail message (puts it in the Outbox)
        End With
    Next Sup
    
    Application.ScreenUpdating = True
End Sub
