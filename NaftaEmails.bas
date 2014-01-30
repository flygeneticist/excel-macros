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
    ADS_dir = "\\DSUS061-NT0001\KEKEL1$\Desktop\NAFTA Sup Rpts\NAFTA_Sup_ADS_Art\Completed ADS\"
    
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
            .ReplyRecipients.Add "NAFTA@IKEA.COM"
            .To = SupContactID_1
            .CC = SupContactID_2
            .Subject = SupName & " - 2014 ADS"
            ' build HTML encoded string to embed into the email
            HTMLContent = "<HTML><BODY><p>Dear Suppliers,</p><p>It’s almost that time of year again. Time to renew our Article Detail Sheets (ADS) and to verify our article NAFTA eligibility for 2014. The IKEA NAFTA Group has implemented some changes aimed at claiming NAFTA benefits on as many articles as possible and streamlining the process while becoming more compliant.</p>"
            HTMLContent = HTMLContent + "<p>To that end, we will be asking for ADS sheets for ALL IKEA articles, either in production or soon to be in production, on a yearly basis. This means, even if the process and raw materials do not change, we will require an ADS sheet with an appropriate date and new signature each year. To assist in this huge undertaking, the NAFTA Group will be sending out newly revised ADS sheets starting in week 39 (September 26, 2013).</p>"
            HTMLContent = HTMLContent + "<p>These new ADS sheets will have the article numbers by IKEA Family name  pre-populated for your convenience, along with detailed instructions to assist you in providing the level of detail IKEA needs in order to review product eligibility and document product processes for U.S. Customs. These new ADS will be easier to complete and will have detailed instructions to assist you. We will require the ADS form to be filled out in the EXCEL format as well as a signed and dated copy in PDF format returned to us at <a href='mailto:NAFTA@IKEA.COM'>NAFTA@IKEA.COM</a> no later than week 42 (October 11, 2013).</p>"
            HTMLContent = HTMLContent + "<p>During weeks 43 and 44 our NAFTA team will verify the eligibility of all articles and begin sending out during week 45 (Nov 4,2013) pre-populated NAFTA forms for you to inspect sign, date and return no later than week 47 (Nov. 22, 2013)</p>"
            HTMLContent = HTMLContent + "<p>We do realize this is a large project and there will be challenges to meeting all the deadlines. This is why communication with each other will be of utmost importance. All correspondence with the NAFTA Group regarding NAFTA should be sent to the mailbox <a href='mailto:NAFTA@IKEA.COM'>NAFTA@IKEA.COM</a>. This is a group mailbox that will be monitored by the IKEA NAFTA team on a daily basis. Below is a time line for those who are visual in nature. We have attached a blank copy of the ADS sheet for you to look over and to familiarize yourself with.</p>"
            HTMLContent = HTMLContent + "<img src=""cid:NAFTA_Sup_Timeline.jpg""><br><p>We believe this new routine will allow us to improve the level of communication between the IKEA NAFTA Group, IKEA Trading and your our supplier in a way that will allow us to maximize the benefits of the NAFTA program offers. If you have any questions or concerns about any of the forms or information provided please do not hesitate to contact us</p><p>Best regards,</p>"
            HTMLContent = HTMLContent + "<p>IKEA NAFTA GROUP<br>IKEA Customs Compliance Center N.A.<br>100 IKEA Drive<br>Westampton, NJ 08060<br>609-261-1208<br><a href='mailto:NAFTA@IKEA.COM'>NAFTA@IKEA.COM</a><br></p></BODY></HTML>"
            
            .HTMLBody = HTMLContent
            .Attachments.Add "\\DSUS061-NT0001\KEKEL1$\Desktop\NAFTA Sup Rpts\NAFTA_Sup_Timeline.jpg"
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
