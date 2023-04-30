Sub SendEmails()

Dim olApp As Outlook.Application
Dim olMail As Outlook.MailItem
Dim rngCompanies As Range
Dim rngEmails As Range
Dim strCompany As String
Dim strEmail As String
Dim arrEmails() As String
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim strFolderPath As String
Dim strFileName As String
Dim bAttachmentsAdded As Boolean
Dim CurrentMonth As String

'Set the folder path where the PDF files are located
strFolderPath = "\\192.168.0.152\scan\deposit\Email_Monthly_Statement\"
CurrentMonth = Format(Date, "MMMM")
'Create a reference to the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Get a reference to the Outlook Application object
Set olApp = CreateObject("Outlook.Application")

'Loop through all the companies in column A
Set rngCompanies = Range("A2:A" & Range("A" & Rows.Count).End(xlUp).Row)
For Each cell In rngCompanies
    strCompany = cell.Value
    
    'Get the list of email addresses for this company
    strEmail = cell.Offset(0, 1).Value
    arrEmails = Split(strEmail, ";")
    
    'Create a new email message
    Set olMail = olApp.CreateItem(olMailItem)
    
    'Attach only the Exchange_rate.pdf file for this company
    Set objFolder = objFSO.GetFolder(strFolderPath)
    For Each objFile In objFolder.Files
        strFileName = objFile.Name
        If InStr(strFileName, "Exchange_rate.pdf") > 0 Then
            olMail.Attachments.Add objFile.Path
            bAttachmentsAdded = True
            Exit For
        End If
    Next
    
    'Send the email if attachments were added for the current company
    If bAttachmentsAdded Then
        'Set the email properties
        With olMail
            .Subject = "Exchange rate " & CurrentMonth & " from Busan Bank Ho Chi Minh City Branch"
            .To = strEmail
            '.Body = "Dear Customer, please find attached your monthly statement."
            .HTMLBody = "<font size='2' face='Arial'>" & _
                                "Dear Customer,<br><br>" & _
                                "We are writing to you the pdf file of the exchange rate of the month. Please find attached a PDF file with the details.<br><br>" & _
                                 "Sincerely,<br>--<br>Busan Bank- HCM City Branch<br>Room 1502, 15th Floor, MPlaza<br>39 Le Duan, Ben Nghe ward, Dist 1, HCMC<br>Tel: 028 7301 6200/6203 <br>F:028 7301 6201/028 3822 1143<br><br>" & _
                                "<i><font color='navy'>CONFIDENTIAL: This email and any files attached are intended solely for the use of the individual or entity to whom they are addressed and may contain confidential and/or privileged information. If you have received this email in error, please notify the sender immediately and delete it from your system. Any unauthorized use, dissemination, distribution, or copying of this email and its attachments is strictly prohibited. Thank you for your cooperation. </i> </font>" & _
                    "</font>"
            .Send
        End With
    End If
    
    'Clean up the email object
    Set olMail = Nothing
    
    'Reset the flag for attachments added
    bAttachmentsAdded = False
Next
    
    'Clean up the Outlook Application object
    Set olApp = Nothing
    
    'Clean up the FileSystemObject
    Set objFSO = Nothing
    
End Sub
