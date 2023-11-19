Sub Monthly_Interest_Payment()

    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    Dim rngCompanies As Range
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
    Dim PreviousMonth As String

    'Set the folder path where the PDF files are located
    strFolderPath = "\\192.168.0.152\scan\deposit\Email_statement_DAILY\"
    PreviousMonth = Format(DateAdd("m", -1, Date), "MMMM")
    
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
        
        'Set the flag for attachments added to False at the beginning of each loop
        bAttachmentsAdded = False

        'Attach all the PDF files for this company
        Set objFolder = objFSO.GetFolder(strFolderPath)
        For Each objFile In objFolder.Files
            strFileName = objFile.Name

            ' Remove the file extension from the file name
            Dim fileCompany As String
            fileCompany = Left(strFileName, Len(strFileName) - Len(".pdf"))

            ' Perform exact string comparison between the extracted company name and the company name in column A
            If StrComp(fileCompany, strCompany, vbTextCompare) = 0 Then
                olMail.Attachments.Add objFile.Path
                bAttachmentsAdded = True
                Exit For ' No need to continue searching if a match is found
            End If
        Next
        
        'Send the email if attachments were added for the current company
        If bAttachmentsAdded Then
            'Set the email properties
            With olMail
                .Subject = "Interest Payment " & PreviousMonth & " from Busan Bank Ho Chi Minh City Branch"
                .To = strEmail
                ' Rest of the email body and sending process remains unchanged
                ' ...
            End With
            ' Send the email
            olMail.Send
        End If
        
        'Clean up the email object
        Set olMail = Nothing

    Next
    
    'Clean up the Outlook Application object
    Set olApp = Nothing
    
    'Clean up the FileSystemObject
    Set objFSO = Nothing
    
End Sub
