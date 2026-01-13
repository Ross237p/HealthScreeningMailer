Attribute VB_Name = "HealthScreeningMailer"
'=============================================================================
' Health Screening Invitation Mailer
' Version: 1.0
' Description: Send personalised HTML email invitations via Outlook from Excel
'=============================================================================
Option Explicit

' Constants for Outlook
Private Const olMailItem As Integer = 0
Private Const olFormatHTML As Integer = 2
Private Const olImportanceHigh As Integer = 2
Private Const olImportanceNormal As Integer = 1
Private Const olImportanceLow As Integer = 0

' CDO Configuration constants
Private Const cdoSendUsingPickup As Integer = 1
Private Const cdoSendUsingPort As Integer = 2
Private Const cdoAnonymous As Integer = 0
Private Const cdoBasic As Integer = 1
Private Const cdoNTLM As Integer = 2

' Module-level SMTP settings
Private m_SmtpServer As String
Private m_SmtpPort As Integer
Private m_SmtpUser As String
Private m_SmtpPass As String
Private m_FromAddress As String
Private m_UseCDO As Boolean

'=============================================================================
' Main Entry Point
'=============================================================================
Public Sub SendHealthScreeningInvites()
    Dim ws As Worksheet
    Dim headerRow As Range
    Dim dataRange As Range
    Dim currentRow As Range
    Dim outlookApp As Object
    Dim templatePath As String
    Dim htmlTemplate As String
    Dim sendMode As VbMsgBoxResult
    Dim rowCount As Long
    Dim sentCount As Long
    Dim errorCount As Long
    Dim skippedCount As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim toColIndex As Integer
    Dim subjectColIndex As Integer
    Dim statusColIndex As Integer

    On Error GoTo ErrorHandler

    ' Get active worksheet
    Set ws = ActiveSheet

    ' Find data boundaries
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Validate we have data
    If lastRow < 2 Then
        MsgBox "No data rows found in the spreadsheet. Please ensure row 1 contains headers and row 2+ contains data.", _
               vbExclamation, "No Data"
        Exit Sub
    End If

    ' Set header and data ranges
    Set headerRow = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))

    ' Validate required columns exist
    toColIndex = GetColumnIndex(headerRow, "To")
    subjectColIndex = GetColumnIndex(headerRow, "Subject")

    If toColIndex = 0 Then
        MsgBox "Required column 'To' not found in header row.", vbCritical, "Missing Column"
        Exit Sub
    End If

    If subjectColIndex = 0 Then
        MsgBox "Required column 'Subject' not found in header row.", vbCritical, "Missing Column"
        Exit Sub
    End If

    ' Check for Status column, add if missing
    statusColIndex = GetColumnIndex(headerRow, "Status")
    If statusColIndex = 0 Then
        lastCol = lastCol + 1
        ws.Cells(1, lastCol).Value = "Status"
        statusColIndex = lastCol
        Set headerRow = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    End If

    ' Prompt for HTML template file
    templatePath = Application.GetOpenFilename( _
        FileFilter:="HTML Files (*.html;*.htm),*.html;*.htm", _
        Title:="Select HTML Email Template")

    If templatePath = "False" Or templatePath = "" Then
        MsgBox "No template selected. Operation cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If

    ' Load HTML template
    htmlTemplate = LoadHTMLTemplate(templatePath)
    If htmlTemplate = "" Then
        MsgBox "Failed to load HTML template or template is empty.", vbCritical, "Template Error"
        Exit Sub
    End If

    ' Prompt for sending method
    Dim methodChoice As VbMsgBoxResult
    methodChoice = MsgBox("Which sending method would you like to use?" & vbCrLf & vbCrLf & _
                          "YES = Outlook (may be blocked by security)" & vbCrLf & _
                          "NO = SMTP Direct (bypasses Outlook security)" & vbCrLf & _
                          "CANCEL = Abort operation", _
                          vbYesNoCancel + vbQuestion, "Sending Method")

    If methodChoice = vbCancel Then
        MsgBox "Operation cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If

    m_UseCDO = (methodChoice = vbNo)

    ' If using CDO/SMTP, get server settings
    If m_UseCDO Then
        If Not GetSmtpSettings(headerRow) Then
            Exit Sub
        End If
    End If

    ' Prompt for send mode
    sendMode = MsgBox("How would you like to send the emails?" & vbCrLf & vbCrLf & _
                      "YES = Preview each email before sending" & vbCrLf & _
                      "NO = Send all emails automatically" & vbCrLf & _
                      "CANCEL = Abort operation", _
                      vbYesNoCancel + vbQuestion, "Send Mode")

    If sendMode = vbCancel Then
        MsgBox "Operation cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If

    ' Create Outlook application (needed for preview mode or Outlook sending)
    Set outlookApp = CreateObject("Outlook.Application")

    ' Initialize counters
    sentCount = 0
    errorCount = 0
    skippedCount = 0

    ' Process each row
    Application.ScreenUpdating = False
    Application.StatusBar = "Processing emails..."

    For rowCount = 2 To lastRow
        Set currentRow = ws.Range(ws.Cells(rowCount, 1), ws.Cells(rowCount, lastCol))

        ' Check if already sent
        If InStr(1, CStr(ws.Cells(rowCount, statusColIndex).Value), "Sent", vbTextCompare) > 0 Then
            skippedCount = skippedCount + 1
            GoTo NextRow
        End If

        ' Check for empty To field
        If Trim(CStr(ws.Cells(rowCount, toColIndex).Value)) = "" Then
            LogStatus ws, rowCount, statusColIndex, "Skipped: No recipient email"
            skippedCount = skippedCount + 1
            GoTo NextRow
        End If

        ' Validate email format
        If Not ValidateEmailAddress(CStr(ws.Cells(rowCount, toColIndex).Value)) Then
            LogStatus ws, rowCount, statusColIndex, "Error: Invalid email format"
            errorCount = errorCount + 1
            GoTo NextRow
        End If

        ' Process this row
        Application.StatusBar = "Processing row " & rowCount & " of " & lastRow & "..."
        DoEvents

        ' Create and send email
        If m_UseCDO Then
            If ProcessEmailRowCDO(htmlTemplate, currentRow, headerRow, ws, rowCount, statusColIndex, sendMode, outlookApp) Then
                sentCount = sentCount + 1
            Else
                errorCount = errorCount + 1
            End If
        Else
            If ProcessEmailRow(outlookApp, htmlTemplate, currentRow, headerRow, ws, rowCount, statusColIndex, sendMode) Then
                sentCount = sentCount + 1
            Else
                errorCount = errorCount + 1
            End If
        End If

        ' Add delay every 50 emails to avoid throttling
        If sentCount > 0 And sentCount Mod 50 = 0 Then
            Application.Wait Now + TimeValue("00:00:02")
        End If

NextRow:
    Next rowCount

    ' Cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Set outlookApp = Nothing

    ' Show summary
    MsgBox "Email processing complete!" & vbCrLf & vbCrLf & _
           "Sent: " & sentCount & vbCrLf & _
           "Skipped: " & skippedCount & vbCrLf & _
           "Errors: " & errorCount, _
           vbInformation, "Complete"

    Exit Sub

ErrorHandler:
    Dim errMsg As String
    errMsg = "Error: " & Err.Description & vbCrLf & _
             "Error number: " & Err.Number & vbCrLf & _
             "Processing row: " & rowCount
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Set outlookApp = Nothing
    MsgBox errMsg, vbCritical, "Error"
End Sub

'=============================================================================
' Process a single email row
'=============================================================================
Private Function ProcessEmailRow(outlookApp As Object, htmlTemplate As String, _
                                  dataRow As Range, headerRow As Range, _
                                  ws As Worksheet, rowNum As Long, _
                                  statusColIndex As Integer, sendMode As VbMsgBoxResult) As Boolean
    Dim mail As Object
    Dim processedHTML As String
    Dim toAddress As String
    Dim subject As String
    Dim ccAddress As String
    Dim bccAddress As String
    Dim sendAsAddress As String
    Dim importance As String
    Dim i As Integer
    Dim attachPath As String
    Dim stepName As String

    On Error GoTo RowError

    ProcessEmailRow = False
    stepName = "Reading cell values"

    ' Get email properties from row
    toAddress = Trim(GetCellValue(dataRow, headerRow, "To"))
    subject = Trim(GetCellValue(dataRow, headerRow, "Subject"))
    ccAddress = Trim(GetCellValue(dataRow, headerRow, "CC"))
    bccAddress = Trim(GetCellValue(dataRow, headerRow, "BCC"))
    sendAsAddress = Trim(GetCellValue(dataRow, headerRow, "SendAs"))
    importance = Trim(GetCellValue(dataRow, headerRow, "Importance"))

    stepName = "Processing template"
    ' Process HTML template with placeholders
    processedHTML = ReplacePlaceholders(htmlTemplate, dataRow, headerRow)

    stepName = "Creating mail item"
    ' Create mail item
    Set mail = outlookApp.CreateItem(olMailItem)

    stepName = "Setting mail properties"
    With mail
        ' Set body format first (important for HTMLBody to work correctly)
        .BodyFormat = olFormatHTML

        ' Set recipients
        .To = toAddress
        If ccAddress <> "" Then .CC = ccAddress
        If bccAddress <> "" Then .BCC = bccAddress

        ' Set subject
        .Subject = subject

        ' Set HTML body
        .HTMLBody = processedHTML

        ' Set importance
        Select Case UCase(importance)
            Case "HIGH"
                .importance = olImportanceHigh
            Case "LOW"
                .importance = olImportanceLow
            Case Else
                .importance = olImportanceNormal
        End Select

        stepName = "Setting shared mailbox"
        ' Set shared mailbox if specified
        If sendAsAddress <> "" Then
            SetSharedMailbox mail, outlookApp, sendAsAddress
        End If

        stepName = "Adding attachments"
        ' Add attachments
        For i = 1 To 10  ' Support up to 10 attachments
            attachPath = Trim(GetCellValue(dataRow, headerRow, "Attachment" & i))
            If attachPath <> "" Then
                If Dir(attachPath) <> "" Then
                    .Attachments.Add attachPath
                End If
            End If
        Next i

        ' Also check for single "Attachment" column
        attachPath = Trim(GetCellValue(dataRow, headerRow, "Attachment"))
        If attachPath <> "" Then
            If Dir(attachPath) <> "" Then
                .Attachments.Add attachPath
            End If
        End If

        stepName = "Resolving recipients"
        ' Resolve recipients to validate email addresses
        If Not .Recipients.ResolveAll Then
            ' If resolution fails, still try to send (external addresses may not resolve)
        End If

        stepName = "Sending/displaying email"
        ' Send or display based on mode
        If sendMode = vbYes Then
            .Display
            LogStatus ws, rowNum, statusColIndex, "Previewed: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
        Else
            ' Try to send, fall back to display if blocked by security
            On Error Resume Next
            .Send
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo RowError
                .Display
                LogStatus ws, rowNum, statusColIndex, "Displayed (auto-send blocked): " & Format(Now, "yyyy-mm-dd hh:mm:ss")
            Else
                On Error GoTo RowError
                LogStatus ws, rowNum, statusColIndex, "Sent: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
            End If
        End If
    End With

    Set mail = Nothing
    ProcessEmailRow = True
    Exit Function

RowError:
    LogStatus ws, rowNum, statusColIndex, "Error at [" & stepName & "]: " & Err.Description
    Set mail = Nothing
    ProcessEmailRow = False
End Function

'=============================================================================
' Load HTML template from file
'=============================================================================
Private Function LoadHTMLTemplate(filePath As String) As String
    Dim fileNum As Integer
    Dim content As String
    Dim textLine As String

    On Error GoTo FileError

    ' Check if file exists
    If Dir(filePath) = "" Then
        LoadHTMLTemplate = ""
        Exit Function
    End If

    ' Use native VBA file reading (more compatible)
    fileNum = FreeFile
    content = ""

    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, textLine
        content = content & textLine & vbCrLf
    Loop
    Close #fileNum

    LoadHTMLTemplate = content
    Exit Function

FileError:
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    LoadHTMLTemplate = ""
End Function

'=============================================================================
' Replace placeholders in template with row values
'=============================================================================
Private Function ReplacePlaceholders(template As String, dataRow As Range, headerRow As Range) As String
    Dim result As String
    Dim i As Integer
    Dim colName As String
    Dim placeholder As String
    Dim cellValue As String

    result = template

    For i = 1 To headerRow.Columns.Count
        colName = Trim(CStr(headerRow.Cells(1, i).Value))
        If colName <> "" Then
            ' Convert column name to placeholder format (spaces to underscores)
            placeholder = "{{" & Replace(colName, " ", "_") & "}}"
            ' Get cell value, handle empty cells
            cellValue = Trim(CStr(dataRow.Cells(1, i).Value))
            ' Replace all occurrences (case-insensitive)
            result = Replace(result, placeholder, cellValue, Compare:=vbTextCompare)
        End If
    Next i

    ReplacePlaceholders = result
End Function

'=============================================================================
' Set shared mailbox as sender
'=============================================================================
Private Sub SetSharedMailbox(mail As Object, outlookApp As Object, sendAsAddress As String)
    Dim acct As Object
    Dim accountFound As Boolean

    On Error Resume Next

    accountFound = False

    ' Try to find account in session accounts (for Send As)
    For Each acct In outlookApp.Session.Accounts
        If LCase(acct.SmtpAddress) = LCase(sendAsAddress) Then
            mail.SendUsingAccount = acct
            accountFound = True
            Exit For
        End If
    Next

    ' Fall back to SentOnBehalfOfName (for Send on Behalf Of)
    If Not accountFound Then
        mail.SentOnBehalfOfName = sendAsAddress
    End If

    On Error GoTo 0
End Sub

'=============================================================================
' Get column index by name
'=============================================================================
Private Function GetColumnIndex(headerRow As Range, columnName As String) As Integer
    Dim i As Integer

    For i = 1 To headerRow.Columns.Count
        If LCase(Trim(CStr(headerRow.Cells(1, i).Value))) = LCase(columnName) Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i

    GetColumnIndex = 0
End Function

'=============================================================================
' Get cell value by column name
'=============================================================================
Private Function GetCellValue(dataRow As Range, headerRow As Range, columnName As String) As String
    Dim colIndex As Integer

    colIndex = GetColumnIndex(headerRow, columnName)

    If colIndex > 0 Then
        GetCellValue = CStr(dataRow.Cells(1, colIndex).Value)
    Else
        GetCellValue = ""
    End If
End Function

'=============================================================================
' Log status to spreadsheet
'=============================================================================
Private Sub LogStatus(ws As Worksheet, rowNum As Long, statusColIndex As Integer, statusText As String)
    On Error Resume Next
    ws.Cells(rowNum, statusColIndex).Value = statusText
    On Error GoTo 0
End Sub

'=============================================================================
' Basic email format validation
'=============================================================================
Private Function ValidateEmailAddress(email As String) As Boolean
    Dim atPos As Integer
    Dim dotPos As Integer
    Dim cleanEmail As String

    ' Handle multiple emails separated by semicolons
    cleanEmail = Trim(Split(email, ";")(0))

    If cleanEmail = "" Then
        ValidateEmailAddress = False
        Exit Function
    End If

    ' Basic validation: must contain @ and at least one dot after @
    atPos = InStr(1, cleanEmail, "@")

    If atPos < 2 Then ' @ must not be first character
        ValidateEmailAddress = False
        Exit Function
    End If

    dotPos = InStr(atPos + 1, cleanEmail, ".")

    If dotPos = 0 Or dotPos = Len(cleanEmail) Then ' Must have dot after @ and not at end
        ValidateEmailAddress = False
        Exit Function
    End If

    ValidateEmailAddress = True
End Function

'=============================================================================
' Utility: Test email configuration (for debugging)
'=============================================================================
Public Sub TestEmailSetup()
    Dim outlookApp As Object
    Dim mail As Object
    Dim acct As Object
    Dim msg As String

    On Error GoTo TestError

    ' Create Outlook application
    Set outlookApp = CreateObject("Outlook.Application")

    msg = "Outlook connection: OK" & vbCrLf & vbCrLf
    msg = msg & "Available accounts:" & vbCrLf

    For Each acct In outlookApp.Session.Accounts
        msg = msg & "  - " & acct.SmtpAddress & vbCrLf
    Next

    ' Create test mail item
    Set mail = outlookApp.CreateItem(olMailItem)
    msg = msg & vbCrLf & "Mail item creation: OK" & vbCrLf

    Set mail = Nothing
    Set outlookApp = Nothing

    MsgBox msg, vbInformation, "Email Setup Test"
    Exit Sub

TestError:
    MsgBox "Error testing email setup:" & vbCrLf & _
           Err.Description, vbCritical, "Test Failed"
End Sub

'=============================================================================
' Get SMTP settings from user or spreadsheet
'=============================================================================
Private Function GetSmtpSettings(headerRow As Range) As Boolean
    Dim smtpColIndex As Integer
    Dim fromColIndex As Integer

    GetSmtpSettings = False

    ' Check if SMTP server is in spreadsheet
    smtpColIndex = GetColumnIndex(headerRow, "SmtpServer")
    fromColIndex = GetColumnIndex(headerRow, "FromAddress")

    ' Prompt for SMTP server
    m_SmtpServer = InputBox("Enter your SMTP server address:" & vbCrLf & vbCrLf & _
                            "Examples:" & vbCrLf & _
                            "  - smtp.office365.com (Office 365)" & vbCrLf & _
                            "  - smtp.yourcompany.com (Internal)" & vbCrLf & _
                            "  - mail.yourcompany.com", _
                            "SMTP Server", "smtp.office365.com")

    If m_SmtpServer = "" Then
        MsgBox "SMTP server is required for direct sending.", vbExclamation, "Cancelled"
        Exit Function
    End If

    ' Prompt for port
    Dim portStr As String
    portStr = InputBox("Enter SMTP port:" & vbCrLf & vbCrLf & _
                       "Common ports:" & vbCrLf & _
                       "  - 25 (Internal/no encryption)" & vbCrLf & _
                       "  - 587 (TLS - recommended)" & vbCrLf & _
                       "  - 465 (SSL)", _
                       "SMTP Port", "587")

    If portStr = "" Then
        MsgBox "SMTP port is required.", vbExclamation, "Cancelled"
        Exit Function
    End If
    m_SmtpPort = CInt(portStr)

    ' Prompt for From address
    m_FromAddress = InputBox("Enter the FROM email address:" & vbCrLf & vbCrLf & _
                             "This should be your shared mailbox or sending address.", _
                             "From Address", "")

    If m_FromAddress = "" Then
        MsgBox "From address is required.", vbExclamation, "Cancelled"
        Exit Function
    End If

    ' Ask about authentication
    Dim authChoice As VbMsgBoxResult
    authChoice = MsgBox("Does your SMTP server require authentication?" & vbCrLf & vbCrLf & _
                        "YES = Enter username/password" & vbCrLf & _
                        "NO = Anonymous/Windows auth", _
                        vbYesNo + vbQuestion, "Authentication")

    If authChoice = vbYes Then
        m_SmtpUser = InputBox("Enter SMTP username (usually your email):", "SMTP Username", m_FromAddress)
        If m_SmtpUser = "" Then
            MsgBox "Username is required for authenticated SMTP.", vbExclamation, "Cancelled"
            Exit Function
        End If

        m_SmtpPass = InputBox("Enter SMTP password:", "SMTP Password", "")
        If m_SmtpPass = "" Then
            MsgBox "Password is required for authenticated SMTP.", vbExclamation, "Cancelled"
            Exit Function
        End If
    Else
        m_SmtpUser = ""
        m_SmtpPass = ""
    End If

    GetSmtpSettings = True
End Function

'=============================================================================
' Process email row using CDO (SMTP direct)
'=============================================================================
Private Function ProcessEmailRowCDO(htmlTemplate As String, _
                                     dataRow As Range, headerRow As Range, _
                                     ws As Worksheet, rowNum As Long, _
                                     statusColIndex As Integer, sendMode As VbMsgBoxResult, _
                                     outlookApp As Object) As Boolean
    Dim cdoMsg As Object
    Dim cdoConfig As Object
    Dim processedHTML As String
    Dim toAddress As String
    Dim subject As String
    Dim ccAddress As String
    Dim bccAddress As String
    Dim stepName As String
    Dim i As Integer
    Dim attachPath As String

    On Error GoTo RowError

    ProcessEmailRowCDO = False
    stepName = "Reading cell values"

    ' Get email properties from row
    toAddress = Trim(GetCellValue(dataRow, headerRow, "To"))
    subject = Trim(GetCellValue(dataRow, headerRow, "Subject"))
    ccAddress = Trim(GetCellValue(dataRow, headerRow, "CC"))
    bccAddress = Trim(GetCellValue(dataRow, headerRow, "BCC"))

    stepName = "Processing template"
    processedHTML = ReplacePlaceholders(htmlTemplate, dataRow, headerRow)

    ' If preview mode, use Outlook to display
    If sendMode = vbYes Then
        stepName = "Creating preview in Outlook"
        Dim mail As Object
        Set mail = outlookApp.CreateItem(olMailItem)
        With mail
            .BodyFormat = olFormatHTML
            .To = toAddress
            If ccAddress <> "" Then .CC = ccAddress
            If bccAddress <> "" Then .BCC = bccAddress
            .Subject = subject
            .HTMLBody = processedHTML
            .Display
        End With
        Set mail = Nothing
        LogStatus ws, rowNum, statusColIndex, "Previewed: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
        ProcessEmailRowCDO = True
        Exit Function
    End If

    stepName = "Creating CDO message"
    Set cdoMsg = CreateObject("CDO.Message")
    Set cdoConfig = CreateObject("CDO.Configuration")

    stepName = "Configuring SMTP"
    With cdoConfig.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = m_SmtpServer
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = m_SmtpPort
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

        ' Set SSL/TLS based on port
        If m_SmtpPort = 465 Then
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        ElseIf m_SmtpPort = 587 Then
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        End If

        ' Set authentication if provided
        If m_SmtpUser <> "" Then
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = m_SmtpUser
            .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = m_SmtpPass
        End If

        .Update
    End With

    Set cdoMsg.Configuration = cdoConfig

    stepName = "Setting message properties"
    With cdoMsg
        .From = m_FromAddress
        .To = toAddress
        If ccAddress <> "" Then .CC = ccAddress
        If bccAddress <> "" Then .BCC = bccAddress
        .Subject = subject
        .HTMLBody = processedHTML

        stepName = "Adding attachments"
        ' Add attachments
        For i = 1 To 10
            attachPath = Trim(GetCellValue(dataRow, headerRow, "Attachment" & i))
            If attachPath <> "" Then
                If Dir(attachPath) <> "" Then
                    .AddAttachment attachPath
                End If
            End If
        Next i

        ' Also check single Attachment column
        attachPath = Trim(GetCellValue(dataRow, headerRow, "Attachment"))
        If attachPath <> "" Then
            If Dir(attachPath) <> "" Then
                .AddAttachment attachPath
            End If
        End If

        stepName = "Sending via SMTP"
        .Send
    End With

    Set cdoMsg = Nothing
    Set cdoConfig = Nothing

    LogStatus ws, rowNum, statusColIndex, "Sent (SMTP): " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    ProcessEmailRowCDO = True
    Exit Function

RowError:
    LogStatus ws, rowNum, statusColIndex, "Error at [" & stepName & "]: " & Err.Description
    Set cdoMsg = Nothing
    Set cdoConfig = Nothing
    ProcessEmailRowCDO = False
End Function
