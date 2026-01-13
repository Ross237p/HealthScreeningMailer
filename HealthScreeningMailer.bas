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

    ' Create Outlook application (late binding for compatibility)
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
        If ProcessEmailRow(outlookApp, htmlTemplate, currentRow, headerRow, ws, rowCount, statusColIndex, sendMode) Then
            sentCount = sentCount + 1
        Else
            errorCount = errorCount + 1
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
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Set outlookApp = Nothing
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Error number: " & Err.Number, vbCritical, "Error"
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

    On Error GoTo RowError

    ProcessEmailRow = False

    ' Get email properties from row
    toAddress = Trim(GetCellValue(dataRow, headerRow, "To"))
    subject = Trim(GetCellValue(dataRow, headerRow, "Subject"))
    ccAddress = Trim(GetCellValue(dataRow, headerRow, "CC"))
    bccAddress = Trim(GetCellValue(dataRow, headerRow, "BCC"))
    sendAsAddress = Trim(GetCellValue(dataRow, headerRow, "SendAs"))
    importance = Trim(GetCellValue(dataRow, headerRow, "Importance"))

    ' Process HTML template with placeholders
    processedHTML = ReplacePlaceholders(htmlTemplate, dataRow, headerRow)

    ' Create mail item
    Set mail = outlookApp.CreateItem(olMailItem)

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

        ' Set shared mailbox if specified
        If sendAsAddress <> "" Then
            SetSharedMailbox mail, outlookApp, sendAsAddress
        End If

        ' Add attachments
        For i = 1 To 10  ' Support up to 10 attachments
            attachPath = Trim(GetCellValue(dataRow, headerRow, "Attachment" & i))
            If attachPath <> "" Then
                If Dir(attachPath) <> "" Then
                    .Attachments.Add attachPath
                Else
                    ' Log warning but continue
                    Debug.Print "Warning: Attachment not found - " & attachPath
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

        ' Send or display based on mode
        If sendMode = vbYes Then
            .Display
            LogStatus ws, rowNum, statusColIndex, "Previewed: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
        Else
            .Send
            LogStatus ws, rowNum, statusColIndex, "Sent: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
        End If
    End With

    Set mail = Nothing
    ProcessEmailRow = True
    Exit Function

RowError:
    LogStatus ws, rowNum, statusColIndex, "Error: " & Err.Description
    Set mail = Nothing
    ProcessEmailRow = False
End Function

'=============================================================================
' Load HTML template from file
'=============================================================================
Private Function LoadHTMLTemplate(filePath As String) As String
    Dim stream As Object
    Dim content As String

    On Error GoTo FileError

    ' Use ADODB.Stream for proper UTF-8 encoding support
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Charset = "utf-8"
        .Open
        .LoadFromFile filePath
        content = .ReadText
        .Close
    End With

    Set stream = Nothing
    LoadHTMLTemplate = content
    Exit Function

FileError:
    Set stream = Nothing
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
