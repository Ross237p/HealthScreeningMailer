# Development Instructions for Coding Agents

This document provides technical context and instructions for AI coding agents working on the Health Screening Invitation Mailer project.

## Project Context

### Business Problem

A healthcare company needs to send personalised email invitations to employees of client companies, inviting them to book health screenings. The emails must:

- Originate from a shared Outlook mailbox (not personal accounts)
- Support HTML formatting for professional appearance
- Include personalisation (name, company, invitation code)
- Support file attachments
- Maintain a log of sent invitations
- Work within a locked-down IT environment (no third-party software)

### Technical Constraints

- **No external dependencies**: Cannot install Python, Node.js, or any third-party software
- **Microsoft Office only**: Must use VBA within Excel/Outlook
- **Windows 11 compatibility**: Previous Word-based mail merge broke after OS upgrade
- **Healthcare security**: No cloud services or external data transmission
- **Shared mailbox**: Must use `SendOnBehalfOfName` or `SendUsingAccount` for shared inbox

### Previous Solution (Deprecated)

The previous solution used Word's mail merge with the `MailEnvelope` object. This approach:

- Required opening HTML in Word (which corrupts HTML with Word-specific markup)
- Used Word's MERGEFIELD syntax
- Relied on `Document.MailEnvelope.Item` to create Outlook emails
- Broke on Windows 11 due to COM/security changes

The new solution eliminates Word entirely.

## Architecture

### Component Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                        Excel Workbook                           │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │                    Recipient Data                         │  │
│  │  (To, Subject, Title, Surname, Company, Attachments...)   │  │
│  └───────────────────────────────────────────────────────────┘  │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │                    VBA Module                             │  │
│  │  - Read recipient data from active sheet                  │  │
│  │  - Load HTML template from file system                    │  │
│  │  - Replace placeholders with row data                     │  │
│  │  - Create Outlook MailItem via COM                        │  │
│  │  - Set shared mailbox as sender                           │  │
│  │  - Attach files                                           │  │
│  │  - Send or display for review                             │  │
│  │  - Log status back to spreadsheet                         │  │
│  └───────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                     HTML Template File                          │
│  - Clean HTML with {{Placeholder}} syntax                       │
│  - No Word/Office markup                                        │
│  - Inline CSS for email client compatibility                    │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    Outlook Application                          │
│  - Receives MailItem via COM automation                         │
│  - Sends via Exchange using shared mailbox identity             │
└─────────────────────────────────────────────────────────────────┘
```

### Data Flow

1. User triggers macro from Excel
2. Macro prompts for HTML template file path
3. Macro reads HTML template into memory
4. For each row in the spreadsheet:
   a. Skip if Status column shows "Sent"
   b. Clone HTML template string
   c. Replace all `{{Placeholder}}` tokens with cell values
   d. Create Outlook.MailItem object
   e. Set To, CC, BCC, Subject from columns
   f. Set HTMLBody to processed template
   g. Set SentOnBehalfOfName to shared mailbox
   h. Add attachments if paths provided
   i. Send or Display based on user choice
   j. Update Status column with timestamp
5. Display completion summary

## Code Structure

### Main Module: HealthScreeningMailer.bas

```
HealthScreeningMailer.bas
├── SendHealthScreeningInvites()     ' Main entry point
├── LoadHTMLTemplate()               ' Reads HTML file to string
├── ReplacePlaceholders()            ' Substitutes {{tokens}} with values
├── CreateMailItem()                 ' Builds Outlook email
├── SetSharedMailbox()               ' Configures sender identity
├── AddAttachments()                 ' Attaches files from paths
├── GetColumnIndex()                 ' Maps column name to index
├── LogStatus()                      ' Writes send status to sheet
└── ValidateEmailAddress()           ' Basic email format check
```

### Key Functions

#### SendHealthScreeningInvites()

Main orchestration function. Handles:
- User prompts (template selection, send mode)
- Row iteration
- Error handling with continue/abort option
- Progress indication

#### ReplacePlaceholders(template As String, row As Range, headers As Range) As String

Performs token substitution:
- Iterates through all columns
- Converts column names to placeholder format (spaces → underscores)
- Uses case-insensitive replacement
- Returns processed HTML string

#### SetSharedMailbox(mail As Outlook.MailItem, sendAsAddress As String)

Configures shared mailbox sending:
- Sets `SentOnBehalfOfName` property
- Optionally sets `SendUsingAccount` if account column provided
- Handles permission errors gracefully

## Placeholder Syntax

### Format

```
{{Column_Name}}
```

- Double braces for unambiguous parsing
- Underscores replace spaces in column names
- Case-insensitive matching

### Transformation Rules

| Excel Column | Placeholder |
|--------------|-------------|
| `First Name` | `{{First_Name}}` |
| `Screen Type` | `{{Screen_Type}}` |
| `Invitation Code` | `{{Invitation_Code}}` |
| `Company` | `{{Company}}` |

### Implementation

```vba
Function ReplacePlaceholders(template As String, dataRow As Range, headerRow As Range) As String
    Dim result As String
    Dim i As Integer
    Dim colName As String
    Dim placeholder As String
    Dim cellValue As String

    result = template

    For i = 1 To headerRow.Columns.Count
        colName = Trim(headerRow.Cells(1, i).Value)
        If colName <> "" Then
            ' Convert column name to placeholder format
            placeholder = "{{" & Replace(colName, " ", "_") & "}}"
            ' Get cell value, handle empty cells
            cellValue = Trim(CStr(dataRow.Cells(1, i).Value))
            ' Replace all occurrences (case-insensitive)
            result = Replace(result, placeholder, cellValue, Compare:=vbTextCompare)
        End If
    Next i

    ReplacePlaceholders = result
End Function
```

## Outlook COM Integration

### Late Binding vs Early Binding

Use **late binding** to avoid reference issues across different Office versions:

```vba
' Late binding (recommended)
Dim outlookApp As Object
Set outlookApp = CreateObject("Outlook.Application")

' Early binding (requires reference, can break)
Dim outlookApp As Outlook.Application
Set outlookApp = New Outlook.Application
```

### MailItem Properties

| Property | Type | Description |
|----------|------|-------------|
| `.To` | String | Recipient email(s), semicolon-separated |
| `.CC` | String | CC recipient(s) |
| `.BCC` | String | BCC recipient(s) |
| `.Subject` | String | Email subject line |
| `.HTMLBody` | String | HTML content of email body |
| `.BodyFormat` | Integer | 2 = HTML format |
| `.Importance` | Integer | 0=Low, 1=Normal, 2=High |
| `.SentOnBehalfOfName` | String | Shared mailbox address |
| `.SendUsingAccount` | Account | Account object to send from |
| `.Attachments.Add` | Method | Add file attachment |
| `.Send` | Method | Send immediately |
| `.Display` | Method | Open for preview |

### Shared Mailbox Configuration

Two approaches depending on Exchange setup:

```vba
' Option 1: Send on Behalf Of (shows "sent by X on behalf of Y")
mail.SentOnBehalfOfName = "shared@company.com"

' Option 2: Send As (shows only shared mailbox, requires full Send As permission)
Dim acct As Object
For Each acct In outlookApp.Session.Accounts
    If acct.SmtpAddress = "shared@company.com" Then
        mail.SendUsingAccount = acct
        Exit For
    End If
Next
```

## Error Handling Strategy

### Per-Row Error Handling

Errors on individual rows should not abort the entire batch:

```vba
On Error Resume Next
mail.Send
If Err.Number <> 0 Then
    LogStatus row, "Error: " & Err.Description
    Err.Clear
Else
    LogStatus row, "Sent: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
End If
On Error GoTo 0
```

### Critical Errors

These should abort with user notification:
- Outlook not installed/available
- HTML template file not found
- No data rows in spreadsheet
- No To/Subject columns found

### Warning Conditions

Log but continue:
- Empty To field for a row
- Attachment file not found
- Invalid email format
- Empty placeholder values

## Testing Approach

### Manual Testing Checklist

1. **Basic send**: Single recipient, no attachments, preview mode
2. **Batch send**: Multiple recipients, auto-send mode
3. **Attachments**: Valid path, invalid path, multiple attachments
4. **Shared mailbox**: Verify sender shows correctly
5. **Placeholders**: All placeholders replaced, missing columns handled
6. **Status logging**: Sent status written, errors logged
7. **Skip sent**: Re-running skips already-sent rows
8. **Edge cases**: Empty rows, special characters in names, long subjects

### Test Data

Create test rows with:
- Valid internal email addresses
- One row with invalid email format
- One row with missing attachment
- One row with all optional fields empty
- Special characters: O'Brien, Müller, José

## Extension Points

### Adding New Placeholders

No code changes required:
1. Add column to spreadsheet
2. Add `{{Column_Name}}` to HTML template

### Adding New Email Properties

Modify `CreateMailItem()` function:
1. Add column name to header detection
2. Add property assignment in the mail item creation block

Example - adding Read Receipt:
```vba
If GetColumnIndex(headers, "ReadReceipt") > 0 Then
    If UCase(GetCellValue(row, headers, "ReadReceipt")) = "YES" Then
        mail.ReadReceiptRequested = True
    End If
End If
```

### Supporting Multiple Templates

Modify to accept template selection:
```vba
' Add Template column to spreadsheet
templatePath = GetCellValue(row, headers, "Template")
If templatePath = "" Then templatePath = defaultTemplatePath
htmlContent = LoadHTMLTemplate(templatePath)
```

## Common Issues and Solutions

### Issue: "Operation aborted" error

**Cause**: Outlook security prompt blocking automation
**Solution**:
- Use Outlook's Trust Center to allow programmatic access
- Or use Redemption library (if IT permits third-party)

### Issue: Emails appear as plain text

**Cause**: HTMLBody not being applied correctly
**Solution**: Ensure `.BodyFormat = 2` (olFormatHTML) is set before `.HTMLBody`

### Issue: Shared mailbox not found

**Cause**: Account not added to Outlook profile
**Solution**:
- Add shared mailbox via File → Account Settings → Account Settings → Change → More Settings → Advanced
- Or use `SentOnBehalfOfName` instead of `SendUsingAccount`

### Issue: Attachments showing inline

**Cause**: Attachments added after HTMLBody contains CID references
**Solution**: Add attachments before setting HTMLBody, or ensure no `cid:` references in template

### Issue: Special characters corrupted

**Cause**: HTML template encoding mismatch
**Solution**:
- Save HTML template as UTF-8
- Read file with proper encoding:
```vba
Dim stream As Object
Set stream = CreateObject("ADODB.Stream")
stream.Charset = "utf-8"
stream.Open
stream.LoadFromFile templatePath
htmlContent = stream.ReadText
stream.Close
```

## Performance Considerations

### Large Batches

For >100 recipients:
- Add `DoEvents` in loop to prevent Excel freezing
- Consider adding progress bar via UserForm
- Add batch delay to avoid Exchange throttling:
```vba
If rowCount Mod 50 = 0 Then
    Application.Wait Now + TimeValue("00:00:05")
End If
```

### Memory Management

Release objects explicitly:
```vba
Set mail = Nothing
Set outlookApp = Nothing
```

## Security Considerations

### Macro Security

- Code should be signed if organisation requires it
- Document Trust Center settings in README
- Never store credentials in code

### Data Protection

- No PII should be logged externally
- Status column should not contain recipient details
- Consider adding audit trail worksheet

### Email Content

- Validate that placeholders don't allow injection
- Sanitize any user-provided content in cells
- Be cautious with HTML in cell values

## Files in This Project

| File | Description |
|------|-------------|
| `HealthScreeningMailer.bas` | Complete VBA module |
| `invitation_template.html` | Clean HTML template with placeholders |
| `sample_recipient_list.csv` | Example data structure (open in Excel) |
| `README.md` | User documentation |
| `DEVELOPMENT.md` | This file - technical documentation |

## Version Control Notes

### What to Commit

- `.bas` files (VBA exported modules)
- `.html` template files
- `.md` documentation
- Sample `.csv` with dummy data only

### What to Ignore

- Working spreadsheets with real recipient data
- Any files containing PII
- `.xlsm` files (can contain cached data)

## Contact and Handover

This project was developed to replace a Word mail merge workflow that broke after Windows 11 upgrade. The key improvement is eliminating Word from the process entirely, making it more reliable and easier to maintain.

The original macro reference (InviteMacro.txt) is from Imnoss Ltd and provided useful patterns for handling mail merge fields, but the new implementation takes a simpler approach using direct placeholder substitution.
