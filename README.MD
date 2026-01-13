# Health Screening Invitation Mailer

A VBA-based tool for sending personalised health screening invitations via Outlook from a shared mailbox. Designed for environments with IT restrictions where third-party software installation is not permitted.

## Overview

This tool allows you to:

- Send bulk HTML email invitations from a shared Outlook inbox
- Personalise each email with recipient details (name, company, invitation code, etc.)
- Include attachments
- Log sent status directly in your Excel spreadsheet
- Preview emails before sending or send automatically

## Requirements

- Microsoft Excel (2016 or later)
- Microsoft Outlook (2016 or later)
- "Send As" or "Send on Behalf Of" permissions on the shared mailbox
- Macros enabled in Excel

## Files Included

| File | Description |
|------|-------------|
| `HealthScreeningMailer.bas` | VBA macro to import into Excel |
| `invitation_template.html` | HTML email template with placeholders |
| `sample_recipient_list.xlsx` | Example spreadsheet with correct column structure |

## Setup

### 1. Enable Macros in Excel

1. Open Excel
2. Go to **File → Options → Trust Center → Trust Center Settings**
3. Select **Macro Settings**
4. Choose "Disable VBA macros with notification" (recommended) or "Enable VBA macros"
5. Click OK

### 2. Import the Macro

1. Open your recipient list spreadsheet in Excel
2. Press **Alt + F11** to open the VBA Editor
3. Go to **File → Import File**
4. Select `HealthScreeningMailer.bas`
5. Close the VBA Editor

### 3. Save Your Workbook

Save your workbook as a macro-enabled file:
- **File → Save As**
- Change file type to **Excel Macro-Enabled Workbook (.xlsm)**

### 4. Place the HTML Template

Save `invitation_template.html` to a known location (e.g., `C:\EmailTemplates\`). You'll need this path when running the macro.

## Excel Column Structure

Your spreadsheet should include these columns (column names must match exactly):

### Required Columns

| Column | Description | Example |
|--------|-------------|---------|
| `To` | Recipient email address | john.smith@company.com |
| `Subject` | Email subject line | Schedule Your Health Assessment |

### Optional Email Columns

| Column | Description | Example |
|--------|-------------|---------|
| `CC` | CC recipient(s) | manager@company.com |
| `BCC` | BCC recipient(s) | records@yourcompany.com |
| `Account` | Your Outlook account | you@yourcompany.com |
| `SendAs` | Shared mailbox address | shared@yourcompany.com |
| `Importance` | High, Normal, or Low | High |
| `Attachment1` | Full path to attachment | C:\Files\Brochure.pdf |
| `Attachment2` | Full path to second attachment | C:\Files\Form.pdf |

### Personalisation Columns

| Column | Description | Template Placeholder |
|--------|-------------|---------------------|
| `Title` | Mr, Mrs, Ms, Dr, etc. | `{{Title}}` |
| `First Name` | Recipient's first name | `{{First_Name}}` |
| `Surname` | Recipient's surname | `{{Surname}}` |
| `Company` | Recipient's company | `{{Company}}` |
| `Screen Type` | Type of health screening | `{{Screen_Type}}` |
| `Invitation Code` | Unique booking code | `{{Invitation_Code}}` |

### Status Column

| Column | Description |
|--------|-------------|
| `Status` | Automatically populated with "Sent" and timestamp |

## HTML Template Placeholders

The HTML template uses double-brace placeholders that get replaced with values from your spreadsheet:

```html
<p>Dear {{Title}} {{Surname}},</p>

<p>Your company, {{Company}}, has arranged for you to receive a {{Screen_Type}}.</p>

<p>Your invitation code is: <strong>{{Invitation_Code}}</strong></p>
```

You can add additional placeholders by:
1. Adding a column to your spreadsheet
2. Adding `{{Column_Name}}` to your HTML template (replace spaces with underscores)

## Usage

### Running the Macro

1. Open your recipient list spreadsheet
2. Press **Alt + F8**
3. Select `SendHealthScreeningInvites`
4. Click **Run**

### Options

When the macro runs, you'll be prompted to:

1. **Select the HTML template file** — Browse to your `invitation_template.html`
2. **Choose send mode:**
   - **Preview** — Opens each email for review before sending
   - **Send All** — Sends all emails automatically
   - **Cancel** — Abort the operation

### Monitoring Progress

- The macro processes one row at a time
- The `Status` column updates after each successful send
- If an error occurs, the status will show the error message
- Already-sent rows (with "Sent" in Status column) are skipped

## Troubleshooting

### "Macros have been disabled"

- Check Trust Center settings (see Setup step 1)
- Ensure the file is saved as `.xlsm`
- Check if your IT policy blocks macros via Group Policy

### "Outlook cannot send from this account"

- Verify you have Send As permissions on the shared mailbox
- Contact your IT administrator to grant permissions
- Check the `SendAs` column has the correct email address

### Emails stuck in Outbox

- Ensure Outlook is open and connected
- Check your network connection
- Verify the shared mailbox is accessible

### Attachments not found

- Use full file paths (e.g., `C:\Files\document.pdf`)
- Verify the file exists at the specified location
- Check for typos in the path

### HTML formatting looks wrong

- Avoid editing the HTML template in Word (it adds extra markup)
- Use a plain text editor like Notepad
- Test with a single recipient first using Preview mode

## Security Notes

- This macro only uses built-in Microsoft Office features
- No data is sent to external services
- Emails are sent through your organisation's Exchange server
- All logging stays within your local spreadsheet

## Customisation

### Adding New Placeholders

1. Add a new column to your Excel spreadsheet
2. In the HTML template, add `{{Column_Name}}` where you want the value to appear
3. Replace spaces in column names with underscores in the placeholder

### Modifying the Template

The HTML template can be fully customised. Ensure you:

- Keep the placeholder syntax: `{{Placeholder_Name}}`
- Test changes with Preview mode first
- Use inline CSS for best email client compatibility

## Support

For issues or questions:

1. Check the Troubleshooting section above
2. Review the sample files for correct formatting
3. Test with a small batch using Preview mode

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | January 2026 | Initial release |

## Credits

Developed to replace Word mail merge workflow for health screening invitations in restricted IT environments.
