# Tele Space Form Deployment

## Google Apps Script

1. Open the Apps Script project used by the form.
2. Replace its existing code with the contents of `Code.gs`.
3. Run `authorizeServices` once and approve Spreadsheet, Drive and Mail permissions.
4. Deploy a new Web App version:
   - Execute as: Me
   - Who has access: Anyone
5. If Google issues a different `/exec` URL, update `SCRIPT_URL` in `index.html`.

The script automatically creates a worksheet named `Service Requests` with all required columns. Existing worksheets are not modified.

If an earlier test version already created `Service Requests`, create a fresh worksheet or change its final two headers to `Pre Issues` and `Post Issues` before deploying this version.

## Frontend

Deploy the single `index.html` file.

Each submission:

- creates one new Sheet row and Job No;
- stores Pre and Post function issues separately;
- treats unmarked functions as having no recorded problem;
- stores device condition and notes;
- saves the signature image to Google Drive;
- sends a confirmation email to the customer.

Any later changes can be made directly in Google Sheet.
