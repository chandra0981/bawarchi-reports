# Bawarchi Google Drive to GoDaddy Production Setup

Flow:

Google Drive Paytype folder -> GitHub Actions scheduled job -> data/paytype-data.json -> GoDaddy hosting.

Users upload daily Paytype Excel files only to Google Drive.

Required GitHub Secrets:
- GOOGLE_SERVICE_ACCOUNT_JSON
- DRIVE_PAYTYPE_FOLDER_ID
- GODADDY_FTP_SERVER
- GODADDY_FTP_USERNAME
- GODADDY_FTP_PASSWORD
- GODADDY_FTP_REPORTS_DIR
