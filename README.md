# Email Filter App

The app automatically reads file from email-blacklist.xlsx, and export the file to Gmail.

## Add Credentials
- Create the local credentials folder.
- Add your email account in a specific personal folder.
- modify the code to include the account.

## Access Google Gmail API
- Enable the API: https://console.cloud.google.com/flows/enableapi?apiid=gmail.googleapis.com
- Go to Credentials: https://console.cloud.google.com/apis/credentials
- Install Google client library:
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
- For details, visit: https://developers.google.com/gmail/api/quickstart/python