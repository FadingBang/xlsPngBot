# Google Sheets Setup Guide

1. Go to the Google Cloud Console (https://console.cloud.google.com/)
2. Create a new project
3. Enable the Google Sheets API for your project
4. Create credentials (Service Account Key):
   - Go to APIs & Services > Credentials
   - Click "Create Credentials" > "Service Account"
   - Fill in the service account details
   - Click "Create and Continue"
   - Skip role selection (or select Basic > Viewer)
   - Click "Done"
5. Once created, click on the service account
6. Go to the "Keys" tab
7. Add Key > Create New Key > JSON
8. Save the downloaded JSON file as "credentials.json" in your project directory
9. Share your Google Sheet with the email address from the service account (found in credentials.json)

Important: Never commit credentials.json to your repository. Add it to .gitignore!