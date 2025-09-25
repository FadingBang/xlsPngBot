from Image_To_Excel_Spreadsheet import Convert_Google_Sheet_To_Image

# Replace this with your Google Sheets URL
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1OOn0nY3f_SZGn8zDjatO7apu3usObTdhzh42oXOYCKU/edit"

# Convert Google Sheet to image
Convert_Google_Sheet_To_Image(
    spreadsheet_url=GOOGLE_SHEET_URL,
    output_file="output_image.png",
    credentials_file="credentials.json"  # Make sure you have this file from Google Cloud Console
)