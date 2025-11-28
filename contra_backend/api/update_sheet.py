import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import os
from django.conf import settings

def update_google_sheets(summary_report, final_file_label):
    """
    Update Google Sheets with the summary report data
    """
    try:
        
        SCOPE = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]

        # Use the path from settings
        key_file_path = settings.GOOGLE_SHEETS_CREDENTIALS_PATH
        
        print(f"Looking for key file at: {key_file_path}")
        print(f"File exists: {os.path.exists(key_file_path)}")
        
        if not os.path.exists(key_file_path):
            print("❌ Google Sheets key file not found in static/config/")
            print("Please ensure the file is at: static/config/robust-shadow-471605-k1-6152c9ae90ff.json")
            
            # List files in static/config to help debug
            config_dir = os.path.join(settings.BASE_DIR, 'static', 'config')
            if os.path.exists(config_dir):
                print(f"Files in {config_dir}:")
                for file in os.listdir(config_dir):
                    print(f"  - {file}")
            else:
                print(f"Config directory does not exist: {config_dir}")
                
            return False

        print("✅ Google Sheets key file found!")
        
        SPREADSHEET_TITLE = 'Software Testing Report'
        WORKSHEET_INDEX = 2 

        credentials = Credentials.from_service_account_file(
            key_file_path,
            scopes=SCOPE
        )

        gc = gspread.authorize(credentials)
        spreadsheet = gc.open(SPREADSHEET_TITLE) 
        sheet = spreadsheet.get_worksheet(WORKSHEET_INDEX)
        
        # Get existing entries to determine next serial number
        existing_entries = sheet.col_values(1)
        new_s_no = len(existing_entries)

        # Prepare data for each file in summary
        for item in summary_report:
            # Calculate percentage
            if item["Manual Matched"] > 0:
                percentage = (item["Software Matched"] / item["Manual Matched"]) * 100
            else:
                percentage = 0
            
            new_row_data = [
                new_s_no,
                datetime.now().strftime("%d-%m-%Y"),
                item["Bank Name"],
                item["File Name"],
                item["Total Entries (Manual)"],
                item["Total Entries (Software)"],
                item["Manual Matched"],
                item["Software Matched"],
                f"{percentage:.2f} %"
            ]
            
            sheet.append_row(
                new_row_data, 
                value_input_option='USER_ENTERED'
            )
            
            print(f"✅ Success: Added data for {item['File Name']} to Google Sheets.")
            new_s_no += 1

        print("✅ All data appended to Google Sheets.")
        return True

    except gspread.exceptions.SpreadsheetNotFound:
        print(f"❌ Error: Spreadsheet with title '{SPREADSHEET_TITLE}' not found.")
        print("Please ensure the title is exactly correct and the Service Account has 'Viewer' or 'Editor' access.")
        return False
    except gspread.exceptions.APIError as e:
        print(f"❌ API Error: Check if the Service Account has been shared with the spreadsheet.")
        print(f"Details: {e}")
        return False
    except FileNotFoundError:
        print(f"❌ Error: Service account key file not found.")
        return False
    except Exception as e:
        print(f"❌ An unexpected error occurred: {e}")
        return False
    
