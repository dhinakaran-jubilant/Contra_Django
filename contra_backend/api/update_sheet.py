import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import os
from django.conf import settings

def update_google_sheets(summary_report, final_file_label):
    """
    Update Google Sheets with the summary report data
    Checks for existing entries to avoid duplicates
    """
    try:
        
        SCOPE = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]

        # Use the path from settings
        key_file_path = r"../safe_storage/robust-shadow-471605-k1-6152c9ae90ff.json"
        
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
        
        # Get all existing data to check for duplicates
        print("📋 Checking for existing entries...")
        all_data = sheet.get_all_values()
        
        # If there's no data yet, start from row 2 (skip header)
        if len(all_data) <= 1:
            existing_entries = []
            next_row = 2
        else:
            # Extract existing file names from column D (index 3)
            existing_entries = [row[3] for row in all_data[1:] if len(row) > 3]  # Skip header row
            next_row = len(all_data) + 1
        
        print(f"Found {len(existing_entries)} existing entries")
        
        # Prepare data for each file in summary
        new_rows_added = 0
        skipped_rows = 0
        
        for item in summary_report:
            file_name = item["File Name"]
            
            # Check if this file already exists in the sheet
            if file_name in existing_entries:
                print(f"⏭️  Skipping {file_name} - already exists in Google Sheets")
                skipped_rows += 1
                continue
            
            # Calculate percentage
            if item["Manual Matched"] > 0:
                percentage = (item["Software Matched"] / item["Manual Matched"]) * 100
            else:
                percentage = 0
            
            new_row_data = [
                next_row - 1,  # Serial number (adjusting for header row)
                datetime.now().strftime("%d-%m-%Y"),
                item["Bank Name"],
                file_name,
                item["Total Entries (Manual)"],
                item["Total Entries (Software)"],
                item["Manual Matched"],
                item["Software Matched"],
                f"{percentage:.2f} %"
            ]
            
            # Append the new row
            sheet.append_row(
                new_row_data, 
                value_input_option='USER_ENTERED'
            )
            
            print(f"✅ Success: Added data for {file_name} to Google Sheets.")
            new_rows_added += 1
            next_row += 1

        print(f"📊 Google Sheets update summary:")
        print(f"   ✅ New rows added: {new_rows_added}")
        print(f"   ⏭️  Rows skipped (already exist): {skipped_rows}")
        print(f"   📋 Total files processed: {len(summary_report)}")
        
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