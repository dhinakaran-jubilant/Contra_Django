from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from django.conf import settings
from datetime import datetime
import gspread
import os

def update_google_sheets(summary_report, final_file_label):
    """
    Update Google Sheets with the summary report data
    Checks for existing entries to avoid duplicates
    Checks last row background color BEFORE adding new rows
    """
    try:
        
        SCOPE = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]

        # Use the path from settings
        key_file_path = r"../safe_storage/robust-shadow-471605-k1-6152c9ae90ff.json"
        key_file_path = settings.GOOGLE_SHEETS_CREDENTIALS_PATH
        
        print(f"Looking for key file at: {key_file_path}")
        print(f"File exists: {os.path.exists(key_file_path)}")
        
        if not os.path.exists(key_file_path):
            print("❌ Google Sheets key file not found")
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
        
        # If there's no data yet (only header), start from row 2
        if len(all_data) <= 1:
            existing_entries = []
            next_row = 2
            total_existing_rows = 0
            # First batch after header should have NO background
            should_have_bg = False
            print("🔍 No existing data rows - first batch will have NO background")
        else:
            # Extract existing file names from column D (index 3)
            existing_entries = [row[3] for row in all_data[1:] if len(row) > 3]  # Skip header row
            next_row = len(all_data) + 1
            total_existing_rows = len(all_data) - 1  # Exclude header
            
            # CHECK LAST ROW BACKGROUND COLOR BEFORE ADDING NEW ROWS
            last_row_num = len(all_data)
            print(f"🔍 Checking background color of last row {last_row_num} BEFORE adding new rows...")
            last_row_has_bg = check_last_row_background(sheet, last_row_num)
            
            # Determine if this batch should have background color
            # If last row has NO background → apply grey background to this batch
            # If last row HAS background → apply NO background to this batch
            should_have_bg = not last_row_has_bg
        
        print(f"Found {len(existing_entries)} existing entries")
        print(f"Total existing data rows: {total_existing_rows}")
        print(f"Last row has background: {last_row_has_bg if 'last_row_has_bg' in locals() else 'N/A'}")
        print(f"🎨 This batch will have {'LIGHT GREY background' if should_have_bg else 'NO background (white)'}")
        
        # Prepare data for each file in summary
        new_rows_added = 0
        skipped_rows = 0
        new_rows_data = []  # Store new rows to add formatting later
        
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
            
            # Store row info for formatting
            new_rows_data.append({
                'row_number': next_row,
                'file_name': file_name,
                'data': new_row_data
            })
            
            print(f"✅ Success: Added data for {file_name} to Google Sheets at row {next_row}.")
            new_rows_added += 1
            next_row += 1

        # Apply batch color to ALL new rows
        if new_rows_added > 0:
            print(f"🎨 Applying {'light grey' if should_have_bg else 'no'} background to {new_rows_added} new rows...")
            apply_batch_color_simple(sheet, new_rows_data, should_have_bg)
        
        print(f"📊 Google Sheets update summary:")
        print(f"   ✅ New rows added: {new_rows_added}")
        print(f"   ⏭️  Rows skipped (already exist): {skipped_rows}")
        print(f"   🎨 Batch color applied: {'Light Grey' if should_have_bg else 'None (White)'}")
        print(f"   📋 Total files processed: {len(summary_report)}")
        
        return True

    except gspread.exceptions.SpreadsheetNotFound:
        print(f"❌ Error: Spreadsheet with title '{SPREADSHEET_TITLE}' not found.")
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
        import traceback
        traceback.print_exc()
        return False


def check_last_row_background(sheet, last_row_num):
    """
    Check the background color of the last row in the sheet
    Returns True if the row has grey background, False if white
    """
    try:
        print(f"🔍 Checking background color for row {last_row_num}...")
        
        # Method 1: Try using gspread's format method
        try:
            # Get the format of the first cell in the last row
            cell_format = sheet.get(f"A{last_row_num}", params={"valueRenderOption": "UNFORMATTED_VALUE", "dateTimeRenderOption": "FORMATTED_STRING"})
            
            # Try to get the background color using cell properties
            # This is a workaround since direct background color access might not be available
            print(f"   Cell format data: {cell_format}")
        except:
            pass
        
        # Method 2: Use Google Sheets API directly with simpler approach
        service = build('sheets', 'v4', credentials=sheet.spreadsheet.client.auth)
        spreadsheet_id = sheet.spreadsheet.id
        
        # Simple approach - get the entire row data
        response = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            ranges=[f"{sheet.title}!A{last_row_num}:I{last_row_num}"],
            includeGridData=True
        ).execute()
        
        print(f"   API Response received: {'sheets' in response}")
        
        if 'sheets' in response and response['sheets']:
            sheet_data = response['sheets'][0]
            if 'data' in sheet_data and sheet_data['data']:
                grid_data = sheet_data['data'][0]
                if 'rowData' in grid_data and grid_data['rowData']:
                    row_data = grid_data['rowData'][0]
                    if 'values' in row_data and row_data['values']:
                        # Check the first cell's background color
                        first_cell = row_data['values'][0]
                        if 'effectiveFormat' in first_cell and 'backgroundColor' in first_cell['effectiveFormat']:
                            bg_color = first_cell['effectiveFormat']['backgroundColor']
                            
                            red = bg_color.get('red', 1.0)
                            green = bg_color.get('green', 1.0)
                            blue = bg_color.get('blue', 1.0)
                            
                            print(f"   🔍 Background color RGB: ({red:.2f}, {green:.2f}, {blue:.2f})")
                            
                            # Check if it's white (no background)
                            is_white = (red == 1.0 and green == 1.0 and blue == 1.0)
                            
                            # Check if it's light grey (our color)
                            is_grey = (0.85 <= red <= 0.95 and 
                                      0.85 <= green <= 0.95 and 
                                      0.85 <= blue <= 0.95)
                            
                            # Check if it's header blue (#99ebeb ≈ 0.6, 0.92, 0.92)
                            is_header_blue = (0.5 <= red <= 0.7 and 
                                            0.85 <= green <= 0.95 and 
                                            0.85 <= blue <= 0.95)
                            
                            print(f"   Is white: {is_white}")
                            print(f"   Is grey: {is_grey}")
                            print(f"   Is header blue: {is_header_blue}")
                            
                            # Return True if row has grey background
                            # Return False for white or header blue
                            if is_grey:
                                print("   ✅ Last row has GREY background")
                                return True
                            else:
                                print("   ⬜ Last row has WHITE or HEADER background")
                                return False
        
        # If we can't determine, use fallback logic based on row number
        print("   ⚠️  Could not determine background color, using row number fallback")
        # Data rows start from row 2
        # Row 2 = index 0 = should be white, Row 3 = index 1 = should be grey, etc.
        data_row_index = last_row_num - 2
        has_bg_fallback = (data_row_index % 2 == 1)  # Odd index = has background
        print(f"   Fallback: row index {data_row_index}, has background: {has_bg_fallback}")
        
        return has_bg_fallback
        
    except Exception as e:
        print(f"   ❌ Error checking background color: {e}")
        # Fallback to row number logic
        data_row_index = last_row_num - 2
        has_bg_fallback = (data_row_index % 2 == 1)
        print(f"   Error fallback: row index {data_row_index}, has background: {has_bg_fallback}")
        return has_bg_fallback


def apply_batch_color_simple(sheet, new_rows_data, should_have_bg):
    """
    Simple and reliable batch color application using gspread formatting
    Applies background color and borders to all cells
    """
    try:
        print(f"🔄 Starting batch color and border application for {len(new_rows_data)} rows...")
        
        for row_info in new_rows_data:
            row_num = row_info['row_number']
            file_name = row_info['file_name']
            
            # Define the range for this row (A to I columns)
            cell_range = f"A{row_num}:I{row_num}"
            
            if should_have_bg:
                # Apply light grey background with borders
                sheet.format(cell_range, {
                    "backgroundColor": {
                        "red": 0.90,
                        "green": 0.90, 
                        "blue": 0.90
                    },
                    "borders": {
                        "top": {"style": "SOLID"},
                        "bottom": {"style": "SOLID"},
                        "left": {"style": "SOLID"},
                        "right": {"style": "SOLID"}
                    }
                })
                print(f"   ✅ Applied light grey background + borders to row {row_num} ({file_name})")
            else:
                # Apply white background with borders
                sheet.format(cell_range, {
                    "backgroundColor": {
                        "red": 1.0,
                        "green": 1.0,
                        "blue": 1.0
                    },
                    "borders": {
                        "top": {"style": "SOLID"},
                        "bottom": {"style": "SOLID"},
                        "left": {"style": "SOLID"},
                        "right": {"style": "SOLID"}
                    }
                })
                print(f"   ✅ Applied white background + borders to row {row_num} ({file_name})")
        
        print(f"🎉 Successfully applied colors and borders to {len(new_rows_data)} rows")
        
    except Exception as e:
        print(f"❌ Error in apply_batch_color_simple: {e}")
        import traceback
        traceback.print_exc()
