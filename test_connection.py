import gspread
import json
import os

def test_connection():
    try:
        with open('sync_config.json', 'r') as f:
            config = json.load(f)
        
        credentials_path = config['credentials_path']
        spreadsheet_id = config['spreadsheet_id']
        
        print(f"Testing connection with:")
        print(f" - Credentials: {credentials_path}")
        print(f" - Spreadsheet ID: {spreadsheet_id}")
        
        gc = gspread.service_account(filename=credentials_path)
        sh = gc.open_by_key(spreadsheet_id)
        
        print(f"\nSUCCESS! Connected to Spreadsheet: '{sh.title}'")
        print("Sheets found:")
        for worksheet in sh.worksheets():
            print(f" - {worksheet.title}")
            
    except Exception as e:
        print(f"\nERROR: {str(e)}")
        if "Project has not enabled" in str(e):
            print("\nTIP: You need to Enable Google Sheets API in GCP Console.")
        if "not found" in str(e).lower():
            print("\nTIP: Make sure you shared the Spreadsheet with the Service Account email.")

if __name__ == "__main__":
    test_connection()
