import pandas as pd
import json
import sys

def excel_to_json(excel_file, json_file):
    try:
        print(f"Starting conversion process...")
        print(f"Reading Excel file: {excel_file}")
        # Read the Excel file
        df = pd.read_excel(excel_file)
        
        print(f"Excel file read successfully. Shape: {df.shape}")
        print("Converting DataFrame to JSON...")
        # Convert DataFrame to JSON
        json_data = df.to_json(orient='records')
        
        print(f"Saving JSON to file: {json_file}")
        # Save JSON to file
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(json.loads(json_data), f, ensure_ascii=False, indent=2)
        
        print("Conversion completed successfully.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        print(f"Error type: {type(e).__name__}")
        print(f"Error details: {sys.exc_info()}")

if __name__ == "__main__":
    print("Script started.")
    excel_file = 'master_duel_cards.xlsx'  # Your Excel file name
    json_file = 'master_duel_cards.json'   # Output JSON file name
    
    excel_to_json(excel_file, json_file)
    print("Script finished.")