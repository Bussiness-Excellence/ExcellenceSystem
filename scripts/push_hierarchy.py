import pandas as pd
import requests
import numpy as np

# --- Configuration ---
excel_path = r"C:\Users\Administrator\Downloads\hierarchy_export_FIXED.xlsx"
supabase_url = "https://xxbfwvlqixnmonxytdxq.supabase.co"
anon_key = "sb_publishable_EmW0iOc7qMdpEtGdKbjRgQ_-ETHGD2m"

headers = {
    "apikey": anon_key,
    "Authorization": f"Bearer {anon_key}",
    "Content-Type": "application/json",
    "Prefer": "return=minimal"
}

def main():
    print("Loading Excel file...")
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    print(f"Loaded {len(df)} rows. Columns found: {', '.join(df.columns)}")
    
    # Replace NaN/NaT with None for proper JSON serialization
    df = df.replace({np.nan: None})
    records = df.to_dict('records')

    print("\nDeleting existing hierarchy data...")
    # Using a dummy filter like id=not.is.null or employee_code=not.is.null to delete all rows
    # Alternatively, you could just truncate the table from the Supabase dashboard
    delete_res = requests.delete(
        f"{supabase_url}/rest/v1/hierarchy?employee_code=not.is.null", 
        headers=headers
    )
    
    if delete_res.status_code >= 400:
        print(f"Failed to delete existing records (Status: {delete_res.status_code}):", delete_res.text)
        print("Note: If you get a 401 Unauthorized or 403 Forbidden, you might need to use the Supabase Service Role Key instead of the Anon Key, or manually empty the table in the Supabase Dashboard before inserting.")
    else:
        print("Successfully cleared existing hierarchy data.")

    print(f"\nPushing {len(records)} records to Supabase...")
    batch_size = 500
    for i in range(0, len(records), batch_size):
        batch = records[i:i + batch_size]
        res = requests.post(
            f"{supabase_url}/rest/v1/hierarchy",
            headers=headers,
            json=batch
        )
        if res.status_code >= 400:
            print(f"Failed to insert batch {i} to {i+len(batch)}:")
            print(res.text)
        else:
            print(f"Successfully inserted batch {i} to {i+len(batch)}")

    print("\nUpload complete!")

if __name__ == "__main__":
    main()
