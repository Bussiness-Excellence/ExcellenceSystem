import pandas as pd
import requests
import json
import sys

excel_path = r"C:\Users\Administrator\Downloads\hierarchy_export_FIXED.xlsx"

try:
    df = pd.read_excel(excel_path)
    print("Excel columns:", df.columns.tolist())
    print("First 3 rows:")
    print(df.head(3).to_dict('records'))
except Exception as e:
    print(f"Error reading Excel: {e}")

supabase_url = "https://xxbfwvlqixnmonxytdxq.supabase.co"
anon_key = "sb_publishable_EmW0iOc7qMdpEtGdKbjRgQ_-ETHGD2m"

headers = {
    "apikey": anon_key,
    "Authorization": f"Bearer {anon_key}",
}

print("\nFetching current hierarchy table from Supabase...")
res = requests.get(f"{supabase_url}/rest/v1/hierarchy?select=*&limit=3", headers=headers)
print("Status:", res.status_code)
if res.status_code == 200:
    print(json.dumps(res.json(), indent=2))
else:
    print(res.text)

print("\nFetching teams table from Supabase...")
res_teams = requests.get(f"{supabase_url}/rest/v1/teams?select=*", headers=headers)
print("Status:", res_teams.status_code)
if res_teams.status_code == 200:
    print(json.dumps(res_teams.json(), indent=2))
else:
    print(res_teams.text)

