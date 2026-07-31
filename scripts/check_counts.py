import os
import requests

# Read .env manually
env = {}
env_path = "e:\\ExcellenceSystem\\scripts\\.env"
if os.path.exists(env_path):
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip().strip("'\"")

url = env.get("SUPABASE_URL")
key = env.get("SUPABASE_SERVICE_ROLE_KEY")

headers = {
    "apikey": key,
    "Authorization": f"Bearer {key}",
    "Content-Type": "application/json"
}

print("Fetching counts...")
h_res = requests.get(f"{url}/rest/v1/hierarchy?select=count", headers=headers)
s_res = requests.get(f"{url}/rest/v1/summaries?select=count&period=eq.July 2026", headers=headers)
v_res = requests.get(f"{url}/rest/v1/visits?select=count&period=eq.July 2026", headers=headers)

print("Hierarchy row count:", h_res.json() if h_res.ok else h_res.text)
print("Summaries (July 2026) row count:", s_res.json() if s_res.ok else s_res.text)
print("Visits (July 2026) row count:", v_res.json() if v_res.ok else v_res.text)

print("\nFetching data...")
# Fetch samples
h_data = requests.get(f"{url}/rest/v1/hierarchy?select=employee_code,employee_name&limit=10", headers=headers).json()
s_data = requests.get(f"{url}/rest/v1/summaries?select=employee_code,user_name&period=eq.July 2026&limit=10", headers=headers).json()

print("\nHierarchy codes sample:", [r.get("employee_code") for r in h_data])
print("Summaries codes sample:", [r.get("employee_code") for r in s_data])

# Check intersection
h_all = requests.get(f"{url}/rest/v1/hierarchy?select=employee_code", headers=headers).json()
s_all = requests.get(f"{url}/rest/v1/summaries?select=employee_code&period=eq.July 2026", headers=headers).json()

h_codes = set(r.get("employee_code") for r in h_all if isinstance(r, dict) and r.get("employee_code"))
s_codes = set(r.get("employee_code") for r in s_all if isinstance(r, dict) and r.get("employee_code"))

print(f"\nTotal unique codes in hierarchy: {len(h_codes)}")
print(f"Total unique codes in summaries (July 2026): {len(s_codes)}")
print(f"Intersection of codes: {len(h_codes.intersection(s_codes))}")
