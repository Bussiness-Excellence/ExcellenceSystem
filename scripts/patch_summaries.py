import os
import json
import requests

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
    "Content-Type": "application/json",
    "Prefer": "return=minimal"
}

print("Fetching hierarchy...")
h_res = requests.get(f"{url}/rest/v1/hierarchy?select=employee_code,employee_name", headers=headers)
if not h_res.ok:
    print("Failed to fetch hierarchy", h_res.text)
    exit(1)

hierarchy = h_res.json()
name_to_code = {}
for h in hierarchy:
    if h.get("employee_name") and h.get("employee_code"):
        name = str(h["employee_name"]).strip().lower()
        name_to_code[name] = str(h["employee_code"])

print(f"Loaded {len(name_to_code)} name-to-code mappings from hierarchy.")

print("Fetching summaries with missing employee_code...")
s_res = requests.get(f"{url}/rest/v1/summaries?select=id,user_name&employee_code=is.null", headers=headers)
if not s_res.ok:
    print("Failed to fetch summaries", s_res.text)
    exit(1)

summaries = s_res.json()
print(f"Found {len(summaries)} summaries missing employee_code.")

updates = []
for s in summaries:
    user = s.get("user_name")
    if user:
        name = str(user).strip().lower()
        if name in name_to_code:
            updates.append({
                "id": s["id"],
                "employee_code": name_to_code[name]
            })

print(f"Prepared {len(updates)} updates.")

# Send updates in chunks
chunk_size = 500
success_count = 0
for i in range(0, len(updates), chunk_size):
    chunk = updates[i:i+chunk_size]
    res = requests.post(f"{url}/rest/v1/summaries", headers={
        **headers,
        "Prefer": "resolution=merge-duplicates,return=minimal"
    }, params={"on_conflict": "id"}, json=chunk)
    
    if res.ok:
        success_count += len(chunk)
        print(f"Updated chunk {i} to {i+len(chunk)}")
    else:
        print(f"Error updating chunk: {res.text}")

print(f"Done. Successfully patched {success_count} summaries.")
