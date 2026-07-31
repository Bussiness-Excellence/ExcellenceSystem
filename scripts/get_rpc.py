import os
import json
import psycopg2
from dotenv import dotenv_values

env = dotenv_values("e:\\ExcellenceSystem\\scripts\\.env")
url = env.get("SUPABASE_URL")
key = env.get("SUPABASE_SERVICE_ROLE_KEY")

# Extract connection string parameters from URL or use direct postgres connection if available,
# or we can query it via PostgREST by executing an ad-hoc query!
# Wait, can we execute custom SQL via PostgREST? No, RLS prevents it unless we call a function.
# But wait! We can connect directly to PostgreSQL database since the host and credentials can be inferred!
# The connection string for Supabase is:
# postgresql://postgres:[password]@db.[ref].supabase.co:5432/postgres
# Do we know the password? No, it's not in .env.
# But wait! Can we fetch the function definition via a simple HTTP GET using PostgREST?
# Yes! We can query `pg_proc` via the postgrest API if RLS allows or if we use the service_role key!
# PostgREST exposes all views and tables in the `public` schema. If pg_proc is not exposed, we might not see it.
# Let's try to query pg_proc or information_schema via a PostgREST query!

import requests
headers = {
    "apikey": key,
    "Authorization": f"Bearer {key}",
    "Content-Type": "application/json"
}

# Let's see if we can call a system catalog or check available RPCs
res = requests.get(f"{url}/rest/v1/", headers=headers)
if res.ok:
    swagger = res.json()
    paths = swagger.get("paths", {})
    print("Available REST paths:")
    for p in paths:
        if "rpc" in p:
            print(f"  {p}")
            
# Let's run a test query to get the function definition if pg_proc is exposed
# Or we can query the function metadata
res_func = requests.get(f"{url}/rest/v1/rpc/get_dashboard_data", headers=headers)
print(f"RPC get_dashboard_data info: {res_func.status_code}")
