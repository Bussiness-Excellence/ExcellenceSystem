import os
from supabase import create_client

url = "https://xxbfwvlqixnmonxytdxq.supabase.co"
key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inh4YmZ3dmxxaXhubW9ueHl0ZHhxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI3NTYxNjUsImV4cCI6MjA5ODMzMjE2NX0.7CXnsUbOjC3JDFcZoQBr1TKbtt31GLbu_0bIrmNK8wk"

sb = create_client(url, key)

v_periods = sb.from("visits").select("period").limit(100).execute()
s_periods = sb.from("summaries").select("period").limit(100).execute()

v_unique = set(r["period"] for r in (v_periods.data or []) if "period" in r)
s_unique = set(r["period"] for r in (s_periods.data or []) if "period" in r)

print("Visits periods sample:", v_unique)
print("Summaries periods sample:", s_unique)
