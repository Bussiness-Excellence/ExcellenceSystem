import requests
from pulpoplus_config import supabase_config
import urllib.parse

def main():
    url, key = supabase_config()
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal"
    }

    # Count visits by period
    print("Fetching distinct periods in visits table...")
    r = requests.get(
        f"{url}/rest/v1/visits?select=period,visit_date&limit=1000",
        headers=headers
    )
    if r.status_code == 200:
        data = r.json()
        print(f"Sampled {len(data)} visits.")
        periods = {}
        dates = set()
        for row in data:
            p = row.get("period")
            d = row.get("visit_date")
            periods[p] = periods.get(p, 0) + 1
            if d:
                dates.add(d)
        print("Periods distribution:", periods)
        print("A few dates:", sorted(list(dates))[:5])
    else:
        print(f"Error: {r.text}")

if __name__ == "__main__":
    main()
