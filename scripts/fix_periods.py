import requests
from pulpoplus_config import supabase_config

def main():
    url, key = supabase_config()
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal"
    }

    # Update visits where period is null
    print("Updating visits with NULL period to 'July 2026'...")
    r = requests.patch(
        f"{url}/rest/v1/visits?period=is.null",
        headers=headers,
        json={"period": "July 2026"}
    )
    if r.status_code in (200, 204):
        print("Success!")
    else:
        print(f"Error: {r.text}")

if __name__ == "__main__":
    main()
