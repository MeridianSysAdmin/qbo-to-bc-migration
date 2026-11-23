"""Show how many records an entity has and the fields on the first one.

Usage: python peek.py Vendor
"""
import json
import sys

entity = sys.argv[1]
with open(f"data/{entity}.json", encoding="utf-8") as f:
    rows = json.load(f).get("QueryResponse", {}).get(entity, [])

print(f"{entity}: {len(rows)} records\n")
print(json.dumps(rows[0], indent=2) if rows else "No records")
