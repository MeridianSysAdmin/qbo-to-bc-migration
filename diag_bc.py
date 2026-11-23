"""Diagnose why the BC API returns no companies. Prints no secrets.

Usage (same env vars as reconcile.py --api): python diag_bc.py
"""
import base64
import json
import os

import msal
import requests

tenant = os.environ["BC_TENANT_ID"]
app = msal.ConfidentialClientApplication(os.environ["BC_CLIENT_ID"],
                                         authority=f"https://login.microsoftonline.com/{tenant}",
                                         client_credential=os.environ["BC_CLIENT_SECRET"])
tok = app.acquire_token_for_client(scopes=["https://api.businesscentral.dynamics.com/.default"])
if "access_token" not in tok:
    raise SystemExit(f"No token: {tok.get('error')}: {tok.get('error_description')}")

payload = tok["access_token"].split(".")[1]
claims = json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4)))
print("Token tenant :", claims.get("tid"))
print("Token app    :", claims.get("appid"))
print("Token roles  :", claims.get("roles"))   # should include API.ReadWrite.All
print()

h = {"Authorization": f"Bearer {tok['access_token']}", "Accept": "application/json"}
for env in [os.environ.get("BC_ENVIRONMENT", "Production"), "Production", "Sandbox"]:
    url = f"https://api.businesscentral.dynamics.com/v2.0/{tenant}/{env}/api/v2.0/companies"
    r = requests.get(url, headers=h)
    print(f"[{env}] HTTP {r.status_code}")
    print("   ", r.text[:400].replace("\n", " "))
