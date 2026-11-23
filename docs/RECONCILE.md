# Reconcile: one-time API setup (about 15 minutes)

`reconcile.py` reads your BC trial through Business Central's API. To do that it needs an **app registration** in Microsoft Entra ID (the same Azure AD you work with for SSO at your job), plus permission inside BC.

## 1. Register the app (Azure portal)

1. Go to https://portal.azure.com. Sign in with the **same account your BC trial uses**.
2. Search **App registrations** → **New registration**. Name it `qbo-bc-reconcile`, keep "Single tenant", and click **Register**.
3. On the overview page, copy the **Application (client) ID** and the **Directory (tenant) ID**.
4. **API permissions** → **Add a permission** → **Dynamics 365 Business Central** → **Application permissions** → tick **API.ReadWrite.All** → Add. Then **Grant admin consent** and choose Yes.
5. **Certificates & secrets** → **New client secret** → copy the **Value** right away. It's only shown once.

## 2. Allow the app inside Business Central

1. In BC, Alt+Q → **Microsoft Entra Applications** → **New**.
2. **Client ID** = the Application (client) ID from step 1.3. Description: `qbo-bc-reconcile`. Set **State** = **Enabled**.
3. In **User Permission Sets**, add **D365 READ**, with **Company** blank so it applies to all companies.
4. Click **Grant Consent** if BC asks.

## 3. Run it

Run this in PowerShell from this folder, filling in your own values. Keep the secret out of chat and out of git.

```powershell
$env:BC_TENANT_ID="..."; $env:BC_CLIENT_ID="..."; $env:BC_CLIENT_SECRET="..."; $env:BC_ENVIRONMENT="Production"; $env:BC_COMPANY="Craigs Landscaping MIG"
pip install msal
python reconcile.py
python reconcile.py --break
```

`BC_ENVIRONMENT` is the environment name shown in your BC URL (`businesscentral.dynamics.com/<tenant>/<environment>`), usually `Production` on a trial.

**What good looks like:** `python reconcile.py` prints all PASS lines. `--break` prints at least one FAIL and ends with "Break test WORKED".
