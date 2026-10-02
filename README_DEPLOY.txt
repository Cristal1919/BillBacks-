MMD BILLBACK VOUCHER — DEPLOY-READY (web + Microsoft login)
===========================================================
This is the standalone voucher tool, ready to host as a private website for
the MMD team. Your existing DA app is NOT part of this and is NOT modified.

WHAT'S HERE
  voucher_tool_web.py        the web app (Microsoft login + shared storage)
  voucher.py, engine.py,     your existing classifier (unchanged), imported
    da_audit.py, brand_hierarchy.py, brands.py
  voucher_template.xlsx      the MMD voucher form
  requirements.txt           python dependencies (incl. Authlib for login)
  render.yaml                one-click config for Render.com (host + disk)
  .streamlit/secrets.toml.EXAMPLE   auth config template (fill from IT, keep private)
  .gitignore                 keeps secrets + data out of git

COST (annual, all-in): ~$100-150
  - Render Starter instance (always-on, private)        ~$84/yr
  - 1 GB persistent disk (shared learned mappings)      ~$3/yr
  - Microsoft 365 login (uses MMD's existing Entra ID)  $0
  - Subdomain billbacks.mmdusa.com (MMD already owns domain)  $0

------------------------------------------------------------------
DEPLOY STEPS (in order)
------------------------------------------------------------------
1. PUT CODE ON GITHUB
   - Push these files to a private repo (e.g. your existing Cristal1919 account).
   - .gitignore already excludes secrets and data files.

2. GET 3 VALUES FROM MMD IT (Entra ID app registration)
   Ask IT to register an app and return:
     - Application (client) ID
     - A client secret VALUE
     - Directory (tenant) ID
   Give IT this redirect URL to whitelist:
     https://billbacks.mmdusa.com/oauth2callback
   (IT also adds one DNS CNAME so billbacks.mmdusa.com points to Render.)

3. DEPLOY ON RENDER
   - New > Web Service > connect the GitHub repo. render.yaml is auto-detected.
   - It provisions the instance + the 1 GB disk at /var/data automatically.

4. ADD THE SECRETS (not in git)
   - In Render > the service > Environment, add a Secret File at
     path  .streamlit/secrets.toml  with the contents of
     secrets.toml.EXAMPLE, filled in:
        client_id              = (from IT)
        client_secret          = (from IT)
        server_metadata_url    = https://login.microsoftonline.com/<TENANT_ID>/v2.0/.well-known/openid-configuration
        redirect_uri           = https://billbacks.mmdusa.com/oauth2callback
        cookie_secret          = any long random string you generate
   - Restart the service.

5. POINT THE DOMAIN
   - Render > Settings > Custom Domain > add billbacks.mmdusa.com.
   - IT adds the CNAME Render shows. HTTPS is automatic.

6. TEST
   - Visit the domain -> "Sign in with Microsoft" -> log in with an
     @mmdusa.com account -> you're in. Non-MMD accounts are refused.

------------------------------------------------------------------
NOTES
  - Local testing: running without secrets.toml runs the app OPEN (no login),
    so you can test logic on your PC. Login only activates once [auth] exists.
  - Learned mappings are SHARED and persistent (stored on /var/data), so a
    code a teammate assigns applies for everyone and survives redeploys.
  - The two baked-in fixes (Clos Jordanne->AW, BRUT 18W->LRV) live in this
    tool only; your app files are untouched.
  - This hosts distributor pricing data on Render (a third-party, behind MMD
    login). Confirm that's acceptable with whoever owns data policy at MMD.
