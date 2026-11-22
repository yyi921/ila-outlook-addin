# iLA for Outlook (test add-in)

A task pane that opens beside an email in Outlook. For every DDTA PDF attached to the email it builds an AP purchase invoice page, puts it in front of the DDTA scan, and emails each combined PDF on its own to Lightyear.

```
docs/              the web files Outlook loads (host these over HTTPS)
  config.js        Lightyear address and the Entra client ID
  ddta.js          reads the DDTA file name, draws the invoice, joins the PDFs
  taskpane.*       the pane: review, download, send
  vendor/          pdf-lib 1.17.1 and @azure/msal-browser 3.30.0
package/           manifest.json + color.png + outline.png
build.py           regenerates the manifest for a base URL and zips it to dist/
data/              sample DDTAs and invoice templates (git-ignored: real patron bank details)
```

## How the DDTA function works

1. Attachments named like `DDTA 1234 - 260105 CITIZEN Jane $2,500 DDTA#1234.pdf` are picked up. The DDTA number, date (YYMMDD), patron and amount come from the file name, because the scans have no readable text. Each value can be corrected in the pane before sending.
2. The invoice page copies the Sage Intacct print in `data/DDTA invoice template.pdf`: supplier `SV0753--Gaming wins`, invoice and reference number `DDTA1234`, description `DDTA1234 - Jane Citizen`, the amount on `1054--Cash on Hand - Main Bank/Safe`, and $0.00 on `1080--GST Input Tax`, location `10--Breakwater Island Trust`. All three dates are the DDTA date. These constants are `DDTA_PROFILE` in `docs/ddta.js`.
3. **Send to Lightyear** sends one email per ticked DDTA from the signed-in user's mailbox to the address in `config.js`. The subject is the description and the attachment is `DDTA1234 - Jane Citizen.pdf`. A copy lands in Sent Items.
4. While **Test: send to me, not Lightyear** is ticked, the email goes to the signed-in user's own inbox with `[TEST]` in the subject. The box starts ticked while `testMode` is `true` in `config.js`; set it to `false` once testing is done.
5. A DDTA already sent to Lightyear from this browser is unticked and marked "Already sent", so a second click does not create a duplicate bill.

To check the output without Outlook, serve `docs/` locally (`python -m http.server` inside `docs/`), open `taskpane.html` and pick DDTA PDFs. That mode only downloads.

## Let the add-in send email (one-off, about 5 minutes)

Until this is done the pane builds and downloads PDFs but the Send button stays off.

1. entra.microsoft.com → App registrations → New registration. Name: `iLA Outlook`. Supported accounts: *Accounts in any organizational directory* (so users in other tenants can use it).
2. Redirect URI: platform *Single-page application*, value `brk-multihub://yyi921.github.io` (the host only; use `brk-multihub://ila.works` after the move).
3. API permissions → Add → Microsoft Graph → Delegated → `Mail.Send`.
4. Copy the *Application (client) ID* into `clientId` in `docs/config.js` and redeploy `docs/`.

Each user is asked once to let iLA send mail on their behalf. Sign-in uses nested app authentication, which works in Outlook on the web, new Outlook and current classic Outlook.

## 1. Host the web files (free, about 5 minutes)

1. Create a public GitHub repo called `ila-outlook-addin` and push this folder to it.
2. In the repo, go to Settings → Pages → Source: *Deploy from a branch* → `main` / `/docs` → Save.
3. Check that https://yyi921.github.io/ila-outlook-addin/taskpane.html loads.

## 2. Build the package

```bash
python build.py https://yyi921.github.io/ila-outlook-addin
```
This creates `dist/ila-outlook.zip`, which is the file you upload.

## 3. Install it in your tenant (sign in as ErinYin@AFISCONSULTANCYPTYLTD.onmicrosoft.com)

**Option A: Teams (quickest)**
Teams → Apps → Manage your apps → Upload an app → *Upload a custom app* → choose `ila-outlook.zip` → Add.
If the upload option is missing, turn it on in the Teams admin center: Teams apps → Setup policies → Global → *Upload custom apps* = On. Allow a few hours for the change to apply.

**Option B: Microsoft 365 admin center (deploy to users)**
admin.microsoft.com → Settings → Integrated apps → *Upload custom apps* → upload the zip → assign it to yourself.

## 4. Try it

Open an email in Outlook on the web (outlook.office.com) or the new Outlook. The **Open iLA** button appears on the ribbon; if you don't see it, check the **Apps** button or the "…" menu. It can take a few minutes to appear, and up to 24 hours in classic Outlook for Windows.

## Moving to ila.works/ila2 later

1. Serve the contents of `docs/` at `https://ila.works/ila2/` (as Django static files, or one view per file).
2. Run `python build.py https://ila.works/ila2`.
3. Increase `VERSION` in `build.py` (for example to 1.0.1), rebuild, and upload the new zip.
4. Keep `APP_ID` unchanged so Outlook treats it as the same app.

## Notes
- If you change only files in `docs/`, you just redeploy those files. You don't need to re-upload the package.
- PDFs over 3 MB are not sent automatically; download them from the pane and email them by hand.
- Replace `make_icons.py`'s placeholder icons with real brand icons before you publish.
- Publishing to the Marketplace requires live privacy and terms pages at the URLs in `build.py`.
