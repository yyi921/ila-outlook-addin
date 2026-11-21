# iLA for Outlook (test add-in)

A task pane that opens beside an email in Outlook. It shows the sender, subject, received date, attachments and a body preview, and has a "Copy summary" button. "Send to iLA" is a placeholder for the next version.

```
docs/      the web files Outlook loads (host these over HTTPS)
package/   manifest.json + color.png + outline.png
build.py   regenerates the manifest for a base URL and zips it to dist/
```

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
- If you change only `taskpane.html`, `.js` or `.css`, you just redeploy those files. You don't need to re-upload the package.
- Replace `make_icons.py`'s placeholder icons with real brand icons before you publish.
- Publishing to the Marketplace requires live privacy and terms pages at the URLs in `build.py`.
