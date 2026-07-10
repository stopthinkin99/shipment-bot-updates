# Uni Creation Shipment Bot

## For the developer (Aayan) — building the installer

### One-time setup on your laptop

1. Make sure you have these downloaded already:
   - Tesseract portable: `C:\Users\aayan.boradia\Downloads\Tesseract-OCR\`
   - Poppler portable:   `C:\Users\aayan.boradia\Downloads\poppler-26.02.0\`
   - Inno Setup: https://jrsoftware.org/isinfo.php (free, install once)

2. Run `build_bundle.bat` — this builds the `bundle\` folder with
   everything self-contained (embedded Python 3.12, all packages,
   Tesseract, Poppler, bot code).

3. Open `installer.iss` in Inno Setup → Build → Compile.
   Output: `dist\ShipmentBotSetup.exe` (~50MB single file)

4. Send `ShipmentBotSetup.exe` to the target PC. Done.

---

## For the user (finance team PC)

1. Double-click `ShipmentBotSetup.exe`
2. Click Next → Install → Finish
3. The app opens automatically
4. Click **Browse** next to "Labels folder" → select the folder
   where shipping label PDFs will be saved
5. Click **Browse** next to "Excel file" → select TRACKING_SHIPMENT.xlsx
6. Enter the alert email address
7. Tick **"Start automatically when Windows starts"**
8. Click **▶ Start Watching**

That's it. The bot runs in the background from now on.

---

## For Aayan — pushing fixes remotely (auto-updater)

1. Create a private GitHub repo: e.g. `unicreation/shipment-bot-updates`
2. In the repo root, create `version.txt` containing: `1.0.0`
3. In the repo, create a `bot/` folder and put all `.py` files there
4. In `updater.py` fill in:
   - `GITHUB_OWNER` = your GitHub username
   - `GITHUB_REPO`  = repo name
   - `GITHUB_TOKEN` = a read-only Personal Access Token
     (GitHub → Settings → Developer settings → Personal access tokens
      → Classic → scopes: repo read only)

When you fix a bug:
1. Edit the `.py` file(s) in the `bot/` folder of the GitHub repo
2. Bump `version.txt` (e.g. `1.0.0` → `1.0.1`)
3. Commit and push
4. Next time the bot starts on the target PC, it downloads
   the new files silently and restarts itself

The user sees nothing except a brief "Updating…" message in the log.

---

## File structure inside the bundle

```
ShipmentBot\
├── ShipmentBot.exe       <- launcher
├── version.txt           <- current version (e.g. 1.0.0)
├── config.json           <- saved paths (written by app on first run)
├── app.py
├── updater.py
├── processor.py
├── parser.py
├── extractor.py
├── excel_writer.py
├── mailer.py
├── python\               <- embedded Python 3.12 + all packages
├── Tesseract-OCR\        <- portable Tesseract
└── poppler\bin\          <- portable Poppler
```
