# Dronetiks booth questionnaire

A lead-capture form for the convention booth that works on phones and laptops. Booth staff search for an attendee by name or company, and the attendee's details fill in from the organiser's attendee list. Staff complete the rest of the questionnaire, and each submission becomes a row in a Google Sheet.

It runs on **Google Sheets + Google Apps Script**, which is free and needs no server or database. You only need a Google account.

```
Google Sheet
 ├─ attendee tab  ← the organiser's Excel file, imported as-is
 └─ Responses     ← created by the app, one row per lead
Apps Script web app (apps-script/) → one link, open it on any device
```

## Files

| File | What it is |
| --- | --- |
| `apps-script/Code.gs` | Server code and **the questionnaire definition** (`FORM`). Edit questions here. |
| `apps-script/Index.html` | The page: attendee search, the form, and saving entries on the device when offline. |
| `apps-script/Logo.html` | Optional logo (embedded image). If you skip it, the page has no logo. |

## One-time setup (about 10 minutes)

1. **Create the sheet.** Go to [sheets.new](https://sheets.new) and name the sheet (e.g. "Edge 2026 Leads").
   Then use **File → Import → Upload**, choose the attendee `.xlsx`, set *Import location* to **Insert new sheet(s)**, and click **Import data**.
   You don't need to rename or tidy the tab. The app finds the tab that has first/last-name columns and reads the standard headers (`AccessCode`, `AttendeeFirstName`, `AttendeeLastName`, `AttendeeJobTitle`, `AttendeeCompany`, `Zip`, `Country`, `Added/Cancelled`).
2. **Add the code.** In the sheet, open **Extensions → Apps Script**.
   - Replace everything in `Code.gs` with the contents of [`apps-script/Code.gs`](apps-script/Code.gs).
   - Click **+ → HTML** and name the file `Index`. Paste [`apps-script/Index.html`](apps-script/Index.html).
   - Optional: click **+ → HTML** again, name it `Logo`, and paste [`apps-script/Logo.html`](apps-script/Logo.html).
   - Save (Ctrl/Cmd + S).
3. **Set a booth PIN (recommended).** Go to **Project Settings** (gear icon) → **Script Properties** → **Add script property**. Set the name to `BOOTH_PIN` and the value to a passphrase longer than a few digits, e.g. `blue-drone-42`.
   Each device asks for the PIN once and then remembers it. Without a PIN, anyone who gets the link can see the attendee list.
4. **Deploy.** Click **Deploy → New deployment**, then the gear icon → **Web app**, and set:
   - *Execute as:* **Me**
   - *Who has access:* **Anyone** (protected by the PIN), or **Anyone within *your company*** if you use Google Workspace and all staff log in with company accounts.

   Click **Deploy** and then **Authorize access**. With a personal Gmail account, Google shows "Google hasn't verified this app". Click **Advanced → Go to … (unsafe)**; this is your own script. Then copy the **Web app URL**.
5. **Put it on the booth devices.** Open the URL on each phone or laptop. On phones, use *Share → Add to Home Screen* so it opens like an app.
6. **Do a test run.** Submit a test lead, check the **Responses** tab, and then delete the test row.

## Using it at the booth

- **Booth staff** (top of the form): each person enters their name once per device. It is saved with every lead.
- **Find attendee**: type part of a name or company, e.g. `jon smi` or `northway`, then tap the person. First name, last name, company and job title fill in. The attendee list has no emails, so **Corporate Email is always typed in**.
  - A **✓ Submitted** badge means this attendee already has a lead. If you open them again, their earlier answers load, and submitting **updates the same row** instead of adding a duplicate.
  - **Registration cancelled** comes from the organiser's list. You can still select these people.
  - **Not on the list? Fill in manually** is for walk-ins. Each walk-in gets its own row.
- **Date** defaults to today and **Event Name** to `Edge_2026`. Both can be edited.
- **Weak Wi-Fi**: if a save fails, the entry is kept on the device and a yellow banner shows how many entries are waiting. They sync automatically when the connection comes back, or when you tap **Sync now**. Search keeps working from the list saved on the device. Leave the page open, or at least reopen it on the same device, until the banner is gone.

## Checking the results

Open the Google Sheet and go to the **Responses** tab. To get an Excel file, use **File → Download → Microsoft Excel (.xlsx)**.

| Columns | Meaning |
| --- | --- |
| Submitted At / Updated At | When the lead was first saved and last changed |
| Captured By | Booth staff name |
| Date … WMS | The form answers |
| Access Code, Zip, Country, Registration Status | Copied from the attendee list (empty for walk-ins) |
| Attendee Key, Response ID | Internal IDs the app uses to update the right row. Don't edit these. |

You can add your own columns (e.g. *Follow-up owner*, *Notes*) to the right. The app never overwrites them.

## Making changes

- **New attendee list:** paste the new rows into the attendee tab, keeping the header row, then tap **↻** next to the search box.
- **Questions or default event name:** edit `FORM` in `Code.gs`. Each question's `header` is its column name in Responses, so don't rename headers mid-event.
- **Publish code changes:** saving in the editor does not change the live link. Go to **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**. The URL stays the same.

## Limits and notes

- Google shows a small "This application was created by a Google Apps Script user" banner on the page. This can't be removed on the free setup.
- Each save takes about 1–2 seconds. Google's free quotas are far above what a booth needs.
- The attendee list is personal data. Don't commit the organiser's spreadsheet to this repository; `.gitignore` blocks `.xlsx`/`.csv` files. Booth devices cache the list in the browser for offline search, so clear site data on shared devices after the event.
