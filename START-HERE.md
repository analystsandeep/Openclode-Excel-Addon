# Start here: no technical knowledge needed

## What this gives you

An **OpenCode** button on Excel's Home tab. Click it and a panel opens on the right where
you can chat with the same AI models you already use in the OpenCode desktop app. It can
see the cells you have selected, and it can write answers, formulas and tables back into
your sheet.

> You do **not** need to type any commands. You double-click one file and it does the rest.

---

## What you need (one time)

1. **Excel 2019, 2021 or Microsoft 365** on Windows.
2. The **OpenCode desktop app**, installed and working (you can already chat with a model in it).
3. **Node.js**, a small free program this add-in needs.
   Install it once from **https://nodejs.org**, choose the **LTS** version, then complete the
   installer. You never have to open it yourself.

---

## Install (once, about 1 minute)

1. **Unzip** this folder somewhere easy to find, for example
   `Documents\OpenCodeExcel`.
   Do not run it from inside the zip. Extract it first.
2. Double-click **`Start OpenCode for Excel.bat`**.
   - The first time it downloads a few parts. Wait until it prints **Ready**.
   - If Windows shows a blue "Windows protected your PC" box, click **More info**, then **Run anyway**.
   - If it says Node.js is missing, it opens the download page for you. Install Node, then
     double-click the `.bat` again.

---

## Use it (every time)

1. Open the **OpenCode desktop app** and leave it running. *This is what answers your questions.*
2. Double-click **`Start OpenCode for Excel.bat`** and wait for **Ready**.
3. Excel opens with a small **OpenCode for Excel** workbook and the panel already open on
   the right. The **OpenCode** button is on the **Home** tab.
4. You can close that starter workbook and use the **OpenCode** button in your own
   workbooks too — it stays for as long as Excel is open.
5. Select some cells, type your question in the box at the bottom, press **Send**.
   - To use a different AI model, use the model selector at the bottom of the panel. The list
      comes from your OpenCode app, and your choice is remembered. Leave it on **Default** to
      let OpenCode decide.
   - Click **Attach** for PDFs, images, documents, presentations, spreadsheets, CSVs, or
     text files. Files stay in the add-in's local scratch folder and are sent through OpenCode.
6. To make the **whole sheet** look professional, click **Format sheet** — the blue button at
   the top of the panel. It creates a real table, adds a summary strip, sets date/money
   formats, freezes the header, turns off gridlines, sets up printing and adds data bars.
   It works without asking the AI, and it is safe to click again. On a sheet that has no
   summary yet, it inserts three summary rows above the table; your original data stays
   intact.
7. Under each answer you get buttons:
   - **Copy** copies the answer.
   - **Insert below** writes the answer into the sheet below your selection.
   - **Insert table** writes a returned table as real rows and columns.
   - **Use as formula** puts a returned formula into the selected cell.

**Keep the black window open** while you use Excel. Closing it stops the add-in.

---

## If something goes wrong

| What you see | What to do |
|---|---|
| *"Node.js is not installed"* | Install it from the page that opened, then double-click the `.bat` again |
| The panel shows a **red dot / "Not connected"** | Open the OpenCode desktop app, wait ~10 seconds. If it stays red, close the black window and run the `.bat` again |
| Excel says *"We can't open this add-in from localhost"* | Double-click **`Fix Certificates.bat`**, choose **Yes** if Windows asks, close Excel, run the `.bat` again |
| Yellow *"ADD-IN ERROR"* mentions Visual Studio debugging | Save your work and close **all** Excel windows, then run `Start OpenCode for Excel.bat` again. Do not reopen the old starter workbook from Excel's Recent list. |
| No **OpenCode** button on the Home tab | Close Excel completely, then run the `.bat` again. It opens the starter workbook that activates the button. If it still does not show, use *Insert > My Add-ins > Developer Add-ins*. |
| *"Port 3000 is already in use"* | It is already running in another black window. Use that one, or close the other window first. |
| The panel opens but is blank | Close Excel and run the `.bat` again; if it repeats, run **`Fix Certificates.bat`** |
| The answer takes long | That's the model thinking. Press **Stop** to cancel |
| The reply says the model is unavailable or returns no text | The add-in retries once with a verified free OpenCode model and updates the model selector. If no fallback is available, choose another model manually. |
| *Format sheet* says it did not load | Close and reopen the panel (click **OpenCode** on the Home tab), or close the black window and run the `.bat` again |
| The summary shows `#` instead of numbers | Click **Format sheet** again — it widens the columns to fit |
| Nothing happens at all | Look at the black window — it prints plain-English messages telling you what to do |

---

## Stopping and removing

- **Stop:** close the black window (the OpenCode button stays in Excel, but shows "Not connected").
- **Remove from Excel:** close Excel, then double-click **`Remove from Excel.bat`**.
- **Remove completely:** delete the folder.

---

## Privacy

Everything runs on your own computer. The Excel panel talks to a small local program on
`localhost`, which talks to your OpenCode app — nothing is hosted by this add-in. Bounded workbook
samples, your prompt, and files you attach are sent to the AI model/provider configured inside
OpenCode. The add-in runs its sessions in a locked folder where the assistant is forbidden from
running shell commands or editing files.

---

*Comfortable with terminals? Everything above is also scriptable — see `README.md`.*
