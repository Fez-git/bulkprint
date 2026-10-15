# Bulk Print

Right-click one or more folders in Windows Explorer and print every Word, Excel, PowerPoint, PDF and email (`.eml`) file in them to the default printer.

## Files

| File | Purpose |
|---|---|
| `Bulk-Print.ps1` | The print script. Run by the Send To shortcut. |
| `Install-SendTo.cmd` | Double-click to install or update. Runs `Install-SendTo.ps1`. |
| `Install-SendTo.ps1` | Downloads the latest `Bulk-Print.ps1` from GitHub, then adds **Send to → Bulk Print** to the right-click menu for the current user. No admin rights needed. |
| `Uninstall-SendTo.cmd` | Removes the Send To shortcut. |

## Requirements

- Windows 10 or 11.
- Desktop Microsoft Office (Word, Excel, PowerPoint). Office on the web and the Microsoft Store edition cannot be automated.
- A default printer. Settings → Printers: switch off "Let Windows manage my default printer" and pick one.
- A PDF program that offers **Print** when you right-click a PDF (Adobe Reader does; Edge may not).
- For `.eml`: **classic** Outlook, set as the program that opens `.eml` files (double-click one to check). The new Outlook app cannot be scripted.

## Install

1. Put the files in a permanent folder, e.g. `C:\Tools\BulkPrint`. The shortcut stores this path, so moving the folder later breaks it.
   - Without git: on GitHub, **Code → Download ZIP**, then extract.
   - With git: `git clone https://github.com/Fez-git/bulkprint.git C:\Tools\BulkPrint`
2. ZIP download only — clear the "downloaded from the internet" mark:
   `Get-ChildItem C:\Tools\BulkPrint | Unblock-File`
3. Double-click `Install-SendTo.cmd`. It downloads the latest `Bulk-Print.ps1`, then should say **Installed: Send to -> Bulk Print**.
4. Test on a folder holding one small file: right-click the folder (Windows 11: **Show more options**) → **Send to → Bulk Print**.

**Update:** double-click `Install-SendTo.cmd` again. It fetches the latest `Bulk-Print.ps1` from GitHub and keeps the previous copy as `Bulk-Print.ps1.bak`. To roll back, delete `Bulk-Print.ps1` and rename the `.bak`.

- The download is checked before it replaces anything: it must be complete, valid PowerShell and actually be `Bulk-Print.ps1`. If it fails any check, or the PC is offline, the existing copy is kept.
- Only `Bulk-Print.ps1` is updated. If the installer or uninstaller changes, download the ZIP again.
- Folder set up with `git clone`: the installer skips the download — run `git -C C:\Tools\BulkPrint pull` instead.
- Run the installer again if the folder moved.

**Uninstall:** double-click `Uninstall-SendTo.cmd`, then delete the folder.

## What it does

- Prints `.doc .docx .xls .xlsx .ppt .pptx .pdf .eml` in each selected folder, in name order. Subfolders are not included.
- Emails print exactly as Outlook's Quick Print does: whole email, Memo Style, default printer. Each email briefly opens in Outlook (the script asks for the window to be minimised; whether it flashes on screen first is not verified) and closes again; Outlook itself is never closed. If Outlook was not running, it starts for each email (about 6 s each) and closes again afterwards.
- If a folder contains ZIP files, asks once per folder whether to extract and print their supported contents.
- Lists everything else as skipped. `.md` files (including triage `.eml.md` notes) are logged but left out of the on-screen summary.
- Writes a log, `Bulk-Print-<date-time>.log`, into each folder it prints (skipped, with a warning, if the folder is read-only).
- Ends with a summary: failed files in red, skipped files in yellow, each labelled with its folder.

## Known limitations

- **PDFs are handed to the PDF program and not tracked.** The log says `SUBMITTED`, not `PRINTED`. Whether it prints depends on that program; Adobe Reader stays open afterwards. PDFs and Office files can come out of the printer interleaved.
- **Extracted ZIP contents are deleted when you press Enter at the end.** Wait until PDFs from ZIPs have printed first.
- **Password-protected files are not printed** — they are listed as skipped or failed. Print them by hand.
- **Office can still hang on an unexpected dialog** (a first-run or licence prompt, for example) because it runs hidden. If the window sits there doing nothing, end `WINWORD`, `EXCEL` or `POWERPNT` in Task Manager.
- **Read-only folders** (some network shares) cannot hold the log. The script prints them anyway, with no log, and shows a yellow warning when it starts that folder and again in the end summary.
- **Emails print in full** — there is no page limit; Outlook's scripting interface has no page range. Attachments inside an `.eml` are not printed (Outlook's default).
- **An email is matched to its Outlook window by subject.** If Outlook shows the email under a different subject than the file's header (unusual encodings), it fails after 60 s rather than risk printing the wrong email.
- **Very large selections** may exceed the Windows command-line length limit for Send To. Not measured.
- **Work PCs:** if the window flashes and closes, a Group Policy execution policy is probably blocking scripts. `-ExecutionPolicy Bypass` does not override that; IT has to allow it.

## Change history

### 15 Oct 2026 — review fixes

| # | Problem | Fix |
|---|---|---|
| — | Selecting more than one folder printed nothing: the script accepted one argument, and PowerShell rejected the call before any code ran. | Accepts any number of folders and processes each in turn. One failing folder no longer stops the rest; Enter is pressed once at the end. |
| 1 | PDFs inside a ZIP were handed to the PDF program, then the temp folder was deleted immediately — often before the program had opened the file. | Temp folders are deleted only after the final Enter. |
| 2 | Word and PowerPoint print in the background by default; the script quit the app straight after `PrintOut()`, which can cancel a job still spooling. | Word: `PrintOut($false)`. PowerPoint: `PrintOptions.PrintInBackground = 0`. Both return only once the job is spooled. |
| 3 | The log was created before the file list was built, so every run listed its own log — and every earlier log — as skipped. | File list is built first and excludes `Bulk-Print-*.log`. |
| 4 | Failed prints went only to the log. The summary could show green "No files were skipped" when every file had failed. | Failures are collected and shown in red in the summary. |
| 5 | PowerPoint runs once per user, so the script attached to the user's open PowerPoint and `Quit()` closed it, with their presentations. | If PowerPoint was already running, the script closes only its own presentations and leaves PowerPoint open. |
| 6 | A password-protected file opened a password box in the hidden Office window; the script hung forever. | Encrypted `.docx/.xlsx/.pptx` are detected from the file header (`D0 CF 11 E0` instead of `PK`) and skipped without opening Office. `.doc/.xls` are opened with a dummy password and `.ppt` with `file::password::`, so Office raises an error instead of prompting. |
| 8 | A read-only folder could not hold the log: the script stopped that folder, so nothing in it printed. | Prints the folder without a log. Yellow warning when the folder starts and in the end summary. Tested on a folder with file creation denied: warning shown, folder processed, no log left behind, the other folder still logged normally. |

**How fixes 5 and 6 were tested** (with `PrintOut` removed, so nothing printed): the script's own `Print-OfficeFile` ran against a protected and an unprotected copy of each of the six formats. All six protected files were skipped or failed without hanging; all six unprotected files opened. With a user presentation open in PowerPoint, it was still open after the script ran. Fixes 1–4 and multi-folder were tested with a dry run on folders holding no printable files. Actual print output has not been tested.

**Gotcha found in testing:** passing a dummy `WriteResPassword` to Excel's `Workbooks.Open` makes every *unprotected* workbook fail to open. It is left as `[Type]::Missing`.

**Installer self-update (15 Oct 2026):** `Install-SendTo.cmd` now downloads the latest `Bulk-Print.ps1` before installing. Tested: cut-off download and wrong file both rejected with the local script kept; newer version replaces it with a `.bak`; repeat run reports "already the latest"; git clone skips the download; no temp files left behind. Needs the repo to be public — a private repo returns 404 without login.

**Email printing (15 Oct 2026):** `.eml` files print through Outlook, replacing the separate `Print-Email.ps1` (removed). Outlook cannot open `.eml` through its scripting interface, so each email is opened the way a double-click does, its window is found by subject, printed with `PrintOut()` and closed without saving.
- Tested with printing removed, on made-up emails (short; long with attachment and folded subject; French/Chinese encoded subject): with Outlook open and another email showing — including one with the same subject — each was found and closed in about 3 s and the user's window left open; with Outlook closed, about 6 s each and Outlook closed again afterwards.
- Real print of two emails: one page each, job named "Microsoft Outlook - Memo Style", output matched a manual Outlook print. `PrintOut()` returned only once the job was in the printer queue.
- Found in testing: `GetActiveObject` cannot see a freshly started Outlook for minutes, so `New-Object -ComObject Outlook.Application` is used; an Outlook with no main window exits when its last email window closes, so the connection is re-checked per email — with a method call, because PowerShell returns `$null` rather than an error when reading a property of a disconnected COM object.
- Also fixed: a log write failing (the log was briefly locked by another process; cause not identified) stopped the whole folder. Log writes now retry, then warn and carry on.

### Before 15 Oct 2026

Original version: single folder, Send To install, ZIP prompt, log per run.
