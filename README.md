# Bulk Print

Right-click one or more folders in Windows Explorer and print every Word, Excel, PowerPoint, PDF, email (`.eml`) and image (`.jpg .jpeg .png`) file in them to the default printer. Images print in colour, 4 per page.

## Files

| File | Purpose |
|---|---|
| `Bulk-Print.ps1` | The print script. Run by the Send To shortcut. |
| `Install-SendTo.cmd` | Double-click to install or update. Runs `Install-SendTo.ps1`. |
| `Install-SendTo.ps1` | Downloads the latest `Bulk-Print.ps1` from GitHub, then adds **Send to → Bulk Print** to the right-click menu for the current user. No admin rights needed. |
| `Uninstall-SendTo.cmd` | Removes the Send To shortcut. |
| `Setup-ImageColour.cmd` | One-time per PC, only if images print in black and white: opens the printer's own settings window; choose Colour and click OK. Saved for image jobs only. |

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

- Prints `.doc .docx .xls .xlsx .ppt .pptx .pdf .eml .jpg .jpeg .png` in each selected folder, in name order. Subfolders are not included.
- Emails print exactly as Outlook's Quick Print does: whole email, Memo Style, default printer. Each email briefly opens in Outlook (the script asks for the window to be minimised; whether it flashes on screen first is not verified) and closes again; Outlook itself is never closed. If Outlook was not running, it starts for each email (about 6 s each) and closes again afterwards.
- If a folder contains ZIP files, asks once per folder whether to extract and print their supported contents.
- Images (including any inside ZIPs) print together as one job at the end of each folder: portrait page (set `$imagePageLandscape = $true` at the top of `Bulk-Print.ps1` for landscape), 2 x 2 grid, each image fitted to its quarter without stretching, file name underneath, colour requested. Phone photos are turned upright from their rotation tag; transparent areas print white. Images are shrunk to 300 dpi at their printed size so the job stays small. An unreadable image leaves a grey note in its place and is listed as failed.
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
- **Images: HEIC** (iPhone's default format) **is not supported** and is listed as skipped. Set the iPhone to "Most Compatible" (JPEG) or convert first.
- **Colour: documents follow the printer, images ask for colour.** Word, Excel, PowerPoint, PDFs and emails print with the printer's default, which the script cannot change for those programs. For black-and-white documents, set the printer to black and white in **Settings → Printers & scanners → printer → Manage → Printing preferences** (affects only your Windows account). Images request colour on their own job, which overrides that default — confirmed with Windows' generic network-printer driver (Microsoft IPP Class Driver). **Some manufacturer drivers (e.g. FUJIFILM) ignore the request** and print images in black and white, because their own colour setting decides. Fix: double-click `Setup-ImageColour.cmd` once on that PC, choose Colour in the printer's window (e.g. *Output Color: Color*) and click OK. The settings are saved per user and printer in `%LOCALAPPDATA%\BulkPrint\` and used for image jobs only; documents keep the black-and-white default. Run it again after a driver update or a printer change. The log's `IMAGES` line shows the printer, driver, colour request and whether a saved profile was used.
- **Images: landscape photos print smaller** than portrait ones on the default portrait page, because photos are never rotated to fit.
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

**Image printing (15 Oct 2026):** `.jpg .jpeg .png` print in colour, 4 per landscape page, with the file name under each image (Ferry's choices: landscape page, no rotation to fit, captions, JPEG + PNG). The script draws the pages itself with .NET `PrintDocument`: Windows' Print Pictures wizard cannot be scripted, and a Word or PDF route was slower or needed an extra tool.
- Preview test (pages rendered to PNG, nothing printed), made-up images: landscape and portrait fit their cells; a phone-style image stored sideways with orientation tag 6 prints upright; a transparent PNG prints on white; a long file name fits its caption; a corrupt file leaves a grey note and is listed as failed; 6 images make 2 pages.
- Real print of one page on a network colour printer: colour, margins, orientation and captions all correct; 4 photos made a 254 KB job.
- Found in testing: `[Math]::Min(1, x)` picks the whole-number version of `Min` when the first argument is `1`, rounding the scale to 0 or 1, so JPEGs shrank to 1x1 pixel and printed blank. Fixed by writing `1.0`.

**Image follow-up (15 Oct 2026):** image pages switched to portrait by default (Ferry), set by `$imagePageLandscape` at the top of the script; preview checked. Each image job now logs an `IMAGES` line (printer, driver, colour support, colour requested). Colour test: with this PC's printer preference set to black and white, an image page still printed in colour (IPP class driver), so documents can default to black and white while images print in colour. On another PC a manufacturer driver printed images in black and white — not yet solved.

**Colour profile for manufacturer drivers (15 Oct 2026):** on a PC with `FF Apeos C4570 PCL 6` (FUJIFILM), images printed black and white although colour was requested — the driver's private "Output Color" setting overrides the standard request. `Setup-ImageColour.cmd` now saves the printer's full settings (DEVMODE, including the driver's private section) from its own Printing Preferences window; image jobs load them with `SetHdevmode`, then apply the layout. Tested without paper: a profile captured here (3,844 bytes, 3,624 driver-private) reached the job byte for byte and set its colour; no profile → standard request as before; damaged profile → warning and fallback. The dialog step and the FUJIFILM result are tested on the other PC.

### Before 15 Oct 2026

Original version: single folder, Send To install, ZIP prompt, log per run.
