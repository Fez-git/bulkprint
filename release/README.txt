BULK PRINT {{VERSION}}
======================

Right-click one or more folders and print every Word, Excel, PowerPoint, PDF,
email (.eml) and image (.jpg .png) file in them to your default printer.
Images print 4 per page, in colour.


INSTALL (about 1 minute)
------------------------
1. Put this folder somewhere permanent, for example C:\Tools\BulkPrint.
   The right-click shortcut points to where the files are: if you move or
   delete the folder later, run Install.cmd again from the new place.

   Tip: before extracting, right-click the ZIP > Properties > tick "Unblock"
   > OK. That stops Windows asking for permission on every file.

2. Double-click Install.cmd.
   Windows may say it "could not verify the publisher". The files are plain
   text you can open in Notepad; the tool is not signed with a paid
   certificate. Click Run.

3. Read the short check at the end. It tells you what this PC is missing for
   each file type (Office, classic Outlook for emails, a default printer).

4. Test it on a folder with one small file: right-click the folder
   (Windows 11: "Show more options" first) > Send to > Bulk Print.


WHAT YOU NEED
-------------
- Windows 10 or 11 and the desktop version of Microsoft Office.
- A default printer.
- For emails (.eml): classic Outlook as the program that opens .eml files.
- For PDFs: a PDF program that offers Print when you right-click a PDF
  (Adobe Reader does).


COLOUR
------
Word, Excel, PowerPoint, PDF and email follow your printer's own default.
If you want documents in black and white, set the printer to black and white
in Settings > Printers & scanners > your printer > Manage > Printing
preferences. Images always ask for colour.

Some printer drivers (for example FUJIFILM) ignore that request. If images
print in black and white, double-click files\Setup-ImageColour.cmd once, choose
Colour in the window that opens, and click OK.


UPDATE
------
Double-click Install.cmd again. It downloads the newest Bulk-Print.ps1 from
the project's Releases page and keeps the old one as Bulk-Print.ps1.bak.
(Changes to the installer itself need a new download of the ZIP.)


UNINSTALL
---------
Double-click files\Uninstall-SendTo.cmd, then delete the folder.


MORE
----
Known limits, change history and source:
https://github.com/Fez-git/bulkprint
Licence: MIT (see LICENSE).
