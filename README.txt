Digonto PDF 3.0  -  create, shrink, organize, sign and edit PDF files
=====================================================================
Part of the Digonto apps (digonto.dev). Install with the installer in the
folder above this one (see ../README.txt):   bash install-linux.sh pdf
Windows: install-windows.bat   macOS: install-macos.command

HOW IT WORKS (version 3.0)
--------------------------
Every PDF opens in its own tab, like a normal PDF viewer. The first tab, "Start",
shows your recent PDFs, with buttons to open a PDF or create a new one.
Left: the page list (only active when the PDF has more than one page).
Right: the tools panel (active when a PDF is open). Click a tool and a bar with
its actions appears above the page.

Tools:  Pages      insert pages from another PDF, insert blank page, move left /
                   right, rotate, delete, extract the selected pages
        Sign       draw a signature or load a picture, place it, move and resize
                   it, Apply. Signatures are kept for next time
        Edit text  change a line, add text, delete text in an area, white-out
        Shrink     makes a smaller copy in a new tab; the original stays as it is
        Add files  pictures, Word (.docx), Excel (.xlsx), text as new pages
        More       page numbers, watermark, password, document info

File menu: Open, Open recent, Create PDF (pictures, Word, Excel, text -> one PDF
or one per file), Save, Save as, Export as (pictures, selected pages), Close tab.
A changed PDF shows * after its name on the tab; the * goes away when you save.
Closing a changed tab asks: Save, Don't save or Cancel. Undo: Ctrl+Z.


ABOUT DEPENDENCIES AND SECURITY
-------------------------------
Digonto PDF uses exactly TWO libraries, listed in requirements.txt:
  - PySide6-Essentials 6.10.3   (Qt for Python - the window and buttons)
  - PyMuPDF 1.27.2.3            (the PDF engine)
  (shiboken6 is the small helper that belongs to PySide6.)

The installers download them from PyPI with --require-hashes: each file is
checked against its SHA-256 fingerprint written in requirements.txt. If a
file was changed by anyone, pip refuses to install it.

Everything else is Digonto's own code, readable in two files:
  digonto_pdf.py       the app
  digonto_convert.py   Word / Excel / text / picture -> PDF converters
The app itself never connects to the internet.

No Microsoft Office, LibreOffice or Ghostscript is needed. If you already
have LibreOffice or MS Office, Create PDF offers an optional "exact copy"
tick box for Word/Excel. If you install Ghostscript yourself, Shrink also
tries it and keeps the smaller result.


LIMITS OF THE OWN WORD/EXCEL CONVERTER
--------------------------------------
Keeps: text, headings, bold/italic/underline, colors, fonts (sans/serif/mono),
alignment, bullet and numbered lists, tables (merged cells, shading), pictures,
page breaks, page size and margins, Bengali/German/Italian letters.
Excel: values as Excel last saved them, number/date/percent/currency formats,
bold/colors/fills/borders, merged cells.
Not (yet): headers/footers, text boxes, charts, footnotes, comments, old
.doc/.xls files (save them as .docx/.xlsx first).


UNINSTALL
---------
  bash install-linux.sh --uninstall pdf       (Windows: install-windows.bat --uninstall pdf)
