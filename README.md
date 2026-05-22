# PRCe360 Service List Converter

![NMPRC Logo](brand/NMPRC_logo-1.png)

Converts a telecom service list (Excel or Word) into the GridTemplate format used by Grid, with columns for First Name, Last Name, Type, and Email.

---

> **Disclaimer:** This repository is an independent passion project and is in no way affiliated with, endorsed by, or produced by the New Mexico Public Regulation Commission (NMPRC). This tool is provided as-is for informational and administrative convenience only. The NMPRC makes no representations or warranties of any kind, express or implied, regarding the accuracy, completeness, reliability, or suitability of this tool for any purpose. The NMPRC assumes no responsibility and shall not be liable for any errors, omissions, data loss, or damages of any kind arising from the use of or reliance on this tool. Use of this tool is entirely at your own risk.

---

## Table of Contents

- [What This App Does](#what-this-app-does)
- [Installing on Windows](#installing-on-windows)
- [Using the App](#using-the-app)
- [Accepted File Types and Requirements](#accepted-file-types-and-requirements)
- [Limitations and Known Restrictions](#limitations-and-known-restrictions)
- [Uninstalling](#uninstalling)
- [Building the Installer (Developers Only)](#building-the-installer-developers-only)
- [Roadmap](#roadmap)

---

## What This App Does

Service lists from regulatory filings often come in Excel or Word format with a party's full name and email address in each row. This app reads those lists and produces a properly formatted Excel file that can be imported directly into Grid, with separate First Name, Last Name, Type, and Email columns.

When a row contains a company name instead of a person's name, the app tries to find a person's name by reading the email address (for example, `gschneider@example.com` becomes "G Schneider"). The **Type** column is always set to **Individual**.

---

## Installing on Windows

> **No Python or technical knowledge is required.** The installer is self-contained — it includes everything the app needs to run.

### Step 1 — Get the installer file

You need the file named something like `PRCe360ServiceListConverter-1.0.0-win-amd64.msi`. This file should have been shared with you by the person who built it. Save it somewhere easy to find, such as your **Downloads** folder or **Desktop**.

### Step 2 — Run the installer

1. Double-click the `.msi` file.
2. If Windows shows a blue "Windows protected your PC" warning, click **More info**, then click **Run anyway**. This appears because the installer is not signed by a large software company — it is safe to proceed.
3. When prompted by User Account Control ("Do you want to allow this app to make changes?"), click **Yes**. Administrator rights are required because the app installs to `Program Files`.

### Step 3 — Follow the installer

1. Click **Next** on the welcome screen.
2. Click **Next** again to accept the default install location (`C:\Program Files\PRCe360ServiceListConverter`).
3. Click **Install**.
4. Click **Finish** when the installer completes.

### Step 4 — Find the app

After installation you will find **PRCe360 PRCe360 Service List Converter** in two places:

- A shortcut on your **Desktop**
- Under **Start Menu → PRCe360 Service List Converter**

Double-click either one to open the app.

---

## Using the App

The app window has two file fields and a **Convert** button.

### 1. Input file

Click **Browse…** next to "Input file" and navigate to your service list. This is the file you received — either an Excel spreadsheet (`.xlsx`) or a Word document (`.docx`).

### 2. Output file

The app fills this in automatically based on your input file name (for example, `output_mylist.xlsx` in the same folder). You can click **Browse…** to save it somewhere else or give it a different name.

### 3. Convert

Click the **Convert** button. Progress messages appear in the log area at the bottom of the window. When you see "Done," the output file is ready.

If the list has more than 500 entries, the app automatically splits the output into multiple files named `output_1.xlsx`, `output_2.xlsx`, and so on. Import each file into Grid separately.

---

## Accepted File Types and Requirements

### Input file — what is accepted

| Format | Extension | Notes |
|--------|-----------|-------|
| Excel workbook | `.xlsx` | Must have Name in column A and Email in column B, with data starting on row 2 (row 1 is the header). |
| Word document | `.docx` | Must contain a table with column headers named **Name** and **Email**. The app reads the first table that has both headers. |

**What is not accepted:**

- `.xls` — old Excel 97–2003 format. Open the file in Excel and save it as `.xlsx` first.
- `.csv` — comma-separated text files. Open the file in Excel and save it as `.xlsx` first.
- `.pdf` — not supported. You must obtain the source Excel or Word file.
- `.ods`, `.numbers`, or other spreadsheet formats — not supported. Convert to `.xlsx` in Excel or Google Sheets first.
- Scanned images or image-only PDFs — not supported.

### GridTemplate file — what is required

- Must be an Excel workbook (`.xlsx`).
- The first row must contain the following column headers (spelling and capitalization must match exactly):
  - `First Name`
  - `Last Name`
  - `Type`
  - `Email`

### Output file

- Always saved as an Excel workbook (`.xlsx`), regardless of the input format.

---

## Limitations and Known Restrictions

**Name guessing is not perfect.** When a row has a company name (or no name), the app tries to extract a person's name from the email address. This works well for common patterns like `jane.doe@example.com` or `jdoe@example.com`, but it can produce unexpected results for unusual email formats. Always review the output before importing into Grid, especially for rows where only an email address was available.

**"Type" is always "Individual."** There is no option to set a different type. All output rows will have `Individual` in the Type column.

**Large lists are split into 500-row chunks.** If your service list has more than 500 entries, the output will be multiple files (e.g. `output_1.xlsx`, `output_2.xlsx`). Each file must be imported into Grid separately.

**Excel input must use columns A and B.** The app always reads Name from the first column and Email from the second. If your source spreadsheet has them in different columns, or has extra header rows above row 1, you will need to rearrange it before converting.

**Word input requires a proper table.** If the Word document lists names and emails in plain paragraphs rather than a table, the app will not be able to find the data and will report an error. You will need to copy the data into Excel and use the `.xlsx` input path instead.

**Multi-value email cells.** If a cell contains more than one email address separated by semicolons, the app uses only the first address.

---

## Uninstalling

Open **Settings → Apps** (or **Control Panel → Programs and Features** on older versions of Windows), find **PRCe360 PRCe360 Service List Converter** in the list, and click **Uninstall**.

Alternatively, run the original `.msi` file again and choose **Remove**.

---

## Running from Source (Linux / macOS / Windows without the installer)

### Requirements

- Python 3.12 or later
- [uv](https://github.com/astral-sh/uv) for dependency management
- **Linux only:** tkinter is not pip-installable — install it via your system package manager before running the GUI:
  ```
  sudo apt install python3-tk      # Debian / Ubuntu
  sudo dnf install python3-tkinter # Fedora / RHEL
  sudo pacman -S tk                # Arch
  ```

### Steps

1. Clone the repository.
2. Install dependencies:
   ```
   uv sync
   ```
3. Run the GUI:
   ```
   uv run python gui.py
   ```
   Or use the CLI directly:
   ```
   uv run python convert.py <source.xlsx|.docx> <template.xlsx> <output.xlsx>
   ```

---

## Building the Installer (Developers Only)

### Requirements

- Windows 10 or 11
- Python 3.12 or later — download from [python.org](https://www.python.org/downloads/). During installation, check **"Add Python to PATH"**.
- The `brand\` folder must be present in the project root (it is gitignored — copy it manually if cloning fresh). It must contain `NMPRC_logo-1.png`.

### Steps

1. Clone or copy this repository to the Windows machine. Ensure `brand\NMPRC_logo-1.png` is present.
2. Open a Command Prompt in the project folder.
3. Run:

   ```bat
   build_windows.bat
   ```

   The script automatically installs `openpyxl`, `python-docx`, `lxml`, `Pillow`, and `cx_Freeze`, then builds the installer. The finished file is written to:

   ```
   dist\PRCe360ServiceListConverter-1.0.0-win-amd64.msi
   ```

4. The MSI is ready to distribute or publish as a GitHub release.

### Upgrading

Increment `version` in `setup_msi.py` before building. The `upgrade_code` GUID must stay the same — Windows uses it to recognise the new build as an upgrade rather than a separate product.

---

## Roadmap

- **Optional shortcuts during install** — the current installer always creates Desktop and Start Menu shortcuts. A future release will prompt the user to opt in or out during installation. This requires migrating from cx_Freeze's MSI builder to a more capable installer tool such as Inno Setup.
