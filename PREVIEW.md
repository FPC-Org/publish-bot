# Publish-Bot Preview Tool

Use `preview.py` to convert a Word research summary draft into a formatted HTML preview — exactly as it would look when published. Nothing gets published, pushed, or saved anywhere outside your machine.

---

## What You Need (One-Time Setup)

### 1. Python

You likely already have Python through ArcGIS Pro. To confirm, open a terminal and run:

```powershell
"C:\Program Files\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\python.exe" --version
```

If that prints a version number, you're good.

### 2. Pandoc

Pandoc is the document converter the tool relies on. Download and install it from:

**https://pandoc.org/installing.html**

After installing, confirm it works:

```powershell
pandoc --version
```

### 3. The publish-bot repository

Clone or download the `FPC-Org/publish-bot` repository to your machine. The easiest way is to open a terminal in the folder where you want it and run:

```powershell
git clone https://github.com/FPC-Org/publish-bot.git
```

Or download it as a ZIP from GitHub and extract it.

### 4. Python dependencies

In a terminal, navigate into the `publish-bot` folder and run:

```powershell
"C:\Program Files\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\python.exe" -m pip install pypandoc gspread
```

---

## Running the Preview

Open a terminal, navigate to the `publish-bot` folder, and run:

```powershell
"C:\Program Files\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\python.exe" preview.py
```

A file picker dialog will open. Select your `.docx` draft and click **Open**.

The tool will convert your document and automatically open the HTML preview in your default browser.

### Passing the file path directly

If you prefer to skip the dialog, you can pass the file path as an argument:

```powershell
"C:\Program Files\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\python.exe" preview.py "C:\path\to\your\report.docx"
```

---

## Where the Output Goes

The HTML preview is saved to:

```
publish-bot\preview_output\<YourDocumentName>\index.html
```

It stays there until you run the tool again on the same file, at which point it is overwritten. You can share this folder or open `index.html` manually in any browser.

---

## Document Requirements

Your Word document needs to include the following labeled metadata lines near the top:

| Field | Example |
|---|---|
| `Title:` | Regionwide 28: Quantifying Carryover Effects |
| `Author:` | Timothy Albaugh |
| `New or Updated:` | Updated |
| `Species:` | Pine |

The preview tool reads these fields to build the report header. If any are missing the tool will still run, but the header may show placeholder values.

---

## Troubleshooting

**"No module named pypandoc"**
Run the pip install step from the setup section above.

**"Pandoc not found" or conversion fails immediately**
Pandoc is not installed or not on your PATH. Install it from pandoc.org and restart your terminal.

**The file picker does not open**
Your Python environment may not have `tkinter`. Pass the file path directly as a command line argument instead (see above).

**The preview opens but images are missing**
Some images in Word are saved as EMF (Windows Enhanced Metafile) format. The tool converts these automatically, but the conversion requires PowerShell. If you are on a machine where PowerShell is restricted, those images may not appear.
