# DocMerge

**Combine `.txt`, `.doc`, `.docx` and `.pdf` files into one editable Word document. Free, open source, and runs entirely on your Mac.**

![macOS](https://img.shields.io/badge/macOS-12%2B-lightgrey)
![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Offline](https://img.shields.io/badge/network-none-brightgreen)

<p align="center">
  <img src="docs/screenshot.png" alt="DocMerge interface" width="640">
</p>

---

## The problem

You have forty call transcripts as `.txt`, a handful of `.docx` briefs, and a few `.pdf` reports. You want them in one document you can read, search and edit.

Every tool you reach for solves a slightly different problem:

- **Terminal** (`cat *.txt > merged.txt`) handles plain text only, and loses everything else
- **PDF mergers** combine PDFs into a PDF, which you then cannot edit
- **Word's Insert > Text from File** handles Word documents but not PDFs, one file at a time
- **Online mergers** want you to upload the documents; often a non-starter for client or government material

DocMerge takes mixed formats in, and gives you one **editable `.docx`** out. Nothing leaves your machine.

---

## How it compares

Prices are US list, checked August 2026. Follow the links; they change often.

| Tool | Price | Mixed formats in | Editable Word out | Fully offline |
|---|---|---|---|---|
| **DocMerge** | **Free, MIT** | **Yes** | **Yes** | **Yes** |
| [Adobe Acrobat Pro](https://www.adobe.com/acrobat/pricing.html) | US$19.99/mo annual, US$29.99/mo month-to-month | Yes | Export step required | Account required |
| [Adobe Acrobat Standard](https://www.adobe.com/acrobat/pricing.html) | US$14.99/mo annual | Yes | Export step required | Account required |
| [PDF Expert Premium](https://pdfexpert.com/) | US$79.99/yr, or US$139.99 lifetime | PDF focused | Conversion step | Yes |
| [PDF Reader Pro](https://www.pdfreaderpro.com/) | US$79.99 one-time | PDF focused | Conversion step | Yes |
| [Foxit PDF Editor](https://www.foxit.com/pdf-editor/) | ~US$139 perpetual | PDF focused | Conversion step | Yes |
| PDF Joiner & Merger (Mac App Store) | US$4.99 one-time | PDF only | No | Yes |
| Various "PDF Merger" subscription apps | ~US$4.99/mo, ~US$39.99/yr | PDF only | No | Varies |
| iLovePDF / Smallpdf (web) | Free tier, paid plans above | PDF focused | Conversion step | **No, upload required** |

### Being straight about scope

The commercial tools above are full PDF suites. They do editing, redaction, OCR, form filling, e-signatures and page-level manipulation. **DocMerge does none of that**, and if you need those things, buy one of them; they are good products and worth the money.

DocMerge does one narrow job the expensive tools handle awkwardly: taking a pile of *mixed* file formats and producing a single editable Word document, without a subscription and without uploading anything.

**What DocMerge does not do:**

- Preserve fonts, colours, images or layout. It extracts text, so formatting is intentionally dropped
- Merge PDFs *as* PDFs. Output is always `.docx`
- OCR. Scanned image-only PDFs are flagged in the output rather than read

---

## Install

Requires macOS 12 or later. Python 3.9+ (Python 3.10+ recommended).

```bash
git clone https://github.com/YOUR_USERNAME/docmerge.git
cd docmerge
bash "Build App.command"
```

The build script creates a virtual environment, installs dependencies, compiles a native `DocMerge.app` with PyInstaller, and installs it to `/Applications`. Takes two to three minutes.

Because the app is compiled on your own machine, Gatekeeper does not block it. If macOS blocks the build script itself after a download, clear the quarantine flag first:

```bash
xattr -r -d com.apple.quarantine .
```

### Run without building

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install python-docx pypdf
python3 docmerge.py
```

---

## Usage

1. Open DocMerge from Launchpad, Spotlight or the Dock
2. **Add Files** or **Add Folder** (folders pull in every supported file, sorted alphabetically)
3. Reorder with **Up** and **Down** if sequence matters
4. **Merge to Word Document** and choose where to save

The status line reports each stage, down to the page number on long PDFs. Finder opens with the result when it finishes.

### Command line

```bash
python3 docmerge.py output.docx file1.txt file2.pdf file3.docx
```

Useful in scripts and Automator actions.

---

## Output structure

- Title page with the source file count
- Each source file begins on a new page, with its filename as a Heading 1
- Word tables are flattened to pipe-separated rows
- PDFs are extracted page by page; unreadable pages are marked inline rather than silently dropped

---

## Supported formats

| Extension | Method |
|---|---|
| `.txt`, `.md` | Direct read, with encoding fallback (UTF-8, CP1252, Latin-1) |
| `.docx` | `python-docx`, including table contents |
| `.doc`, `.rtf` | macOS `textutil`, built into the OS, no extra dependency |
| `.pdf` | `pypdf` text extraction |

---

## Privacy

DocMerge makes no network calls. There is no telemetry, no analytics, no update check and no account. The dependency list is two libraries, both widely used and auditable. Read [`docmerge.py`](docmerge.py); it is a single file, roughly 600 lines.

This matters if you handle client, legal, medical or government documents where uploading to a web merger would breach your obligations.

---

## Built with

- [python-docx](https://github.com/python-openxml/python-docx) for Word reading and writing
- [pypdf](https://github.com/py-pdf/pypdf) for PDF text extraction
- [PyInstaller](https://pyinstaller.org/) for the macOS app bundle
- Tkinter for the interface, so there is no heavyweight UI framework to install

---

## Contributing

Issues and pull requests are welcome. Areas that would genuinely help:

- **OCR support** for scanned PDFs, likely via Tesseract or macOS Vision
- **Windows and Linux builds**; the engine is cross-platform, only `textutil` and the Finder reveal are macOS specific
- **Drag and drop** file adding
- **Formatting preservation** for `.docx` sources, carrying styles rather than plain text
- **Saved presets** for repeated merges of the same folder

---

## Licence

MIT. See [LICENSE](LICENSE). Use it, fork it, ship it commercially; no attribution required, though it is appreciated.

---

## Why this exists

Built in an afternoon because merging a folder of call transcripts into one document should not cost US$240 a year or require uploading client material to a website. Shared publicly in case it saves someone else the same annoyance.
