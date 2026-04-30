# 📥 Google Drive Bulk Downloader from Spreadsheet

> Automates bulk file downloads from Google Drive hyperlinks embedded in Excel spreadsheets — built to solve a real workflow problem in national vocational training program administration.

---

## 🔍 Background

In the context of national vocational training program administration at BBPVP Bekasi, participant data — including photos, ID card scans, and diploma images — was collected via Google Forms and stored as Google Drive links inside a shared Excel spreadsheet.

The manual process: open the spreadsheet → click each hyperlink → download the file → rename it → organize it by category. Repeated for every participant, every training batch.

**This script eliminates that entirely.**

It reads the spreadsheet, extracts all hidden hyperlinks, converts them to direct download URLs, and organizes the files into structured folders automatically — used as part of the certificate validation workflow to verify participant data before issuing official training certificates.

---

## ⚙️ What It Does

- Loads an `.xlsx` file and extracts **hidden hyperlinks** from cells (separate from cell display values)
- Handles **multiple Google Drive URL formats** and converts them to direct download URLs
- Manages Google Drive's **virus-scan confirmation page** automatically
- Downloads files using **streaming** to handle large files without memory issues
- Creates an **organized folder structure** by file category (FOTO, FOTO_IJAZAH, LINK_KTP)
- Handles **duplicate filenames** gracefully with auto-incrementing suffixes
- Full **CLI interface** — configurable file path, output folder, delay, timeout, and error handling
- Prints a clean **summary report** at the end

---

## 📁 Output Structure

```
downloads/
├── FOTO/
│   ├── nama_peserta_1.jpg
│   └── nama_peserta_2.jpg
├── FOTO_IJAZAH/
│   ├── nama_peserta_1.jpg
│   └── nama_peserta_2.jpg
└── LINK_KTP/
    ├── nama_peserta_1.jpg
    └── nama_peserta_2.jpg
```

---

## 🚀 Usage

### Basic
```bash
python download_from_spreadsheet.py
```

### With options
```bash
python download_from_spreadsheet.py --file copy_data.xlsx --output ./hasil
python download_from_spreadsheet.py --file data.xlsx --delay 2.0 --timeout 90 --skip-errors
```

### All flags
| Flag | Default | Description |
|------|---------|-------------|
| `--file` / `-f` | `copy data.xlsx` | Path to the `.xlsx` file |
| `--output` / `-o` | `./downloads` | Root output folder |
| `--delay` | `1.0` | Seconds between downloads (avoid rate limiting) |
| `--timeout` | `60` | HTTP timeout in seconds |
| `--skip-errors` | `False` | Continue even if some downloads fail |

---

## 📦 Requirements

```bash
pip install openpyxl requests
```

---

## 🛠️ Configuration

Edit these constants at the top of the script to match your spreadsheet layout:

```python
COLUMN_MAP = {
    "FOTO":        11,   # Column K
    "FOTO_IJAZAH": 12,   # Column L
    "LINK_KTP":    13,   # Column M
}

NAME_COLUMN  = 3   # Column C = participant full name
HEADER_ROW   = 2   # Row with column headers
DATA_START   = 3   # First row of actual data
```

---

## 💡 Key Technical Notes

- **Hidden hyperlinks**: Excel stores hyperlinks separately from cell display values. This script reads `cell.hyperlink.target` directly — not just `cell.value`.
- **Google Drive URL normalization**: Handles `/file/d/{ID}/view`, `?id={ID}`, and other formats — all converted to `uc?export=download&id=` format.
- **Confirmation page handling**: Large Google Drive files trigger a virus-scan warning page. The script extracts the confirmation token and retries automatically, with a fallback `confirm=t` approach.
- **Streaming download**: Uses `stream=True` with chunked iteration so large files don't load entirely into memory.

---

## 📊 Context

Built as part of a training program data management workflow at **BBPVP Bekasi** (Balai Besar Pelatihan Vokasi dan Produktivitas), where participant documentation needed to be verified and archived before certificate issuance across multiple concurrent training batches.

The script reduced manual download time from ~45 minutes per batch to near-zero, and eliminated human error in file naming and organization.

---

*Built by Mohammad Syaoqi Ramadhan — [linkedin.com/in/msyaoqiramadhan](https://linkedin.com/in/msyaoqiramadhan)*
