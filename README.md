# Office Sensitive-Word Encryptor

**A fully offline Windows desktop tool that makes sensitive words invisible inside Office documents.**

It replaces sensitive terms in Word / Excel / PowerPoint files with random English-letter strings while keeping the original layout intact. Only the matching encrypted word list plus your password can restore the original text. Everything runs locally — no document, word list, or password ever leaves your machine.

- Current version: **v2.4**
- Supported formats: Word `.docx`, Excel `.xlsx`, PowerPoint `.pptx`
- UI: native Python / tkinter desktop application (English UI strings are localized to Chinese in the shipped build)
- License: free for personal, non-commercial use

---

## Highlights

- **Invisible substitution** — sensitive terms are replaced with random letter strings; formatting, tables and styling are preserved.
- **Two encryption modes**
  - *Standard*: the replacement string has the same length as the sensitive word (1:1 mapping).
  - *Strong encryption* (v2.4): for a word of length `L`, the replacement length is randomized in `[L, 2L]`, and every sensitive word gets **5–10 distinct** letter strings. Each occurrence in the document is replaced by a randomly chosen variant, so the ciphertext no longer has a fixed length or a 1-to-1 correspondence with the source word.
- **AES-256-GCM word-list encryption** with a random salt and PBKDF2 key derivation; the legacy v2.2 Fernet word lists can still be decrypted.
- **Document fingerprint binding** — the SHA-256 hash of each encrypted document is stored in the word list and verified on decrypt, preventing word-list mix-ups.
- **Zero-width marker** — a zero-width space (`U+200B`) is appended to substitutions so normal words are never touched and encryption is exactly reversible.
- **Broad content coverage**
  - Word: body, tables, headers/footers, hyperlink text and text boxes.
  - Excel: normal and rich-text cells; numeric / date / formula / boolean cells are type-protected (a replacement that cannot keep the numeric type is skipped with a warning instead of corrupting the cell).
  - PowerPoint: paragraph runs and table cells.
- **Single-file and batch processing** — encrypt or decrypt one document or a whole folder; batch mode supports one shared word list or an independent list per file.
- **Automatic decryption** — standard and strong-encryption word lists are both recognized automatically; no mode selection is needed.
- **Responsive GUI** — processing runs on background threads with a live log and overwrite confirmation.
- **Domain demo assets** — ready-made vocabularies and sample documents for **legal, finance, military and public-security** domains.

## Privacy

The application contains no network code whatsoever. It does not collect data, use cookies, embed analytics/advertising SDKs, download updates, or call any online API. All inputs and outputs stay in local folders you choose. See the “Privacy” section of the product page ([index.html](index.html)) for the full policy.

## Technology stack

| Area | Choice |
| --- | --- |
| Language / runtime | Python 3.9+ |
| GUI | tkinter (standard library) |
| Office parsing | [python-docx](https://pypi.org/project-python-docx/), [openpyxl](https://pypi.org/project/openpyxl/), [python-pptx](https://pypi.org/project/python-pptx/) |
| Cryptography | [cryptography](https://pypi.org/project/cryptography/) (AES-256-GCM, PBKDF2) |
| Packaging | PyInstaller → standalone `.exe` / MSIX / ZIP |

## Project structure

```
SedT/
├─ index.html                              # Product landing page (features, privacy, download)
├─ img/                                    # Landing-page screenshots and reward QR image
└─ app/
   ├─ office_sensitive_encryptor.py        # Main application (single-file source)
   ├─ office_sensitive_encryptor.spec      # PyInstaller spec
   ├─ test_encryptor.py                    # Regression tests (14 cases)
   ├─ icon.ico                             # Application icon
   ├─ build.bat                            # One-click build pipeline (8 steps)
   ├─ clean.bat                            # Remove generated build outputs
   ├─ PACKAGING_GUIDE.txt                  # Detailed packaging / Store / signing guide
   ├─ word/                                # Demo vocabularies + sample documents
   │  ├─ vocab_legal.json   (108 words)    vocab_finance.json  (108 words)
   │  ├─ vocab_military.json (110 words)   vocab_police.json   (108 words)
   │  └─ sample_legal.docx / sample_finance.docx /
   │     sample_military.docx / sample_police.docx
   └─ msix_packaging/
      ├─ build_msix.bat                    # Packs/signs the MSIX
      ├─ app/                             # MSIX content folder
      │  ├─ AppxManifest.xml              # Package identity / capabilities
      │  └─ Assets/                       # Store logos
      └─ zip/                             # Portable Program-Files distribution
         ├─ install.bat                    # Elevated install to C:\Program Files\...
         └─ uninstall.bat
```

## Run from source

Requirements: Windows 10/11 with Python 3.9+.

```powershell
cd app
python -m pip install python-docx openpyxl python-pptx cryptography
python office_sensitive_encryptor.py
```

The “Load word list” and “Save word list” dialogs default to the `word` subdirectory next to the program (creating it if missing).

### Word-list file format

Plain JSON vocabulary files use a simple `words` array; the encrypted list produced by the application uses an `.enc` container:

```json
{
  "words": ["案件当事人", "取保候审"]
}
```

### Command-line flags

The packaged executable also accepts:

- `--gen-files` — release the embedded user manual and license text files next to the executable and exit (used by the build).
- `--uninstall` — remove the desktop shortcut and Explorer right-menu registration entries and exit (used by the ZIP uninstaller).

## Test

```powershell
cd app
python test_encryptor.py
```

The suite (14 cases) covers crypto round-trips, legacy v2.2 compatibility, Word/Excel/PowerPoint encryption+decryption, type protection, SHA-256 fingerprint binding, batch flows, and the strong-encryption rules (5–10 variants per word, `[L, 2L]` length, uniqueness, randomized per-occurrence replacement).

## Build distribution packages

All commands run from the `app` folder on Windows:

```powershell
# 1) Unsigned MSIX for Microsoft Store submission (Store re-signs it)
build.bat

# 2) Signed MSIX for side-loading (needs msix_packaging\mycert.pfx)
build.bat sign PASSWORD your-pfx-password

# 3) Remove generated artifacts (exe, copied word/, txt, MSIX, ZIP)
clean.bat
```

`build.bat` runs an 8-step pipeline: check Python → install dependencies → run the self-test → build the exe with PyInstaller → copy the `word/` assets → generate manual/license files → pack the MSIX → build the ZIP.

Outputs:

- `app/msix_packaging/dist_msix/OfficeSensitiveEncryptor.msix`
- `app/dist_zip/OfficeSensitiveEncryptor_<version>.zip`

The MSIX package identity is **`SamLi.Office`** (publisher `CN=DF81B408-E5B2-48DC-A919-4D0ABCEED8B5`), matching the reserved identity in Microsoft Partner Center. See [PACKAGING_GUIDE.txt](app/PACKAGING_GUIDE.txt) for signing, certificate and Store-submission details.

## Installation options

1. **Microsoft Store** — install the published MSIX; Windows manages updates and signing.
2. **MSIX side-loading** — trust the self-signed certificate (or enable Developer Mode), then `Add-AppxPackage`.
3. **Portable ZIP** — unzip `OfficeSensitiveEncryptor_<version>.zip`, right-click `install.bat` → *Run as administrator*. It installs to `C:\Program Files\OfficeSensitiveEncryptor`, copies the `word/` assets, and creates a desktop shortcut and Explorer context-menu entries; `uninstall.bat` in the install folder removes everything.

## Important notes

- Keep your password and encrypted word list (`.enc`) safe. A lost password cannot be recovered — the author provides no recovery service.
- Strong-encryption replacement strings are letters only; Excel cells that must remain numeric/dates are skipped rather than forced to text.
- You are responsible for distributing encrypted documents in compliance with applicable laws.

## Version history

- **v2.4** — strong-encryption mode (variable-length, multi-variant substitution); word-list format version `2.4`; auto-detect decrypt; bundled legal/finance/military/police vocabularies and samples.
- **v2.3** — AES-256-GCM word lists (with v2.2 Fernet read compatibility), SHA-256 document binding, expanded Word coverage, Excel type protection, background processing, batch mode.
- **v2.2** — initial equal-length substitution release.
