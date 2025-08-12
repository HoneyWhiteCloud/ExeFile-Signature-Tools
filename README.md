# Digital Signature Tool (Non-Certified)

A Windows tool to create self-signed certificates and sign executables.

This release ships a revamped UI that improves usability and reliability with multi-language support, drag-and-drop, batch processing, Unicode-safe paths, and no console popups.

## What’s New (UI Refresh)

- Auto language detection and localized title: switches among Simplified Chinese, Traditional Chinese, and English based on your system UI language
- High-DPI aware UI for crisp rendering on modern displays
- Drag-and-drop to add files (optional dependency: tkinterdnd2)
- No console flashes: signing/verification runs silently without spawning cmd windows
- GUI password prompt with cache: securely prompts for PFX passwords and remembers them per file, works even when packaged without a console
- Full Unicode support: can sign and timestamp files/paths containing CJK and full-width characters
- RFC 3161 timestamps first: uses `/tr` + `/td sha256`, automatically falls back to `/t`
- Batch and concurrency improvements:
  - “Verify” and “Sign (no timestamp)” run concurrently for speed
  - “Sign + timestamp” and “Timestamp only” run sequentially to avoid TSA rate limits
- Clear color-coded logs and summary statistics
- More robust icon loading when packaged with PyInstaller

## Features

- One-click certificate generation and signing: create `.pfx` and sign with a timestamp
- Sign with existing certificates: use your `.pfx` to sign programs
- Timestamp management: add timestamps to signed files with multiple servers (RFC 3161 preferred)
- Multiple file format support: `.exe`, `.dll`, `.sys`, `.msi`, `.cab`, `.cat`, `.ocx`, `.ps1`, `.psm1`, `.psd1`, `.js`, `.vbs`, `.wsf`
- Batch processing with concurrency
- Signature verification with color-coded status
  - 🟢 Trusted signature (CA-certified)
  - 🟡 Self-signed certificate (not CA-certified)
  - 🔴 Unsigned or invalid signature
- Certificate management: generate `.pfx` and `.cer` files

## Requirements

- Windows
- Python 3.10+ (end users of the packaged EXE do not need Python)
- Administrator privileges (recommended for signing)
- Optional: `tkinterdnd2` for drag-and-drop
  - Install: `pip install tkinterdnd2`

## Required Tools

The following tools must be present in the `tools` folder:
- `cert2spc.exe` — Certificate to SPC converter
- `makecert.exe` — Certificate creation tool
- `pvk2pfx.exe` — Private key to PFX converter
- `signtool.exe` — Microsoft signing tool

These tools are part of the Windows SDK and can be downloaded from Microsoft.

## Getting Started

- Run from source
  1. Clone or download the repository
  2. Ensure the `tools` folder contains the required executables listed above
  3. Optionally enable drag-and-drop: `pip install tkinterdnd2`
  4. Start the GUI:
     ```bash
     python gui.py
     ```

- Package for end users (single-file EXE)
  - Use the provided PyInstaller spec:
    ```bash
    pyinstaller tool.spec
    ```
  - Notes
    - `tool.spec` already embeds resources (e.g., `icon.ico`, `tools` directory)
    - `console=False` is recommended; this release handles PFX passwords via GUI and suppresses console windows

## How to Use

1. Add files
   - Click “Add Files” or drag and drop into the list (with `tkinterdnd2` installed)
2. Choose a certificate
   - Select an existing `.pfx` and enter the password (or leave blank and provide it on prompt), or
   - Click “Create Self-signed PFX…” to generate one (a “Create CER only” action is also available)
3. Pick a timestamp server
   - Choose from presets; signing prefers `/tr` (RFC 3161) and falls back to `/t` automatically
4. Run operations
   - Verify signatures: runs concurrently and shows color-coded logs
   - Sign + timestamp: runs sequentially to avoid TSA rate limits
   - Sign (no timestamp): runs concurrently for speed
   - Timestamp only: sequentially add timestamps to already-signed files
5. Inspect results
   - Logs include summary and key fields like signer, issuer, and timestamp

## Troubleshooting

- Failures with CJK/full-width characters in file names?
  - Fixed: the tool invokes `signtool` via Unicode-safe process calls without shell
- Console window flashes during signing/verification?
  - Fixed: child processes are started hidden (`CREATE_NO_WINDOW` + `SW_HIDE`)
- No console packaging breaks password prompts?
  - Fixed: a GUI prompt collects and caches PFX passwords
- Icon not shown or looks blurry?
  - Ensure `icon.ico` contains 16×16 and 32×32 (ideally also 48/64/256) sizes
- Drag-and-drop doesn’t work?
  - Install `tkinterdnd2` and restart: `pip install tkinterdnd2`
- Timestamp errors?
  - Could be TSA throttling or instability; the tool prefers `/tr` and falls back to `/t`. Try another TSA if needed.

## License

This project is intended for learning and internal tooling scenarios. Use responsibly and comply with your environment and certificate policies.
