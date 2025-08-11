# Digital Signature Tool (Non-Certified)

A Windows digital signature generation and signing tool that helps you create self-signed certificates and sign executable files.

## Features

- **One-click certificate generation and signing**: Generate .pfx certificate files and sign programs with timestamps
- **Sign with existing certificates**: Use your existing .pfx files to sign programs
- **Timestamp management**: Add timestamps to signed files with multiple timestamp server options
- **Multiple file format support**: Supports .exe, .dll, .sys, .msi, .cab, .cat, .ocx, .ps1, .psm1, .psd1, .js, .vbs, .wsf
- **Batch processing**: Sign or verify multiple files at once
- **Signature verification**: Verify digital signatures with color-coded status:
  - 🟢 Green: Trusted signature (certified by authority)
  - 🟡 Yellow: Self-signed certificate (not certified)
  - 🔴 Red: Unsigned or invalid signature
- **Certificate management**: Generate .pfx and .cer certificate files

## Requirements

- Windows operating system
- Python 3.6 or higher
- Administrator privileges (recommended for signing operations)

## Required Tools

The following tools must be present in the `tools` folder:
- `cert2spc.exe` - Certificate to SPC converter
- `makecert.exe` - Certificate creation tool
- `pvk2pfx.exe` - Private key to PFX converter
- `signtool.exe` - Microsoft signing tool

These tools are part of the Windows SDK and can be downloaded from Microsoft's official website.

## Usage

1. Clone or download this repository
2. Ensure all required tools are in the `tools` folder
3. Run the script with Python:
   ```bash
   python gui.py
