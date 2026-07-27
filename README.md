# Bulk Certificate Generator

A professional desktop application for designing, previewing, and generating personalized PDF certificates in bulk from Excel spreadsheets.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-orange)

---

## Features

- Modern dark-themed user interface
- Project-based workflow
- Interactive live certificate preview
- Visual drag-and-drop text box editor
- Draw, move, and resize text boxes directly on the certificate
- Multiple independently configurable text fields
- Excel variable placeholders (`{name}`, `{course}`, `{certificate_number}`, etc.)
- Custom fonts, colors, sizes, and text alignment
- PDF metadata customization
- Automatic project management
- Multi-threaded batch PDF certificate generation
- Progress tracking with detailed logs
- SQLite-backed project database
- Cross-platform support (Windows, Linux, and macOS)

---

## Requirements

- Python **3.10** or later

---

## Installation

Clone the repository.

```bash
git clone https://github.com/tusharneje-07/bulk-certificate-generator.git
cd bulk-certificate-generator
```

---

## Running

### Linux

```bash
chmod +x linux_run.sh
./linux_run.sh
```

### macOS

```bash
chmod +x mac_run.sh
./mac_run.sh
```

### Windows (Batch)

```cmd
windows_run.bat
```

### Windows (PowerShell)

```powershell
.\windows_run.ps1
```

The launcher will automatically:

- Create a virtual environment (if needed)
- Upgrade `pip`
- Detect missing Python packages
- Install only the required dependencies
- Launch the application

No manual dependency installation is required.

---

## Project Structure

```text
bulk-certificate-generator/
├── main.py
├── linux_run.sh
├── mac_run.sh
├── windows_run.bat
├── windows_run.ps1
├── requirements.txt
├── README.md
├── LICENSE
├── .gitignore
├── assets/
├── projects/
```

---

## Technologies Used

- Python
- Tkinter
- ttkbootstrap
- Pillow
- OpenPyXL
- pypdf
- SQLite

---

## License

This project is licensed under the **MIT License**.

---


> Made with ❤️ by [@tusharneje-07](https://github.com/tusharneje-07/) and Claude
