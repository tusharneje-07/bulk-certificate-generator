# Bulk Certificate Generator

A professional desktop application for designing and generating personalized PDF certificates in bulk from Excel spreadsheets.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-orange)

---

## Features

- Modern dark-themed user interface
- Project-based workflow
- Live certificate preview
- Draw and resize text boxes visually
- Multiple configurable text fields
- Excel placeholder support (`{name}`, `{course}`, etc.)
- Custom fonts and colors
- PDF metadata support
- Batch PDF certificate generation
- Automatic project saving
- Multi-threaded generation
- Progress tracking and logging

---

## Screenshots

Add screenshots here.

---

## Requirements

- Python 3.10 or newer

---

## Installation

Clone the repository.

```bash
git clone https://github.com/<your-username>/bulk-certificate-generator.git
cd bulk-certificate-generator
```

Create a virtual environment.

```bash
python -m venv venv
```

Activate it.

### Linux / macOS

```bash
source venv/bin/activate
```

### Windows

```cmd
venv\Scripts\activate
```

Install dependencies.

```bash
pip install -r requirements.txt
```

---

## Running

### Linux

```bash
chmod +x run.sh
./run.sh
```

### Windows

```cmd
python main.py
```

---

## Project Structure

```text
bulk-certificate-generator/
├── main.py
├── run.sh
├── requirements.txt
├── README.md
├── LICENSE
├── .gitignore
├── assets/
├── projects/
└── screenshots/
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

This project is licensed under the MIT License.

---

## Author

Developed by **Aptronix / Your Name**