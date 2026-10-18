"""
Certificate Generator Studio
=============================

A production-style desktop application (Tkinter + ttkbootstrap, dark theme) for
bulk-generating PDF certificates from a template image and an Excel roster.

Highlights
----------
- Project-based workflow: every certificate job lives in its own project under
  ./projects/<project>/ (self-contained: template image, font, and Excel roster
  are copied in). Projects are tracked in a local SQLite database and each
  project also keeps a human-readable ".projectinfo.json" with full details.
- Multiple, independently configurable text boxes on the certificate, each
  with its own template text (supporting {variable} placeholders pulled from
  Excel columns), font, size, color and alignment.
- A dedicated "Certificate Number" column selector exposes a {certificate_number}
  variable usable anywhere - in any text box, or in the PDF metadata - to
  authorize / identify each certificate.
- Proper PDF metadata (Title, Author, Subject, Keywords, Creator, Producer,
  plus arbitrary custom fields), all of which can also use {variable} templates.
- An interactive, zoomable canvas preview: draw a rectangle directly on the
  certificate to create a new text box, drag to move it, drag its corner to
  resize it. Each box is drawn in its own color.
- Multithreaded batch generation with progress, a collapsible log, and cancel.

Install:
    pip install ttkbootstrap pillow openpyxl pypdf

Run:
    python main.py
"""

import io
import os
import re
import json
import queue
import random
import shutil
import sqlite3
import platform
import subprocess
import threading
import traceback
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, colorchooser

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from PIL import Image, ImageDraw, ImageFont, ImageTk

import openpyxl
from pypdf import PdfReader, PdfWriter


# ============================================================================ #
# Constants
# ============================================================================ #

APP_DIR = Path(__file__).resolve().parent
PROJECTS_ROOT = APP_DIR / "projects"
DB_PATH = PROJECTS_ROOT / "projects.db"
PROJECT_INFO_FILENAME = ".projectinfo.json"

BOX_COLOR_PALETTE = [
    "#ff5252", "#448aff", "#69f0ae", "#ffab40", "#e040fb",
    "#40c4ff", "#eeff41", "#ff6e40", "#b388ff", "#64ffda",
]

PREVIEW_MIN_ZOOM = 0.1
PREVIEW_MAX_ZOOM = 4.0

ALIGN_OPTIONS = ["center", "left", "right"]

METADATA_STANDARD_FIELDS = [
    ("title", "Title"),
    ("author", "Author"),
    ("subject", "Subject"),
    ("keywords", "Keywords"),
    ("creator", "Creator"),
    ("producer", "Producer"),
]

COLORS = {
    "bg": "#191919",
    "sidebar": "#202020",
    "surface": "#242424",
    "surface_2": "#2b2b2b",
    "surface_hover": "#303030",
    "border": "#383838",
    "text": "#f1f1ef",
    "muted": "#9b9b98",
    "accent": "#f1f1ef",
    "accent_hover": "#ffffff",
    "selection": "#414141",
}

UI_FONT = "Segoe UI"
DISPLAY_FONT = "Segoe UI"


def studio_mark(size):
    """Minimal folded-sheet monogram with a clear silhouette at icon sizes."""
    scale = 8
    image = Image.new("RGBA", (size * scale, size * scale))
    draw = ImageDraw.Draw(image)
    def polygon(points, fill):
        draw.polygon([(round(x * size * scale), round(y * size * scale))
                      for x, y in points], fill=fill)

    # One open, architectural C. The separated fold adds depth without tiny
    # decorative details, outlines, lettering, or a surrounding app tile.
    polygon(((.18, .12), (.62, .12), (.62, .31), (.37, .31),
             (.37, .69), (.82, .69), (.82, .88), (.18, .88)), "#f1f1ef")
    polygon(((.68, .12), (.86, .31), (.68, .31)), "#a6a6a4")
    return image.resize((size, size), Image.Resampling.LANCZOS)


def display_project_date(value):
    try:
        return datetime.fromisoformat(value).strftime("%d %b %Y  ·  %H:%M")
    except (ValueError, TypeError):
        return str(value or "—")


# ============================================================================ #
# Generic helpers
# ============================================================================ #

def sanitize_filename(name):
    """Turn an arbitrary string into a filesystem-safe name (no extension)."""
    name = str(name)
    safe = "".join(c for c in name if c.isalnum() or c in (" ", "-", "_")).strip()
    safe = safe.replace(" ", "_")
    return safe or "untitled"


def make_token(header, existing_tokens):
    """Turn an Excel header into a safe, unique {variable} token."""
    token = re.sub(r"[^0-9a-zA-Z]+", "_", str(header).strip().lower()).strip("_")
    if not token:
        token = "col"
    if token[0].isdigit():
        token = "c_" + token
    base = token
    i = 2
    while token in existing_tokens:
        token = f"{base}_{i}"
        i += 1
    return token


class SafeDict(dict):
    """A dict that leaves unknown {tokens} visible instead of raising, so typos
    in a template are obvious in the preview rather than crashing generation."""

    def __missing__(self, key):
        return "{" + key + "}"


def render_template(template_str, variables):
    if not template_str:
        return ""
    try:
        return str(template_str).format_map(SafeDict(variables))
    except Exception:
        return str(template_str)


def load_font(font_path, font_size):
    if font_path:
        try:
            return ImageFont.truetype(str(font_path), int(font_size))
        except Exception:
            pass
    return ImageFont.load_default()


def open_template(template_source):
    """template_source may be a filesystem path/Path or raw bytes."""
    if isinstance(template_source, (bytes, bytearray)):
        return Image.open(io.BytesIO(template_source))
    return Image.open(template_source)


def build_certificate_image(template_source, text_boxes, variables, default_font_path):
    """Draw every text box's rendered template text onto a fresh copy of the
    template. This is the single source of truth used by both the live
    preview and the final PDF export.

    text_boxes: iterable of dicts with keys:
        template, left, top, width, height, font_file (path or None),
        font_size, color, align
    """
    template = open_template(template_source)
    cert = template.convert("RGB")
    draw = ImageDraw.Draw(cert)

    for box in text_boxes:
        text = render_template(box.get("template", ""), variables)
        if not text:
            continue
        font_path = box.get("font_file") or default_font_path
        font = load_font(font_path, box.get("font_size", 60))
        left = box.get("left", 0)
        top = box.get("top", 0)
        width = box.get("width", 100)
        height = box.get("height", 100)
        color = box.get("color", "#000000")
        align = box.get("align", "center")

        bbox = draw.textbbox((0, 0), text, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        center_y = top + height / 2
        y = center_y - text_h / 2 - bbox[1]

        if align == "left":
            x = left - bbox[0]
        elif align == "right":
            x = left + width - text_w - bbox[0]
        else:
            x = left + width / 2 - text_w / 2 - bbox[0]

        draw.text((x, y), text, fill=color, font=font)

    return cert


def build_metadata_dict(meta_config, variables):
    """Render the configured metadata templates into a pypdf-ready dict with
    '/'-prefixed keys."""
    d = {}
    key_map = {k: f"/{label}" for k, label in METADATA_STANDARD_FIELDS}
    for key, pdf_key in key_map.items():
        val = meta_config.get(key, "")
        if val:
            d[pdf_key] = render_template(val, variables)
    for k, v in (meta_config.get("custom") or {}).items():
        if not k:
            continue
        pdf_key = k if k.startswith("/") else "/" + k
        d[pdf_key] = render_template(v, variables)
    return d


def save_certificate_pdf(cert_image, resolution, metadata_dict, out_path):
    """Save a PIL image as a single-page PDF and stamp it with metadata."""
    buf = io.BytesIO()
    cert_image.save(buf, "PDF", resolution=float(resolution))
    buf.seek(0)
    reader = PdfReader(buf)
    writer = PdfWriter()
    writer.append_pages_from_reader(reader)
    if metadata_dict:
        writer.add_metadata(metadata_dict)
    with open(out_path, "wb") as f:
        writer.write(f)


def open_in_file_manager(path):
    path = str(path)
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(path)  # noqa
        elif system == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# ============================================================================ #
# Project data model + manager (SQLite registry + per-project JSON)
# ============================================================================ #

def default_project_info(name, now):
    return {
        "name": name,
        "created_at": now,
        "updated_at": now,
        "template_image": None,
        "excel_file": None,
        "sheet_name": None,
        "primary_col_header": None,
        "certificate_number_col_header": None,
        "filename_col_header": None,
        "group_col_header": None,
        "default_font_file": None,
        "pdf_resolution": 100,
        "text_boxes": [],
        "metadata": {
            "title": "",
            "author": "",
            "subject": "",
            "keywords": "",
            "creator": "Certificate Generator Studio",
            "producer": "Certificate Generator Studio",
            "custom": {},
        },
    }


class ProjectManager:
    def __init__(self):
        PROJECTS_ROOT.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(DB_PATH))
        self.conn.execute(
            """CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                folder TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )"""
        )
        self.conn.commit()

    def list_projects(self):
        cur = self.conn.execute(
            "SELECT id, name, folder, created_at, updated_at FROM projects ORDER BY updated_at DESC"
        )
        return [
            dict(id=r[0], name=r[1], folder=r[2], created_at=r[3], updated_at=r[4])
            for r in cur.fetchall()
        ]

    def _unique_folder_name(self, base):
        candidate = base
        i = 2
        while (PROJECTS_ROOT / candidate).exists():
            candidate = f"{base}_{i}"
            i += 1
        return candidate

    def create_project(self, name):
        name = name.strip()
        if not name:
            raise ValueError("Project name cannot be empty.")
        exists = self.conn.execute("SELECT 1 FROM projects WHERE name = ?", (name,)).fetchone()
        if exists:
            raise ValueError(f"A project named '{name}' already exists.")
        folder_name = self._unique_folder_name(sanitize_filename(name) or "project")
        folder = PROJECTS_ROOT / folder_name
        (folder / "assets" / "fonts").mkdir(parents=True, exist_ok=True)
        (folder / "output").mkdir(parents=True, exist_ok=True)
        now = datetime.now().isoformat(timespec="seconds")
        info = default_project_info(name, now)
        with open(folder / PROJECT_INFO_FILENAME, "w", encoding="utf-8") as f:
            json.dump(info, f, indent=2)
        self.conn.execute(
            "INSERT INTO projects (name, folder, created_at, updated_at) VALUES (?,?,?,?)",
            (name, folder_name, now, now),
        )
        self.conn.commit()
        return dict(name=name, folder=folder_name, created_at=now, updated_at=now)

    def touch(self, name):
        now = datetime.now().isoformat(timespec="seconds")
        self.conn.execute("UPDATE projects SET updated_at=? WHERE name=?", (now, name))
        self.conn.commit()

    def delete_project(self, project_id, folder_name, delete_files=True):
        self.conn.execute("DELETE FROM projects WHERE id=?", (project_id,))
        self.conn.commit()
        if delete_files:
            shutil.rmtree(PROJECTS_ROOT / folder_name, ignore_errors=True)

    def project_path(self, folder_name):
        return PROJECTS_ROOT / folder_name

    def load_info(self, folder_name):
        path = self.project_path(folder_name) / PROJECT_INFO_FILENAME
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    def save_info(self, folder_name, info, project_name):
        info["updated_at"] = datetime.now().isoformat(timespec="seconds")
        path = self.project_path(folder_name) / PROJECT_INFO_FILENAME
        with open(path, "w", encoding="utf-8") as f:
            json.dump(info, f, indent=2)
        self.touch(project_name)


def copy_asset_into(project_folder, src_path, subfolder, desired_stem=None):
    """Copy an external file into the project folder; returns a project-relative
    posix path string. Overwrites any previous file with the same destination
    name (this is the "proper updation" behaviour when a user re-picks a file)."""
    src_path = Path(src_path)
    ext = src_path.suffix
    dest_dir = Path(project_folder) / subfolder
    dest_dir.mkdir(parents=True, exist_ok=True)
    stem = sanitize_filename(desired_stem or src_path.stem)
    dest_path = dest_dir / f"{stem}{ext}"
    shutil.copy2(src_path, dest_path)
    return (Path(subfolder) / f"{stem}{ext}").as_posix()


# ============================================================================ #
# Small reusable widgets
# ============================================================================ #

class RoundedButton(tk.Canvas):
    """Canvas-backed button with genuinely rounded edges and quiet hover states."""

    def __init__(self, master, text, command=None, kind="secondary", width=None, height=38, **kwargs):
        self.command = command
        self.kind = kind
        self._state = kwargs.pop("state", "normal")
        self._font = kwargs.pop("font", (UI_FONT, 10))
        canvas_bg = kwargs.pop("canvas_bg", self._parent_bg(master))
        self._text = text
        measured = tkfont.Font(font=self._font).measure(text)
        self._button_width = width or max(82, measured + 34)
        self._button_height = height
        super().__init__(
            master, width=self._button_width, height=height, highlightthickness=0,
            borderwidth=0, relief="flat", bg=canvas_bg, cursor="hand2",
        )
        self.bind("<Enter>", lambda _e: self._draw(True))
        self.bind("<Leave>", lambda _e: self._draw(False))
        self.bind("<Button-1>", self._click)
        self.bind("<Configure>", self._resize)
        self.configure(takefocus=True)
        self.bind("<Return>", self._click)
        self.bind("<space>", self._click)
        self.bind("<FocusIn>", lambda _e: self._draw(True))
        self.bind("<FocusOut>", lambda _e: self._draw(False))
        self._draw(False)

    def _resize(self, event):
        self._button_width, self._button_height = event.width, event.height
        self._draw(False)

    @staticmethod
    def _parent_bg(master):
        try:
            return master.cget("background")
        except Exception:
            style = master.cget("style") or "TFrame"
            return tb.Style().lookup(style, "background") or COLORS["sidebar"]

    def _palette(self, hovered=False):
        if self._state == "disabled":
            return "#303030", "#737373", "#383838"
        if self.kind == "primary":
            return ("#ffffff" if hovered else "#f1f1ef"), "#171717", "#f1f1ef"
        if self.kind == "danger":
            return ("#36292b" if hovered else "#292526"), "#df9a9f", "#493438"
        return (COLORS["surface_hover"] if hovered else COLORS["surface"]), COLORS["text"], COLORS["border"]

    def _rounded_rect(self, x1, y1, x2, y2, radius, **options):
        points = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        return self.create_polygon(points, smooth=True, splinesteps=24, **options)

    def _draw(self, hovered):
        self.delete("all")
        fill, text_color, outline = self._palette(hovered)
        self._rounded_rect(1, 1, self._button_width - 1, self._button_height - 1, 9,
                           fill=fill, outline=outline, width=1)
        self.create_text(self._button_width / 2, self._button_height / 2,
                         text=self._text, fill=text_color, font=self._font)

    def _click(self, _event):
        if self._state != "disabled" and self.command:
            self.command()

    def configure(self, cnf=None, **kwargs):
        if cnf:
            kwargs.update(cnf)
        if "state" in kwargs:
            self._state = kwargs.pop("state")
        if "text" in kwargs:
            self._text = kwargs.pop("text")
        if kwargs:
            super().configure(**kwargs)
        self._draw(False)

    config = configure


class ProjectNameDialog(tk.Toplevel):
    """Small first-party dialog instead of the generic ttkbootstrap prompt."""

    def __init__(self, parent):
        super().__init__(parent)
        self.result = None
        self.title("New project")
        self.configure(bg=COLORS["surface"])
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())
        self.grab_set()
        shell = tk.Frame(self, bg=COLORS["surface"], padx=26, pady=24)
        shell.pack(fill="both", expand=True)
        tk.Label(shell, text="Create a new project", bg=COLORS["surface"], fg=COLORS["text"],
                 font=(DISPLAY_FONT, 15, "bold")).pack(anchor="w")
        tk.Label(shell, text="Give this certificate workspace a name.", bg=COLORS["surface"], fg=COLORS["muted"],
                 font=(UI_FONT, 10)).pack(anchor="w", pady=(5, 18))
        self.entry = tk.Entry(shell, bg=COLORS["surface_2"], fg=COLORS["text"], insertbackground=COLORS["text"],
                              relief="flat", highlightthickness=1, highlightbackground=COLORS["border"],
                              highlightcolor="#707070", font=(UI_FONT, 11))
        self.entry.pack(fill="x", ipady=9)
        actions = tk.Frame(shell, bg=COLORS["surface"])
        actions.pack(fill="x", pady=(22, 0))
        RoundedButton(actions, "Create project", self._submit, kind="primary", width=132).pack(side="right")
        RoundedButton(actions, "Cancel", self.destroy, width=84).pack(side="right", padx=(0, 9))
        self.entry.bind("<Return>", lambda _e: self._submit())
        self.bind("<Escape>", lambda _e: self.destroy())
        self.update_idletasks()
        x = parent.winfo_toplevel().winfo_rootx() + (parent.winfo_toplevel().winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_toplevel().winfo_rooty() + (parent.winfo_toplevel().winfo_height() - self.winfo_height()) // 2
        self.geometry(f"360x210+{x}+{y}")
        self.after(60, self.entry.focus_set)

    def _submit(self):
        value = self.entry.get().strip()
        if value:
            self.result = value
            self.destroy()


class ScrollableFrame(tb.Frame):
    """A Frame with a vertical scrollbar for the left-hand control panel."""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0, bg=COLORS["sidebar"])
        self.vbar = tb.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = tb.Frame(self.canvas)

        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self._window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self._window, width=e.width))
        self.canvas.configure(yscrollcommand=self.vbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.vbar.pack(side="right", fill="y")

        for widget in (self.canvas, self.inner):
            widget.bind("<Enter>", self._bind_mousewheel)
            widget.bind("<Leave>", self._unbind_mousewheel)

    def _bind_mousewheel(self, _e):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _unbind_mousewheel(self, _e):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-3, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(3, "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120) * 3), "units")


class Collapsible(tb.Frame):
    """A section that can be expanded/collapsed by clicking its header - used
    for the technical log, which most non-technical users can ignore."""

    def __init__(self, master, title, collapsed=True, **kwargs):
        super().__init__(master, **kwargs)
        self.columnconfigure(0, weight=1)
        self._collapsed = collapsed

        header = tb.Frame(self)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)

        self.toggle_btn = RoundedButton(
            header, text=self._arrow(), width=32, height=30, command=self.toggle
        )
        self.toggle_btn.grid(row=0, column=0)
        tb.Label(header, text=title, bootstyle=SECONDARY).grid(row=0, column=1, sticky="w")

        self.body = tb.Frame(self)
        self.rowconfigure(1, weight=1 if not collapsed else 0)
        if not collapsed:
            self.body.grid(row=1, column=0, sticky="nsew", pady=(4, 0))

    def _arrow(self):
        return "▸" if self._collapsed else "▾"

    def toggle(self):
        self._collapsed = not self._collapsed
        self.toggle_btn.configure(text=self._arrow())
        if self._collapsed:
            self.body.grid_remove()
            self.rowconfigure(1, weight=0)
        else:
            self.body.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
            self.rowconfigure(1, weight=1)


# ============================================================================ #
# Interactive, zoomable preview canvas
# ============================================================================ #

class InteractiveCanvasPreview(tb.Frame):
    """Shows the rendered certificate and lets the user draw / move / resize
    text-box rectangles directly on top of it. Rectangles are lightweight
    (id/left/top/width/height/color/label) - the authoritative text-box data
    (template text, font, etc.) is owned by the caller."""

    HANDLE_PX = 10  # on-screen size of the resize handle, independent of zoom

    def __init__(self, master, on_box_change=None, on_new_box=None, on_select=None, **kwargs):
        super().__init__(master, **kwargs)
        self.on_box_change = on_box_change
        self.on_new_box = on_new_box
        self.on_select = on_select

        self.zoom = 1.0
        self.base_image_size = (0, 0)  # real template pixel size
        self._full_res_image = None
        self.base_photo = None
        self.base_image_id = None
        self.boxes = []          # [{id,left,top,width,height,color,label}, ...]
        self.selected_id = None
        self._box_items = {}     # id -> {'rect':..., 'chip':..., 'text':..., 'handle':...}

        self._drag_mode = None
        self._drag_box_id = None
        self._drag_start_img = (0, 0)
        self._drag_orig = None
        self._temp_rect = None

        self._build_ui()

    # -- construction ----------------------------------------------------
    def _build_ui(self):
        toolbar = tb.Frame(self, padding=(0, 8), style="App.TFrame")
        toolbar.grid(row=0, column=0, columnspan=2, sticky="ew")
        RoundedButton(toolbar, text="−", width=34, command=self.zoom_out).pack(side="left")
        self.zoom_label = tb.Label(toolbar, text="100%", width=6, anchor="center", style="Muted.TLabel")
        self.zoom_label.pack(side="left", padx=4)
        RoundedButton(toolbar, text="+", width=34, command=self.zoom_in).pack(side="left")
        RoundedButton(toolbar, text="Fit", command=self.fit).pack(side="left", padx=(8, 0))
        RoundedButton(toolbar, text="100%", command=self.reset_zoom).pack(side="left", padx=(4, 0))
        tb.Label(
            toolbar, text="Certificate preview", style="Muted.TLabel",
        ).pack(side="right", padx=(12, 0))

        self.canvas = tk.Canvas(self, bg="#171717", highlightthickness=1,
                                highlightbackground=COLORS["border"], bd=0)
        self.vbar = tb.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.hbar = tb.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.vbar.set, xscrollcommand=self.hbar.set)

        self.canvas.grid(row=1, column=0, sticky="nsew")
        self.vbar.grid(row=1, column=1, sticky="ns")
        self.hbar.grid(row=2, column=0, sticky="ew")

        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_motion)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Control-MouseWheel>", self._on_ctrl_wheel)
        self.canvas.bind("<Control-Button-4>", lambda e: self._zoom_step(1.15))
        self.canvas.bind("<Control-Button-5>", lambda e: self._zoom_step(1 / 1.15))
        self.canvas.bind("<Shift-MouseWheel>", lambda e: self.canvas.xview_scroll(int(-1 * (e.delta / 120) * 3), "units"))
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1 * (e.delta / 120) * 3), "units"))
        self.canvas.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-3, "units"))
        self.canvas.bind("<Button-5>", lambda e: self.canvas.yview_scroll(3, "units"))

    # -- public API --------------------------------------------------------
    def set_base_image(self, pil_image):
        """pil_image is the fully rendered certificate at real resolution."""
        self.base_image_size = pil_image.size
        self._full_res_image = pil_image
        self._redraw_all()

    def clear(self):
        self.canvas.delete("all")
        self._full_res_image = None
        self.base_image_size = (0, 0)
        self.boxes = []
        self._box_items = {}

    def set_boxes(self, boxes, selected_id=None):
        """boxes: list of dicts with id/left/top/width/height/color/label."""
        self.boxes = [dict(b) for b in boxes]
        self.selected_id = selected_id
        self._redraw_boxes()

    def fit(self):
        if self.base_image_size == (0, 0):
            return
        self.canvas.update_idletasks()
        pane_w = max(self.canvas.winfo_width(), 100)
        pane_h = max(self.canvas.winfo_height(), 100)
        img_w, img_h = self.base_image_size
        self.zoom = max(PREVIEW_MIN_ZOOM, min(pane_w / img_w, pane_h / img_h, PREVIEW_MAX_ZOOM))
        self._redraw_all()

    def reset_zoom(self):
        self.zoom = 1.0
        self._redraw_all()

    def zoom_in(self):
        self._zoom_step(1.25)

    def zoom_out(self):
        self._zoom_step(1 / 1.25)

    def _zoom_step(self, factor):
        self.zoom = max(PREVIEW_MIN_ZOOM, min(self.zoom * factor, PREVIEW_MAX_ZOOM))
        self._redraw_all()

    def _on_ctrl_wheel(self, event):
        self._zoom_step(1.15 if event.delta > 0 else 1 / 1.15)

    # -- drawing -----------------------------------------------------------
    def _redraw_all(self):
        self.canvas.delete("all")
        self._box_items = {}
        self.zoom_label.configure(text=f"{int(round(self.zoom * 100))}%")
        if self._full_res_image is None:
            return
        w, h = self.base_image_size
        disp_w, disp_h = max(1, int(w * self.zoom)), max(1, int(h * self.zoom))
        resized = self._full_res_image.resize((disp_w, disp_h), Image.LANCZOS)
        self.base_photo = ImageTk.PhotoImage(resized)
        self.base_image_id = self.canvas.create_image(0, 0, anchor="nw", image=self.base_photo)
        self.canvas.configure(scrollregion=(0, 0, disp_w, disp_h))
        self._redraw_boxes()

    def _redraw_boxes(self):
        for items in self._box_items.values():
            for item_id in items.values():
                if item_id:
                    self.canvas.delete(item_id)
        self._box_items = {}
        for box in self.boxes:
            self._draw_one_box(box)

    def _draw_one_box(self, box):
        z = self.zoom
        l, t, w, h = box["left"] * z, box["top"] * z, box["width"] * z, box["height"] * z
        color = box.get("color", "#ff5252")
        selected = box["id"] == self.selected_id
        width = 3 if selected else 2
        rect = self.canvas.create_rectangle(l, t, l + w, t + h, outline=color, width=width)

        label = box.get("label", "")
        chip = txt = None
        if label:
            chip_y = t - 16 if t > 18 else t + 2
            chip = self.canvas.create_rectangle(l, chip_y, l + 10 + 7 * len(label), chip_y + 15, fill=color, outline="")
            txt = self.canvas.create_text(l + 4, chip_y + 7, text=label, anchor="w", fill="#1b1b1b", font=("Segoe UI", 8, "bold"))

        handle = None
        if selected:
            hs = self.HANDLE_PX
            handle = self.canvas.create_rectangle(
                l + w - hs / 2, t + h - hs / 2, l + w + hs / 2, t + h + hs / 2,
                fill=color, outline="#ffffff",
            )

        self._box_items[box["id"]] = {"rect": rect, "chip": chip, "text": txt, "handle": handle}

    def _update_box_visual(self, box_id):
        box = self._get_box(box_id)
        if not box:
            return
        items = self._box_items.get(box_id)
        if not items:
            self._draw_one_box(box)
            return
        z = self.zoom
        l, t, w, h = box["left"] * z, box["top"] * z, box["width"] * z, box["height"] * z
        self.canvas.coords(items["rect"], l, t, l + w, t + h)
        if items["chip"]:
            label = box.get("label", "")
            chip_y = t - 16 if t > 18 else t + 2
            self.canvas.coords(items["chip"], l, chip_y, l + 10 + 7 * len(label), chip_y + 15)
            self.canvas.coords(items["text"], l + 4, chip_y + 7)
        if items["handle"]:
            hs = self.HANDLE_PX
            self.canvas.coords(items["handle"], l + w - hs / 2, t + h - hs / 2, l + w + hs / 2, t + h + hs / 2)

    def _get_box(self, box_id):
        for b in self.boxes:
            if b["id"] == box_id:
                return b
        return None

    # -- hit testing & interaction ------------------------------------------
    def _hit_test(self, img_x, img_y):
        handle_zone = self.HANDLE_PX / max(self.zoom, 0.001)
        for box in reversed(self.boxes):
            l, t, w, h = box["left"], box["top"], box["width"], box["height"]
            if l <= img_x <= l + w and t <= img_y <= t + h:
                if box["id"] == self.selected_id and (l + w - handle_zone) <= img_x and (t + h - handle_zone) <= img_y:
                    return box["id"], "resize"
                return box["id"], "move"
        return None, "draw"

    def _to_img_coords(self, event):
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        z = max(self.zoom, 0.001)
        return cx / z, cy / z

    def _on_press(self, event):
        if self.base_image_size == (0, 0):
            return
        img_x, img_y = self._to_img_coords(event)
        box_id, mode = self._hit_test(img_x, img_y)
        self._drag_mode = mode
        self._drag_start_img = (img_x, img_y)
        if mode in ("move", "resize"):
            self._drag_box_id = box_id
            box = self._get_box(box_id)
            self._drag_orig = dict(box)
            if box_id != self.selected_id:
                self.selected_id = box_id
                self._redraw_boxes()
                if self.on_select:
                    self.on_select(box_id)
        else:
            self._drag_box_id = None
            cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
            self._temp_rect = self.canvas.create_rectangle(cx, cy, cx, cy, outline="#ffffff", dash=(4, 2), width=2)

    def _on_motion(self, event):
        if self.base_image_size == (0, 0) or self._drag_mode is None:
            return
        img_x, img_y = self._to_img_coords(event)
        dx = img_x - self._drag_start_img[0]
        dy = img_y - self._drag_start_img[1]

        if self._drag_mode == "move" and self._drag_box_id:
            orig = self._drag_orig
            box = self._get_box(self._drag_box_id)
            box["left"] = max(0, orig["left"] + dx)
            box["top"] = max(0, orig["top"] + dy)
            self._update_box_visual(self._drag_box_id)
        elif self._drag_mode == "resize" and self._drag_box_id:
            orig = self._drag_orig
            box = self._get_box(self._drag_box_id)
            box["width"] = max(10, orig["width"] + dx)
            box["height"] = max(10, orig["height"] + dy)
            self._update_box_visual(self._drag_box_id)
        elif self._drag_mode == "draw" and self._temp_rect:
            cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
            x0, y0 = self._drag_start_img[0] * self.zoom, self._drag_start_img[1] * self.zoom
            self.canvas.coords(self._temp_rect, x0, y0, cx, cy)

    def _on_release(self, event):
        if self._drag_mode in ("move", "resize") and self._drag_box_id:
            box = self._get_box(self._drag_box_id)
            if box and self.on_box_change:
                self.on_box_change(
                    self._drag_box_id, int(box["left"]), int(box["top"]),
                    int(box["width"]), int(box["height"]),
                )
        elif self._drag_mode == "draw" and self._temp_rect:
            img_x, img_y = self._to_img_coords(event)
            x0, y0 = self._drag_start_img
            left, top = min(x0, img_x), min(y0, img_y)
            width, height = abs(img_x - x0), abs(img_y - y0)
            self.canvas.delete(self._temp_rect)
            self._temp_rect = None
            if width > 8 and height > 8 and self.on_new_box:
                self.on_new_box(int(left), int(top), int(width), int(height))

        self._drag_mode = None
        self._drag_box_id = None
        self._drag_orig = None


# ============================================================================ #
# Launcher view: Start New Project / Open Project
# ============================================================================ #

class LauncherFrame(tb.Frame):
    def __init__(self, master, app):
        super().__init__(master, style="App.TFrame")
        self.app = app
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        rail = tb.Frame(self, padding=(28, 30), style="Sidebar.TFrame", width=292)
        rail.grid(row=0, column=0, sticky="nsew")
        rail.grid_propagate(False)
        brand = tb.Frame(rail, style="Sidebar.TFrame")
        brand.pack(fill="x", pady=(0, 38))
        self.brand_mark = ImageTk.PhotoImage(studio_mark(34))
        tb.Label(brand, image=self.brand_mark, style="Brand.TLabel").pack(side="left", padx=(0, 10))
        tb.Label(brand, text="Certificate\nStudio", style="Brand.TLabel").pack(side="left")
        tb.Label(rail, text="Create certificates", style="Section.TLabel").pack(anchor="w")
        tb.Label(
            rail, text="Design and generate polished PDF certificates from a spreadsheet.",
            style="MutedSidebar.TLabel", wraplength=220, justify="left",
        ).pack(anchor="w", pady=(8, 22))
        RoundedButton(
            rail, "+  New project", self._start_new, kind="primary", width=234,
            height=42, canvas_bg=COLORS["sidebar"],
        ).pack(fill="x")
        tb.Frame(rail, style="Sidebar.TFrame").pack(fill="both", expand=True)
        tb.Label(rail, text="LOCAL WORKSPACE", style="Eyebrow.TLabel").pack(anchor="w")
        tb.Label(rail, text=str(PROJECTS_ROOT), style="Path.TLabel", wraplength=220, justify="left").pack(
            anchor="w", pady=(7, 0)
        )

        body = tb.Frame(self, padding=(44, 34), style="App.TFrame")
        body.grid(row=0, column=1, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        right = tb.Frame(body, style="App.TFrame")
        right.grid(row=0, column=0, sticky="nsew")
        right.rowconfigure(2, weight=1)
        right.columnconfigure(0, weight=1)

        header = tb.Frame(right, style="App.TFrame")
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 24))
        tb.Label(header, text="Your projects", style="Hero.TLabel").pack(side="left")
        RoundedButton(header, "Refresh", self.refresh, width=88, canvas_bg=COLORS["bg"]).pack(side="right")
        self.project_summary = tb.Label(right, text="Continue where you left off", style="Muted.TLabel")
        self.project_summary.grid(
            row=1, column=0, sticky="w", pady=(0, 12)
        )

        columns = ("name", "created", "updated")
        self.tree = tb.Treeview(right, columns=columns, show="headings", style="Projects.Treeview")
        self.tree.heading("name", text="    Project", anchor="w")
        self.tree.heading("created", text="Created", anchor="w")
        self.tree.heading("updated", text="Last edited", anchor="w")
        self.tree.column("name", width=280, minwidth=160, anchor="w")
        self.tree.column("created", width=200, minwidth=170, anchor="w", stretch=False)
        self.tree.column("updated", width=200, minwidth=170, anchor="w", stretch=False)
        self.tree.grid(row=2, column=0, sticky="nsew")
        self.tree.bind("<Double-1>", lambda e: self._open_selected())
        self.tree.bind("<Return>", lambda e: self._open_selected())
        self.tree.tag_configure("even", background=COLORS["bg"])
        self.tree.tag_configure("odd", background="#1e1e1e")

        scroll = tb.Scrollbar(right, orient="vertical", command=self.tree.yview)
        def update_scroll(first, last):
            scroll.set(first, last)
            if float(first) <= 0 and float(last) >= 1:
                scroll.grid_remove()
            else:
                scroll.grid(row=2, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=update_scroll)

        btn_row = tb.Frame(right, style="App.TFrame")
        btn_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(16, 0))
        RoundedButton(btn_row, "Open project", self._open_selected, kind="primary", width=126,
                      canvas_bg=COLORS["bg"]).pack(side="left")
        RoundedButton(btn_row, "Delete", self._delete_selected, kind="danger", width=82,
                      canvas_bg=COLORS["bg"]).pack(
            side="left", padx=(10, 0)
        )

        self._rows = {}
        self.refresh()

    def refresh(self):
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        self._rows = {}
        projects = self.app.pm.list_projects()
        self.project_summary.configure(text=(f"{len(projects)} projects  ·  Your local workspace" if projects
                                             else "Your first project starts here. Choose + New project to begin."))
        for index, proj in enumerate(projects):
            iid = str(proj["id"])
            self.tree.insert("", "end", iid=iid, values=("    " + proj["name"],
                             display_project_date(proj["created_at"]), display_project_date(proj["updated_at"])),
                             tags=("even" if index % 2 == 0 else "odd",))
            self._rows[iid] = proj
        if selected and selected[0] in self._rows:
            self.tree.selection_set(selected[0])

    def _selected_project(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self._rows.get(sel[0])

    def _start_new(self):
        dialog = ProjectNameDialog(self)
        self.wait_window(dialog)
        name = dialog.result
        if not name:
            return
        try:
            proj = self.app.pm.create_project(name)
        except ValueError as exc:
            messagebox.showwarning("Couldn't create project", str(exc))
            return
        self.refresh()
        self.app.open_project(proj)

    def _open_selected(self):
        proj = self._selected_project()
        if not proj:
            messagebox.showinfo("No project selected", "Select a project from the list first.")
            return
        self.app.open_project(proj)

    def _delete_selected(self):
        proj = self._selected_project()
        if not proj:
            messagebox.showinfo("No project selected", "Select a project from the list first.")
            return
        if not messagebox.askyesno(
            "Delete project",
            f"Permanently delete project '{proj['name']}' and ALL of its files "
            f"(template, roster copy, and every generated certificate)?\nThis cannot be undone.",
        ):
            return
        self.app.pm.delete_project(proj["id"], proj["folder"], delete_files=True)
        self.refresh()


# ============================================================================ #
# Workspace view: the main certificate-building UI for the open project
# ============================================================================ #

class WorkspaceFrame(tb.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.pm = app.pm

        self.project = None
        self.info = None
        self.workbook = None
        self.col_index_map = {}
        self.header_labels = ["(none)"]
        self.column_tokens = {}
        self.text_boxes = []
        self.selected_box_id = None
        self._next_box_seq = 1
        self._sample_row_cache = None
        self._sample_variables = {}

        self.msg_queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.generation_running = False
        self._preview_after_id = None
        self._suspend_traces = False

        self._build_ui()
        self.after(150, self._poll_queue)

    # ------------------------------------------------------------------ #
    # UI construction
    # ------------------------------------------------------------------ #
    def _build_ui(self):
        topbar = tb.Frame(self, padding=(20, 12), style="Topbar.TFrame")
        topbar.pack(fill="x", side="top")
        self.project_title_var = tk.StringVar(value="No project open")
        tb.Label(topbar, text="Certificate Studio", style="EyebrowTop.TLabel").pack(side="left")
        tb.Label(topbar, text="   /   ", style="MutedTop.TLabel").pack(side="left")
        tb.Label(topbar, textvariable=self.project_title_var, style="ProjectTitle.TLabel").pack(side="left")
        RoundedButton(topbar, "Save", self.save_project, kind="primary", width=76,
                      canvas_bg=COLORS["surface"]).pack(side="right", padx=(8, 0))
        RoundedButton(topbar, "All projects", self._switch_project, width=104,
                      canvas_bg=COLORS["surface"]).pack(side="right")
        RoundedButton(topbar, "Open folder", self._open_project_folder, width=104,
                      canvas_bg=COLORS["surface"]).pack(
            side="right", padx=(0, 6)
        )
        tb.Separator(self, orient="horizontal").pack(fill="x", side="top")

        body = tb.Frame(self, padding=0, style="App.TFrame")
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=0, minsize=420)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        left_container = tb.Frame(body, padding=(18, 18, 10, 18), style="Sidebar.TFrame")
        left_container.grid(row=0, column=0, sticky="nsew")
        tb.Label(left_container, text="Project editor", style="Section.TLabel").pack(anchor="w", pady=(0, 6))
        tb.Label(left_container, text="Set up, personalize, then export.",
                 style="MutedSidebar.TLabel").pack(anchor="w", pady=(0, 18))
        self.editor_tabs = tb.Notebook(left_container, style="Editor.TNotebook")
        self.editor_tabs.pack(fill="both", expand=True)
        panels = {}
        for title in ("Setup", "Design", "Export"):
            scroller = ScrollableFrame(self.editor_tabs)
            self.editor_tabs.add(scroller, text=title, padding=0)
            panels[title] = scroller.inner
        self._build_files_section(panels["Setup"])
        self._build_columns_section(panels["Setup"])
        self._build_variables_section(panels["Design"])
        self._build_textboxes_section(panels["Design"])
        self._build_metadata_section(panels["Export"])
        self._build_output_section(panels["Export"])
        self._build_actions_section(left_container)

        right = tb.Frame(body, padding=(16, 16, 18, 14), style="App.TFrame")
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(0, weight=1)
        right.rowconfigure(1, weight=0)
        right.rowconfigure(2, weight=0)
        right.columnconfigure(0, weight=1)

        preview_frame = tb.Frame(right, padding=0, style="App.TFrame")
        preview_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
        preview_frame.rowconfigure(0, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        self.canvas_preview = InteractiveCanvasPreview(
            preview_frame,
            on_box_change=self._on_canvas_box_change,
            on_new_box=self._on_canvas_new_box,
            on_select=self._on_canvas_select,
        )
        self.canvas_preview.grid(row=0, column=0, sticky="nsew")

        progress_frame = tb.Frame(right, style="App.TFrame")
        progress_frame.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        progress_frame.columnconfigure(0, weight=1)
        self.progress_bar = tb.Progressbar(progress_frame, mode="determinate", style="Accent.Horizontal.TProgressbar")
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.status_var = tk.StringVar(value="Open or create a project to begin.")
        tb.Label(progress_frame, textvariable=self.status_var, style="Muted.TLabel").grid(row=0, column=1, sticky="e")

        self.log_collapsible = Collapsible(right, "Activity", collapsed=True)
        self.log_collapsible.grid(row=2, column=0, sticky="ew")
        self.log_text = tb.ScrolledText(self.log_collapsible.body, height=6, autohide=True)
        self.log_text.pack(fill="both", expand=True)
        self.log_text.text.configure(state="disabled", font=(UI_FONT, 9), bg=COLORS["bg"],
                                     fg=COLORS["muted"], relief="flat", borderwidth=0)

    # -- Files section ---------------------------------------------------
    def _build_files_section(self, panel):
        box = tb.Labelframe(panel, text="FILES", padding=12)
        box.pack(fill="x", pady=(0, 10))
        box.columnconfigure(0, weight=1)

        self.template_path_var = tk.StringVar()
        self.font_path_var = tk.StringVar()
        self.excel_path_var = tk.StringVar()

        self._path_row(box, 0, "Template image", self.template_path_var, self._browse_template)
        self._path_row(box, 1, "Default font (.ttf/.otf)", self.font_path_var, self._browse_default_font)
        self._path_row(box, 2, "Excel roster (.xlsx)", self.excel_path_var, self._browse_excel)

        tb.Label(box, text="Files are copied into this project's folder automatically.", bootstyle=SECONDARY).grid(
            row=6, column=0, columnspan=2, sticky="w", pady=(6, 0)
        )

    def _path_row(self, parent, row, label, var, command):
        tb.Label(parent, text=label).grid(row=row * 2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        entry = tb.Entry(parent, textvariable=var, state="readonly")
        entry.grid(row=row * 2 + 1, column=0, sticky="ew", padx=(0, 6), pady=(0, 2))
        RoundedButton(parent, text="Browse…", command=command).grid(
            row=row * 2 + 1, column=1, sticky="e", pady=(0, 2)
        )

    def _browse_template(self):
        if not self._require_project():
            return
        path = filedialog.askopenfilename(
            title="Choose certificate template",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp *.webp"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            with Image.open(path) as img:
                img.verify()
        except Exception as exc:
            messagebox.showerror("Invalid image", f"Couldn't open that image:\n{exc}")
            return
        rel = copy_asset_into(self._project_dir(), path, "assets", desired_stem="template")
        self.info["template_image"] = rel
        self.template_path_var.set(rel)
        self.status_var.set("Template loaded.")
        self._schedule_preview()
        self.save_project(silent=True)

    def _browse_default_font(self):
        if not self._require_project():
            return
        path = filedialog.askopenfilename(
            title="Choose a default font file", filetypes=[("Font files", "*.ttf *.otf"), ("All files", "*.*")]
        )
        if not path:
            return
        rel = copy_asset_into(self._project_dir(), path, "assets/fonts", desired_stem="default_font")
        self.info["default_font_file"] = rel
        self.font_path_var.set(rel)
        self._schedule_preview()
        self.save_project(silent=True)

    def _browse_excel(self):
        if not self._require_project():
            return
        path = filedialog.askopenfilename(title="Choose Excel roster", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not path:
            return
        try:
            openpyxl.load_workbook(path, read_only=True, data_only=True)
        except Exception as exc:
            messagebox.showerror("Invalid Excel file", f"Couldn't open that spreadsheet:\n{exc}")
            return
        rel = copy_asset_into(self._project_dir(), path, "assets", desired_stem="roster")
        self.info["excel_file"] = rel
        self.excel_path_var.set(rel)
        # always read back from the project's own copy from here on
        self.workbook = openpyxl.load_workbook(self._abs(rel), read_only=True, data_only=True)
        self.sheet_combo["values"] = self.workbook.sheetnames
        self.sheet_var.set(self.workbook.active.title)
        self.info["sheet_name"] = self.workbook.active.title
        self._load_sheet_columns()
        self.status_var.set(f"Excel loaded — {len(self.workbook.sheetnames)} sheet(s) found.")
        self.save_project(silent=True)

    # -- Columns section ---------------------------------------------------
    def _build_columns_section(self, panel):
        box = tb.Labelframe(panel, text="DATA COLUMNS", padding=12)
        box.pack(fill="x", pady=(0, 10))
        box.columnconfigure(1, weight=1)

        self.sheet_var = tk.StringVar()
        self.primary_col_var = tk.StringVar()
        self.cert_num_col_var = tk.StringVar()
        self.filename_col_var = tk.StringVar()
        self.group_col_var = tk.StringVar()

        tb.Label(box, text="Sheet").grid(row=0, column=0, sticky="w")
        self.sheet_combo = tb.Combobox(box, textvariable=self.sheet_var, state="readonly")
        self.sheet_combo.grid(row=0, column=1, sticky="ew", pady=2)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda e: self._on_sheet_selected())

        tb.Label(box, text="Primary column (required — skips blank rows)").grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.primary_col_combo = tb.Combobox(box, textvariable=self.primary_col_var, state="readonly")
        self.primary_col_combo.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.primary_col_combo.bind("<<ComboboxSelected>>", lambda e: self._on_columns_changed())

        tb.Label(box, text="Certificate Number column (optional → {certificate_number})").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(8, 0)
        )
        self.cert_num_combo = tb.Combobox(box, textvariable=self.cert_num_col_var, state="readonly")
        self.cert_num_combo.grid(row=4, column=0, columnspan=2, sticky="ew")
        self.cert_num_combo.bind("<<ComboboxSelected>>", lambda e: self._on_columns_changed())

        tb.Label(box, text="File name column (optional)").grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.filename_col_combo = tb.Combobox(box, textvariable=self.filename_col_var, state="readonly")
        self.filename_col_combo.grid(row=6, column=0, columnspan=2, sticky="ew")
        self.filename_col_combo.bind("<<ComboboxSelected>>", lambda e: self._on_columns_changed())

        tb.Label(box, text="Group / subfolder column (optional)").grid(row=7, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.group_col_combo = tb.Combobox(box, textvariable=self.group_col_var, state="readonly")
        self.group_col_combo.grid(row=8, column=0, columnspan=2, sticky="ew")
        self.group_col_combo.bind("<<ComboboxSelected>>", lambda e: self._on_columns_changed())

    def _on_sheet_selected(self):
        self.info["sheet_name"] = self.sheet_var.get()
        self._load_sheet_columns()
        self.save_project(silent=True)

    def _on_columns_changed(self):
        self.info["primary_col_header"] = self._strip_none(self.primary_col_var.get())
        self.info["certificate_number_col_header"] = self._strip_none(self.cert_num_col_var.get())
        self.info["filename_col_header"] = self._strip_none(self.filename_col_var.get())
        self.info["group_col_header"] = self._strip_none(self.group_col_var.get())
        self._refresh_sample_row()
        self._schedule_preview()
        self._refresh_variables_reference()
        self.save_project(silent=True)

    @staticmethod
    def _strip_none(label):
        return None if (not label or label == "(none)") else label

    def _load_sheet_columns(self):
        if not self.workbook:
            return
        sheet = self.workbook[self.sheet_var.get()]
        header_row = next(sheet.iter_rows(min_row=1, max_row=1), [])
        labels = ["(none)"]
        self.col_index_map = {"(none)": None}
        self.column_tokens = {}
        used_tokens = set()
        for idx, cell in enumerate(header_row, start=1):
            header = str(cell.value).strip() if cell.value is not None else ""
            col_letter = openpyxl.utils.get_column_letter(idx)
            label = f"{col_letter}: {header}" if header else f"{col_letter}: (blank header)"
            labels.append(label)
            self.col_index_map[label] = idx
            token = make_token(header or f"col_{col_letter}", used_tokens)
            used_tokens.add(token)
            self.column_tokens[label] = token

        self.header_labels = labels
        for combo in (self.primary_col_combo, self.cert_num_combo, self.filename_col_combo, self.group_col_combo):
            combo["values"] = labels

        self._apply_stored_or_guessed_columns(labels)
        self._sample_row_cache = None
        self._refresh_sample_row()
        self._schedule_preview()
        self._refresh_variables_reference()

    def _apply_stored_or_guessed_columns(self, labels):
        def resolve(stored_header, guess_contains=None):
            if stored_header:
                match = next(
                    (lbl for lbl in labels if lbl == stored_header or lbl.endswith(": " + stored_header.split(": ")[-1])),
                    None,
                )
                if match:
                    return match
            if guess_contains:
                match = next((lbl for lbl in labels[1:] if guess_contains in lbl.lower()), None)
                if match:
                    return match
            return "(none)"

        primary = resolve(self.info.get("primary_col_header"), "name")
        if primary == "(none)" and len(labels) > 1:
            primary = labels[1]
        self.primary_col_var.set(primary)
        self.cert_num_col_var.set(resolve(self.info.get("certificate_number_col_header"), "cert"))
        self.filename_col_var.set(resolve(self.info.get("filename_col_header")))
        self.group_col_var.set(resolve(self.info.get("group_col_header")))

        self.info["primary_col_header"] = self._strip_none(self.primary_col_var.get())
        self.info["certificate_number_col_header"] = self._strip_none(self.cert_num_col_var.get())
        self.info["filename_col_header"] = self._strip_none(self.filename_col_var.get())
        self.info["group_col_header"] = self._strip_none(self.group_col_var.get())

    # -- Variables reference ---------------------------------------------------
    def _build_variables_section(self, panel):
        box = tb.Labelframe(panel, text="VARIABLES", padding=12)
        box.pack(fill="x", pady=(0, 10))
        box.columnconfigure(0, weight=1)
        self.variables_text = tk.Text(
            box, height=5, wrap="word", bg=COLORS["surface_2"], fg=COLORS["text"],
            insertbackground=COLORS["text"], relief="flat", padx=8, pady=8,
        )
        self.variables_text.grid(row=0, column=0, sticky="ew")
        self.variables_text.configure(state="disabled")
        RoundedButton(box, text="Shuffle preview", command=self._randomize_sample_row).grid(
            row=1, column=0, sticky="ew", pady=(6, 0)
        )

    def _refresh_variables_reference(self):
        variables = self._current_sample_variables()
        self.variables_text.configure(state="normal")
        self.variables_text.delete("1.0", "end")
        if not variables:
            self.variables_text.insert("end", "Load an Excel roster to see available {variables}.")
        else:
            for token, value in variables.items():
                self.variables_text.insert("end", f"{{{token}}}  →  {value}\n")
        self.variables_text.configure(state="disabled")

    def available_tokens(self):
        tokens = list(self.column_tokens.values())
        for extra in ("certificate_number", "row_number", "excel_row"):
            if extra not in tokens:
                tokens.append(extra)
        return tokens

    def _current_sample_variables(self):
        if not self.workbook:
            return {}
        return self._sample_variables

    def _refresh_sample_row(self):
        self._sample_variables = {}
        if not self.workbook:
            return
        sheet = self.workbook[self.sheet_var.get()]
        row_to_use = self._sample_row_cache
        if row_to_use is None:
            for row in sheet.iter_rows(min_row=2, max_row=2):
                row_to_use = row
                break
            self._sample_row_cache = row_to_use
        variables = {}
        for label, token in self.column_tokens.items():
            idx = self.col_index_map.get(label)
            val = ""
            if row_to_use and idx and idx - 1 < len(row_to_use):
                cell_val = row_to_use[idx - 1].value
                val = str(cell_val).strip() if cell_val is not None else ""
            variables[token] = val or "(sample)"
        cert_label = self.cert_num_col_var.get()
        cert_idx = self.col_index_map.get(cert_label)
        cert_val = ""
        if row_to_use and cert_idx and cert_idx - 1 < len(row_to_use):
            cv = row_to_use[cert_idx - 1].value
            cert_val = str(cv).strip() if cv else ""
        variables["certificate_number"] = cert_val or "CERT-0001"
        variables["row_number"] = "1"
        variables["excel_row"] = "2"
        self._sample_variables = variables

    def _randomize_sample_row(self):
        if not self.workbook:
            return
        sheet = self.workbook[self.sheet_var.get()]
        rows = list(sheet.iter_rows(min_row=2, max_row=min(sheet.max_row, 300)))
        rows = [r for r in rows if any(c.value not in (None, "") for c in r)]
        if rows:
            self._sample_row_cache = random.choice(rows)
        self._refresh_sample_row()
        self._refresh_variables_reference()
        self._schedule_preview()

    # -- Text boxes section ---------------------------------------------------
    def _build_textboxes_section(self, panel):
        box = tb.Labelframe(panel, text="TEXT LAYERS", padding=12)
        box.pack(fill="x", pady=(0, 10))
        box.columnconfigure(0, weight=1)

        list_row = tb.Frame(box)
        list_row.grid(row=0, column=0, sticky="ew")
        self.box_listbox = tk.Listbox(
            list_row, height=5, exportselection=False, bg=COLORS["surface_2"], fg=COLORS["text"],
            selectbackground=COLORS["selection"], selectforeground=COLORS["text"], highlightthickness=1,
            highlightbackground=COLORS["border"], relief="flat", borderwidth=0,
        )
        self.box_listbox.pack(side="left", fill="both", expand=True)
        self.box_listbox.bind("<<ListboxSelect>>", lambda e: self._on_listbox_select())
        list_scroll = tb.Scrollbar(list_row, orient="vertical", command=self.box_listbox.yview)
        self.box_listbox.configure(yscrollcommand=list_scroll.set)
        list_scroll.pack(side="right", fill="y")

        btn_row = tb.Frame(box)
        btn_row.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        RoundedButton(btn_row, text="+ Add", command=self._add_text_box).pack(side="left")
        RoundedButton(btn_row, text="Duplicate", command=self._duplicate_text_box).pack(side="left", padx=(6, 0))
        RoundedButton(btn_row, text="Delete", kind="danger", command=self._delete_text_box).pack(side="left", padx=(6, 0))

        editor = tb.Labelframe(box, text="Selected Box", padding=8)
        editor.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        editor.columnconfigure(1, weight=1)
        self.box_editor = editor

        self.box_label_var = tk.StringVar()
        self.box_template_var = tk.StringVar()
        self.box_left_var = tk.IntVar(value=0)
        self.box_top_var = tk.IntVar(value=0)
        self.box_width_var = tk.IntVar(value=100)
        self.box_height_var = tk.IntVar(value=100)
        self.box_font_size_var = tk.IntVar(value=60)
        self.box_color_var = tk.StringVar(value="#000066")
        self.box_align_var = tk.StringVar(value="center")
        self.box_font_override_var = tk.StringVar(value="")

        r = 0
        tb.Label(editor, text="Label").grid(row=r, column=0, sticky="w")
        e = tb.Entry(editor, textvariable=self.box_label_var)
        e.grid(row=r, column=1, sticky="ew", pady=2)
        e.bind("<KeyRelease>", lambda ev: self._commit_box_field("label", self.box_label_var.get()))
        r += 1

        tb.Label(editor, text="Template text").grid(row=r, column=0, sticky="w", pady=(6, 0))
        r += 1
        tmpl_row = tb.Frame(editor)
        tmpl_row.grid(row=r, column=0, columnspan=2, sticky="ew")
        tmpl_row.columnconfigure(0, weight=1)
        self.box_template_entry = tb.Entry(tmpl_row, textvariable=self.box_template_var)
        self.box_template_entry.grid(row=0, column=0, sticky="ew")
        self.box_template_entry.bind(
            "<KeyRelease>", lambda ev: (self._commit_box_field("template", self.box_template_var.get()), self._schedule_preview())
        )
        self._make_variable_inserter(tmpl_row, self.box_template_entry).grid(row=0, column=1, padx=(4, 0))
        r += 1

        coord_row = tb.Frame(editor)
        coord_row.grid(row=r, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        for i, (lbl, var, key) in enumerate(
            [("Left", self.box_left_var, "left"), ("Top", self.box_top_var, "top"),
             ("Width", self.box_width_var, "width"), ("Height", self.box_height_var, "height")]
        ):
            sub = tb.Frame(coord_row)
            sub.grid(row=0, column=i, padx=(0 if i == 0 else 4, 0), sticky="ew")
            coord_row.columnconfigure(i, weight=1)
            tb.Label(sub, text=lbl, bootstyle=SECONDARY).pack(anchor="w")
            spin = tb.Spinbox(sub, from_=0, to=20000, textvariable=var, width=7)
            spin.pack(fill="x")
            spin.bind("<KeyRelease>", lambda ev, k=key, v=var: self._commit_box_field(k, self._safe_int(v)))
            spin.bind("<<Increment>>", lambda ev, k=key, v=var: self._commit_box_field(k, self._safe_int(v)))
            spin.bind("<<Decrement>>", lambda ev, k=key, v=var: self._commit_box_field(k, self._safe_int(v)))
        r += 1

        font_row = tb.Frame(editor)
        font_row.grid(row=r, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        tb.Label(font_row, text="Font size").pack(side="left")
        fspin = tb.Spinbox(font_row, from_=6, to=500, textvariable=self.box_font_size_var, width=6)
        fspin.pack(side="left", padx=(4, 12))
        fspin.bind("<KeyRelease>", lambda ev: (self._commit_box_field("font_size", self._safe_int(self.box_font_size_var)), self._schedule_preview()))
        fspin.bind("<<Increment>>", lambda ev: (self._commit_box_field("font_size", self._safe_int(self.box_font_size_var)), self._schedule_preview()))
        fspin.bind("<<Decrement>>", lambda ev: (self._commit_box_field("font_size", self._safe_int(self.box_font_size_var)), self._schedule_preview()))

        tb.Label(font_row, text="Align").pack(side="left")
        align_combo = tb.Combobox(font_row, textvariable=self.box_align_var, values=ALIGN_OPTIONS, state="readonly", width=8)
        align_combo.pack(side="left", padx=(4, 0))
        align_combo.bind("<<ComboboxSelected>>", lambda ev: (self._commit_box_field("align", self.box_align_var.get()), self._schedule_preview()))
        r += 1

        color_row = tb.Frame(editor)
        color_row.grid(row=r, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        tb.Label(color_row, text="Color").pack(side="left")
        self.box_color_swatch = tk.Canvas(color_row, width=22, height=18, highlightthickness=1, highlightbackground="#888")
        self.box_color_swatch.pack(side="left", padx=(4, 4))
        RoundedButton(color_row, text="Choose…", command=self._pick_box_color).pack(side="left")
        r += 1

        font_override_row = tb.Frame(editor)
        font_override_row.grid(row=r, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        tb.Label(font_override_row, text="Font override (optional)").pack(side="left")
        tb.Entry(font_override_row, textvariable=self.box_font_override_var, state="readonly", width=16).pack(
            side="left", padx=(4, 4), fill="x", expand=True
        )
        RoundedButton(font_override_row, text="Browse…", command=self._pick_box_font).pack(side="left")
        RoundedButton(font_override_row, text="Clear", command=self._clear_box_font).pack(side="left", padx=(4, 0))

        self._set_editor_enabled(False)

    def _make_variable_inserter(self, parent, entry_widget):
        btn = tb.Menubutton(parent, text="{ }", bootstyle="secondary-outline", width=3)
        menu = tk.Menu(btn, tearoff=0)
        btn.configure(menu=menu)

        def rebuild(*_):
            menu.delete(0, "end")
            for token in self.available_tokens():
                menu.add_command(label="{" + token + "}", command=lambda t=token: self._insert_token(entry_widget, t))

        btn.bind("<Button-1>", rebuild)
        return btn

    def _insert_token(self, entry_widget, token):
        try:
            pos = entry_widget.index(tk.INSERT)
        except Exception:
            pos = tk.END
        entry_widget.insert(pos, "{" + token + "}")
        entry_widget.event_generate("<KeyRelease>")
        entry_widget.focus_set()

    @staticmethod
    def _safe_int(var):
        try:
            return int(var.get())
        except Exception:
            return 0

    def _set_editor_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        for child in self.box_editor.winfo_children():
            self._set_widget_state_recursive(child, state)

    def _set_widget_state_recursive(self, widget, state):
        try:
            widget.configure(state=state)
        except tk.TclError:
            pass
        for child in widget.winfo_children():
            self._set_widget_state_recursive(child, state)

    def _next_color(self):
        idx = len(self.text_boxes) % len(BOX_COLOR_PALETTE)
        return BOX_COLOR_PALETTE[idx]

    def _add_text_box(self, left=None, top=None, width=None, height=None):
        w, h = (self.template_pixel_size() or (1000, 700))
        box = {
            "id": f"box_{self._next_box_seq}",
            "label": f"Text {self._next_box_seq}",
            "template": "",
            "left": left if left is not None else int(w * 0.25),
            "top": top if top is not None else int(h * 0.45),
            "width": width if width is not None else int(w * 0.5),
            "height": height if height is not None else int(h * 0.12),
            "font_file": None,
            "font_size": 60,
            "color": "#000066",
            "align": "center",
            "display_color": self._next_color(),
        }
        self._next_box_seq += 1
        self.text_boxes.append(box)
        self.selected_box_id = box["id"]
        self._refresh_box_listbox()
        self._load_box_into_editor(box)
        self._schedule_preview()
        self.save_project(silent=True)
        return box

    def _duplicate_text_box(self):
        box = self._get_selected_box()
        if not box:
            return
        new_box = dict(box)
        new_box["id"] = f"box_{self._next_box_seq}"
        new_box["label"] = box["label"] + " copy"
        new_box["left"] = box["left"] + 20
        new_box["top"] = box["top"] + 20
        new_box["display_color"] = self._next_color()
        self._next_box_seq += 1
        self.text_boxes.append(new_box)
        self.selected_box_id = new_box["id"]
        self._refresh_box_listbox()
        self._load_box_into_editor(new_box)
        self._schedule_preview()
        self.save_project(silent=True)

    def _delete_text_box(self):
        box = self._get_selected_box()
        if not box:
            return
        if not messagebox.askyesno("Delete text box", f"Delete text box '{box['label']}'?"):
            return
        self.text_boxes = [b for b in self.text_boxes if b["id"] != box["id"]]
        self.selected_box_id = self.text_boxes[0]["id"] if self.text_boxes else None
        self._refresh_box_listbox()
        if self.selected_box_id:
            self._load_box_into_editor(self._get_selected_box())
        else:
            self._set_editor_enabled(False)
        self._schedule_preview()
        self.save_project(silent=True)

    def _get_selected_box(self):
        for b in self.text_boxes:
            if b["id"] == self.selected_box_id:
                return b
        return None

    def _refresh_box_listbox(self):
        self.box_listbox.delete(0, "end")
        for b in self.text_boxes:
            self.box_listbox.insert("end", b["label"] or b["id"])
        if self.selected_box_id:
            for i, b in enumerate(self.text_boxes):
                if b["id"] == self.selected_box_id:
                    self.box_listbox.selection_clear(0, "end")
                    self.box_listbox.selection_set(i)
                    break

    def _on_listbox_select(self):
        sel = self.box_listbox.curselection()
        if not sel:
            return
        box = self.text_boxes[sel[0]]
        self.selected_box_id = box["id"]
        self._load_box_into_editor(box)
        self._redraw_canvas_boxes()

    def _load_box_into_editor(self, box):
        self._suspend_traces = True
        self.box_label_var.set(box.get("label", ""))
        self.box_template_var.set(box.get("template", ""))
        self.box_left_var.set(int(box.get("left", 0)))
        self.box_top_var.set(int(box.get("top", 0)))
        self.box_width_var.set(int(box.get("width", 100)))
        self.box_height_var.set(int(box.get("height", 100)))
        self.box_font_size_var.set(int(box.get("font_size", 60)))
        self.box_color_var.set(box.get("color", "#000066"))
        self.box_align_var.set(box.get("align", "center"))
        self.box_font_override_var.set(box.get("font_file") or "")
        self._redraw_box_color_swatch()
        self._set_editor_enabled(True)
        self._suspend_traces = False

    def _redraw_box_color_swatch(self):
        self.box_color_swatch.delete("all")
        self.box_color_swatch.create_rectangle(0, 0, 22, 18, fill=self.box_color_var.get(), outline="")

    def _commit_box_field(self, field, value):
        if self._suspend_traces:
            return
        box = self._get_selected_box()
        if not box:
            return
        box[field] = value
        if field == "label":
            self._refresh_box_listbox()
        if field in ("left", "top", "width", "height"):
            self._redraw_canvas_boxes()
        self.save_project(silent=True)

    def _pick_box_color(self):
        box = self._get_selected_box()
        if not box:
            return
        _, hex_color = colorchooser.askcolor(color=self.box_color_var.get(), title="Text color")
        if hex_color:
            self.box_color_var.set(hex_color)
            self._redraw_box_color_swatch()
            self._commit_box_field("color", hex_color)
            self._schedule_preview()

    def _pick_box_font(self):
        box = self._get_selected_box()
        if not box:
            return
        path = filedialog.askopenfilename(title="Choose font for this box", filetypes=[("Font files", "*.ttf *.otf"), ("All files", "*.*")])
        if not path:
            return
        rel = copy_asset_into(self._project_dir(), path, "assets/fonts", desired_stem=box["id"])
        self.box_font_override_var.set(rel)
        self._commit_box_field("font_file", rel)
        self._schedule_preview()

    def _clear_box_font(self):
        self.box_font_override_var.set("")
        self._commit_box_field("font_file", None)
        self._schedule_preview()

    def template_pixel_size(self):
        if not self.info or not self.info.get("template_image"):
            return None
        try:
            with Image.open(self._abs(self.info["template_image"])) as img:
                return img.size
        except Exception:
            return None

    # -- canvas callbacks ---------------------------------------------------
    def _on_canvas_box_change(self, box_id, left, top, width, height):
        box = next((b for b in self.text_boxes if b["id"] == box_id), None)
        if not box:
            return
        box["left"], box["top"], box["width"], box["height"] = left, top, width, height
        if box_id == self.selected_box_id:
            self._suspend_traces = True
            self.box_left_var.set(left)
            self.box_top_var.set(top)
            self.box_width_var.set(width)
            self.box_height_var.set(height)
            self._suspend_traces = False
        self._schedule_preview()
        self.save_project(silent=True)

    def _on_canvas_new_box(self, left, top, width, height):
        self.editor_tabs.select(1)
        self._add_text_box(left, top, width, height)

    def _on_canvas_select(self, box_id):
        self.selected_box_id = box_id
        box = self._get_selected_box()
        if box:
            self.editor_tabs.select(1)
            self._load_box_into_editor(box)
            self._refresh_box_listbox()

    def _redraw_canvas_boxes(self):
        lightweight = [
            {
                "id": b["id"], "left": b["left"], "top": b["top"], "width": b["width"], "height": b["height"],
                "color": b.get("display_color", "#ff5252"), "label": b.get("label", ""),
            }
            for b in self.text_boxes
        ]
        self.canvas_preview.set_boxes(lightweight, selected_id=self.selected_box_id)

    # -- Metadata section ---------------------------------------------------
    def _build_metadata_section(self, panel):
        box = tb.Labelframe(panel, text="PDF METADATA", padding=12)
        box.pack(fill="x", pady=(0, 10))
        box.columnconfigure(1, weight=1)

        self.meta_vars = {}
        r = 0
        for key, label in METADATA_STANDARD_FIELDS:
            tb.Label(box, text=label).grid(row=r, column=0, sticky="w", pady=2)
            row_frame = tb.Frame(box)
            row_frame.grid(row=r, column=1, sticky="ew", pady=2)
            row_frame.columnconfigure(0, weight=1)
            var = tk.StringVar()
            self.meta_vars[key] = var
            entry = tb.Entry(row_frame, textvariable=var)
            entry.grid(row=0, column=0, sticky="ew")
            entry.bind("<KeyRelease>", lambda ev, k=key, v=var: self._commit_meta_field(k, v.get()))
            self._make_variable_inserter(row_frame, entry).grid(row=0, column=1, padx=(4, 0))
            r += 1

        tb.Separator(box, orient="horizontal").grid(row=r, column=0, columnspan=2, sticky="ew", pady=6)
        r += 1
        tb.Label(box, text="Custom fields", bootstyle=SECONDARY).grid(row=r, column=0, columnspan=2, sticky="w")
        r += 1

        custom_row = tb.Frame(box)
        custom_row.grid(row=r, column=0, columnspan=2, sticky="ew")
        custom_row.columnconfigure(0, weight=1)
        self.custom_meta_tree = tb.Treeview(custom_row, columns=("key", "value"), show="headings", height=4, style="Editor.Treeview")
        self.custom_meta_tree.heading("key", text="Key")
        self.custom_meta_tree.heading("value", text="Value template")
        self.custom_meta_tree.column("key", width=100)
        self.custom_meta_tree.column("value", width=220)
        self.custom_meta_tree.grid(row=0, column=0, sticky="ew")
        self.custom_meta_tree.bind("<<TreeviewSelect>>", lambda e: self._on_custom_meta_select())
        r += 1

        self.custom_key_var = tk.StringVar()
        self.custom_value_var = tk.StringVar()
        add_row = tb.Frame(box)
        add_row.grid(row=r, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        add_row.columnconfigure(1, weight=1)
        tb.Entry(add_row, textvariable=self.custom_key_var, width=12).grid(row=0, column=0, sticky="w")
        val_entry = tb.Entry(add_row, textvariable=self.custom_value_var)
        val_entry.grid(row=0, column=1, sticky="ew", padx=(4, 4))
        self._make_variable_inserter(add_row, val_entry).grid(row=0, column=2)
        btns = tb.Frame(box)
        btns.grid(row=r + 1, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        RoundedButton(btns, text="Add / update", command=self._add_custom_meta).pack(side="left")
        RoundedButton(btns, text="Remove", kind="danger", command=self._remove_custom_meta).pack(side="left", padx=(6, 0))

        tb.Label(box, text="PDF resolution (DPI)").grid(row=r + 2, column=0, sticky="w", pady=(8, 0))
        self.pdf_resolution_var = tk.IntVar(value=100)
        spin = tb.Spinbox(box, from_=50, to=1200, textvariable=self.pdf_resolution_var, width=8)
        spin.grid(row=r + 2, column=1, sticky="w", pady=(8, 0))
        spin.bind("<KeyRelease>", lambda e: self._commit_pdf_resolution())
        spin.bind("<<Increment>>", lambda e: self._commit_pdf_resolution())
        spin.bind("<<Decrement>>", lambda e: self._commit_pdf_resolution())

    def _commit_meta_field(self, key, value):
        if not self.info:
            return
        self.info["metadata"][key] = value
        self.save_project(silent=True)

    def _commit_pdf_resolution(self):
        if not self.info:
            return
        try:
            self.info["pdf_resolution"] = int(self.pdf_resolution_var.get())
        except Exception:
            pass
        self.save_project(silent=True)

    def _refresh_custom_meta_tree(self):
        self.custom_meta_tree.delete(*self.custom_meta_tree.get_children())
        for k, v in (self.info.get("metadata", {}).get("custom", {}) or {}).items():
            self.custom_meta_tree.insert("", "end", iid=k, values=(k, v))

    def _on_custom_meta_select(self):
        sel = self.custom_meta_tree.selection()
        if not sel:
            return
        k = sel[0]
        v = self.info["metadata"]["custom"].get(k, "")
        self.custom_key_var.set(k)
        self.custom_value_var.set(v)

    def _add_custom_meta(self):
        key = self.custom_key_var.get().strip()
        if not key:
            messagebox.showinfo("Missing key", "Enter a metadata key (e.g. CertificateNumber).")
            return
        self.info.setdefault("metadata", {}).setdefault("custom", {})[key] = self.custom_value_var.get()
        self._refresh_custom_meta_tree()
        self.save_project(silent=True)

    def _remove_custom_meta(self):
        sel = self.custom_meta_tree.selection()
        if not sel:
            return
        key = sel[0]
        self.info.get("metadata", {}).get("custom", {}).pop(key, None)
        self._refresh_custom_meta_tree()
        self.save_project(silent=True)

    # -- Output section ---------------------------------------------------
    def _build_output_section(self, panel):
        box = tb.Labelframe(panel, text="OUTPUT", padding=12)
        box.pack(fill="x", pady=(0, 10))
        box.columnconfigure(0, weight=1)
        self.output_path_var = tk.StringVar()
        tb.Entry(box, textvariable=self.output_path_var, state="readonly").grid(row=0, column=0, sticky="ew", padx=(0, 6))
        RoundedButton(box, text="Open Folder", command=lambda: open_in_file_manager(self.output_path_var.get())).grid(
            row=0, column=1
        )
        tb.Label(box, text="Certificates always save inside this project's own 'output' folder.", bootstyle=SECONDARY).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(6, 0)
        )

    # -- Actions section ---------------------------------------------------
    def _build_actions_section(self, panel):
        box = tb.Frame(panel, padding=(0, 4))
        box.pack(fill="x", pady=(0, 20))
        box.columnconfigure(0, weight=1)

        RoundedButton(box, text="Save one test certificate", command=self._generate_test_certificate).pack(
            fill="x", pady=(0, 8)
        )
        self.generate_btn = RoundedButton(
            box, "Generate all certificates", self._start_generation, kind="primary",
            width=390, height=44, canvas_bg=COLORS["sidebar"],
        )
        self.generate_btn.pack(fill="x")
        self.cancel_btn = RoundedButton(box, text="Cancel", kind="danger", command=self._cancel_generation, state="disabled")
        self.cancel_btn.pack(fill="x", pady=(8, 0))

    # ------------------------------------------------------------------ #
    # Project lifecycle
    # ------------------------------------------------------------------ #
    def _project_dir(self):
        return self.pm.project_path(self.project["folder"])

    def _abs(self, rel_path):
        if not rel_path:
            return None
        return str(self._project_dir() / rel_path)

    def _require_project(self):
        if not self.project:
            messagebox.showinfo("No project", "Open or create a project first.")
            return False
        return True

    def load_project(self, project):
        self.editor_tabs.select(0)
        self.project = project
        self.info = self.pm.load_info(project["folder"])
        self.workbook = None
        self.text_boxes = []
        self.selected_box_id = None
        self._next_box_seq = 1
        self._sample_row_cache = None
        self._sample_variables = {}

        self.project_title_var.set(project['name'])
        self.output_path_var.set(str(self._project_dir() / "output"))
        (self._project_dir() / "output").mkdir(parents=True, exist_ok=True)

        self.template_path_var.set(self.info.get("template_image") or "")
        self.font_path_var.set(self.info.get("default_font_file") or "")
        self.excel_path_var.set(self.info.get("excel_file") or "")
        self.pdf_resolution_var.set(int(self.info.get("pdf_resolution", 100)))

        for key, _ in METADATA_STANDARD_FIELDS:
            self.meta_vars[key].set(self.info.get("metadata", {}).get(key, ""))
        self._refresh_custom_meta_tree()

        if self.info.get("excel_file"):
            try:
                self.workbook = openpyxl.load_workbook(self._abs(self.info["excel_file"]), read_only=True, data_only=True)
                self.sheet_combo["values"] = self.workbook.sheetnames
                sheet_name = self.info.get("sheet_name") or self.workbook.active.title
                if sheet_name not in self.workbook.sheetnames:
                    sheet_name = self.workbook.active.title
                self.sheet_var.set(sheet_name)
                self._load_sheet_columns()
            except Exception as exc:
                messagebox.showwarning("Couldn't reload Excel", str(exc))

        for stored in self.info.get("text_boxes", []):
            box = dict(stored)
            box.setdefault("display_color", self._next_color())
            self.text_boxes.append(box)
            try:
                seq = int(str(box["id"]).split("_")[-1])
                self._next_box_seq = max(self._next_box_seq, seq + 1)
            except Exception:
                pass
        self._refresh_box_listbox()
        if self.text_boxes:
            self.selected_box_id = self.text_boxes[0]["id"]
            self._load_box_into_editor(self.text_boxes[0])
        else:
            self._set_editor_enabled(False)

        self.status_var.set("Project loaded.")
        self._schedule_preview()
        self.after(200, self.canvas_preview.fit)

    def save_project(self, silent=False):
        if not self.project or not self.info:
            return
        self.info["text_boxes"] = [{k: v for k, v in b.items() if k != "display_color"} for b in self.text_boxes]
        try:
            self.pm.save_info(self.project["folder"], self.info, self.project["name"])
        except Exception as exc:
            if not silent:
                messagebox.showerror("Couldn't save project", str(exc))
            return
        if not silent:
            self.status_var.set("Project saved.")

    def _switch_project(self):
        self.save_project(silent=True)
        self.app.show_launcher()

    def _open_project_folder(self):
        if not self.project:
            return
        open_in_file_manager(self._project_dir())

    # ------------------------------------------------------------------ #
    # Live preview
    # ------------------------------------------------------------------ #
    def _schedule_preview(self):
        if self._preview_after_id:
            self.after_cancel(self._preview_after_id)
        self._preview_after_id = self.after(200, self._update_preview)

    def _update_preview(self):
        self._preview_after_id = None
        if not self.info or not self.info.get("template_image"):
            self.canvas_preview.clear()
            return
        variables = self._current_sample_variables() or {}
        default_font = self._abs(self.info.get("default_font_file"))
        boxes_for_render = []
        for b in self.text_boxes:
            b2 = dict(b)
            if b2.get("font_file"):
                b2["font_file"] = self._abs(b2["font_file"])
            boxes_for_render.append(b2)
        try:
            cert = build_certificate_image(self._abs(self.info["template_image"]), boxes_for_render, variables, default_font)
        except Exception as exc:
            self.status_var.set(f"Preview error: {exc}")
            return
        self.canvas_preview.set_base_image(cert)
        self._redraw_canvas_boxes()

    # ------------------------------------------------------------------ #
    # Validation
    # ------------------------------------------------------------------ #
    def _validate_ready(self, need_excel=True):
        if not self._require_project():
            return False
        if not self.info.get("template_image"):
            messagebox.showwarning("Missing template", "Please choose a certificate template image.")
            return False
        if need_excel and not self.info.get("excel_file"):
            messagebox.showwarning("Missing roster", "Please choose an Excel roster file.")
            return False
        if need_excel and not self.info.get("primary_col_header"):
            messagebox.showwarning("Missing primary column", "Please choose the primary (required) column.")
            return False
        if not self.text_boxes:
            messagebox.showwarning("No text boxes", "Add at least one text box (e.g. for the recipient's name).")
            return False
        return True

    # ------------------------------------------------------------------ #
    # Single test certificate
    # ------------------------------------------------------------------ #
    def _generate_test_certificate(self):
        if not self._validate_ready(need_excel=False):
            return
        try:
            out_dir = self._project_dir() / "output"
            out_dir.mkdir(parents=True, exist_ok=True)
            variables = self._current_sample_variables() or {}
            default_font = self._abs(self.info.get("default_font_file"))
            boxes_for_render = []
            for b in self.text_boxes:
                b2 = dict(b)
                if b2.get("font_file"):
                    b2["font_file"] = self._abs(b2["font_file"])
                boxes_for_render.append(b2)
            cert = build_certificate_image(self._abs(self.info["template_image"]), boxes_for_render, variables, default_font)
            metadata = build_metadata_dict(self.info.get("metadata", {}), variables)
            out_path = out_dir / f"TEST_{sanitize_filename(variables.get('row_number', 'sample'))}.pdf"
            save_certificate_pdf(cert, self.info.get("pdf_resolution", 100), metadata, out_path)
            self._log(f"Saved test certificate to {out_path}")
            messagebox.showinfo("Test certificate saved", f"Saved:\n{out_path}")
        except Exception as exc:
            messagebox.showerror("Failed to save test certificate", str(exc))
            self._log("ERROR: " + traceback.format_exc())

    # ------------------------------------------------------------------ #
    # Bulk generation
    # ------------------------------------------------------------------ #
    def _start_generation(self):
        if self.generation_running:
            return
        if not self._validate_ready(need_excel=True):
            return
        if not messagebox.askyesno("Confirm", "Generate a PDF certificate for every matching row in the spreadsheet?"):
            return

        self.save_project(silent=True)
        self.cancel_event.clear()
        self.generation_running = True
        self.generate_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.progress_bar.configure(value=0)
        self._clear_log()
        self._log("Starting certificate generation…")

        boxes_snapshot = []
        for b in self.text_boxes:
            b2 = dict(b)
            if b2.get("font_file"):
                b2["font_file"] = self._abs(b2["font_file"])
            boxes_snapshot.append(b2)

        config = dict(
            excel_path=self._abs(self.info["excel_file"]),
            sheet_name=self.sheet_var.get(),
            primary_col=self.col_index_map.get(self.primary_col_var.get()),
            filename_col=self.col_index_map.get(self.filename_col_var.get()),
            group_col=self.col_index_map.get(self.group_col_var.get()),
            cert_num_col=self.col_index_map.get(self.cert_num_col_var.get()),
            column_tokens=dict(self.column_tokens),
            col_index_map=dict(self.col_index_map),
            template_path=self._abs(self.info["template_image"]),
            default_font=self._abs(self.info.get("default_font_file")),
            text_boxes=boxes_snapshot,
            metadata_config=dict(self.info.get("metadata", {})),
            pdf_resolution=self.info.get("pdf_resolution", 100),
            output_folder=str(self._project_dir() / "output"),
        )
        thread = threading.Thread(target=self._generation_worker, args=(config,), daemon=True)
        thread.start()

    def _cancel_generation(self):
        self.cancel_event.set()
        self.cancel_btn.configure(state="disabled")
        self._log("Cancelling… (in-flight certificates will finish, no new ones will start)")

    def _generation_worker(self, config):
        try:
            with open(config["template_path"], "rb") as f:
                template_bytes = f.read()

            wb = openpyxl.load_workbook(config["excel_path"], read_only=True, data_only=True)
            sheet = wb[config["sheet_name"]]

            rows = []
            row_number = 0
            for excel_row_num, row in enumerate(sheet.iter_rows(min_row=2), start=2):
                primary_idx = config["primary_col"]
                primary_val = row[primary_idx - 1].value if primary_idx and primary_idx - 1 < len(row) else None
                if not primary_val or not str(primary_val).strip():
                    continue
                row_number += 1

                variables = {}
                for label, token in config["column_tokens"].items():
                    idx = config["col_index_map"].get(label)
                    val = ""
                    if idx and idx - 1 < len(row):
                        cv = row[idx - 1].value
                        val = str(cv).strip() if cv is not None else ""
                    variables[token] = val
                cert_idx = config["cert_num_col"]
                cert_val = ""
                if cert_idx and cert_idx - 1 < len(row):
                    cv = row[cert_idx - 1].value
                    cert_val = str(cv).strip() if cv else ""
                variables["certificate_number"] = cert_val
                variables["row_number"] = str(row_number)
                variables["excel_row"] = str(excel_row_num)

                fname_base = str(primary_val).strip()
                if config["filename_col"] and config["filename_col"] - 1 < len(row):
                    fv = row[config["filename_col"] - 1].value
                    if fv and str(fv).strip():
                        fname_base = str(fv).strip()

                group_val = None
                if config["group_col"] and config["group_col"] - 1 < len(row):
                    gv = row[config["group_col"] - 1].value
                    if gv and str(gv).strip():
                        group_val = str(gv).strip()

                rows.append((variables, fname_base, group_val))

            total = len(rows)
            if total == 0:
                self.msg_queue.put(("error", "No data rows found under the selected primary column."))
                return

            self.msg_queue.put(("start", total))
            success, failed = 0, 0
            max_workers = min(8, (os.cpu_count() or 4))
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {}
                for variables, fname_base, group_val in rows:
                    if self.cancel_event.is_set():
                        break
                    folder = config["output_folder"]
                    if group_val:
                        folder = os.path.join(folder, sanitize_filename(group_val))
                    fut = executor.submit(
                        self._render_one, template_bytes, config["text_boxes"], variables,
                        config["default_font"], config["metadata_config"], config["pdf_resolution"],
                        folder, fname_base,
                    )
                    futures[fut] = variables.get("row_number", "?")

                completed = 0
                for fut in as_completed(futures):
                    tag = futures[fut]
                    try:
                        ok, info = fut.result()
                    except Exception as exc:
                        ok, info = False, str(exc)
                    completed += 1
                    if ok:
                        success += 1
                    else:
                        failed += 1
                    self.msg_queue.put(("progress", completed, total, f"row {tag}", ok, info))

            cancelled = self.cancel_event.is_set()
            self.msg_queue.put(("done", success, failed, total, cancelled))
        except Exception:
            self.msg_queue.put(("error", traceback.format_exc()))

    @staticmethod
    def _render_one(template_bytes, text_boxes, variables, default_font, metadata_config, resolution, folder, fname_base):
        try:
            os.makedirs(folder, exist_ok=True)
            out_path = os.path.join(folder, f"{sanitize_filename(fname_base)}.pdf")
            cert = build_certificate_image(template_bytes, text_boxes, variables, default_font)
            metadata = build_metadata_dict(metadata_config, variables)
            save_certificate_pdf(cert, resolution, metadata, out_path)
            return True, out_path
        except Exception as exc:
            return False, str(exc)

    # ------------------------------------------------------------------ #
    # Queue polling / logging
    # ------------------------------------------------------------------ #
    def _poll_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                kind = msg[0]
                if kind == "start":
                    total = msg[1]
                    self.progress_bar.configure(maximum=total, value=0)
                    self.status_var.set(f"0 / {total}")
                elif kind == "progress":
                    completed, total, name, ok, info = msg[1:]
                    self.progress_bar.configure(value=completed)
                    self.status_var.set(f"{completed} / {total}")
                    self._log(f"{'✓' if ok else '✗ FAILED'} {name} → {info}")
                elif kind == "done":
                    success, failed, total, cancelled = msg[1:]
                    label = "Cancelled" if cancelled else "Done"
                    self._log(f"{label}. Success: {success}  Failed: {failed}  Total: {total}")
                    self.status_var.set(f"{label} — {success}/{total} succeeded")
                    self._finish_generation()
                elif kind == "error":
                    self._log(f"ERROR: {msg[1]}")
                    messagebox.showerror("Generation error", str(msg[1]))
                    self._finish_generation()
        except queue.Empty:
            pass
        self.after(150, self._poll_queue)

    def _finish_generation(self):
        self.generation_running = False
        self.generate_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")

    def _log(self, text):
        self.log_text.text.configure(state="normal")
        self.log_text.text.insert("end", text + "\n")
        self.log_text.text.see("end")
        self.log_text.text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.text.configure(state="normal")
        self.log_text.text.delete("1.0", "end")
        self.log_text.text.configure(state="disabled")


# ============================================================================ #
# Main application window
# ============================================================================ #

class CertificateStudioApp(tb.Window):
    def __init__(self):
        super().__init__(themename="darkly", title="Certificate Studio", iconphoto=None)
        self.geometry("1440x900")
        self.minsize(1100, 700)

        # Keep PhotoImages alive for Tk on macOS/Linux. Windows uses its native
        # multi-resolution ICO path for both this window and future dialogs.
        self._app_icons = [ImageTk.PhotoImage(studio_mark(size)) for size in (16, 32, 48, 64)]
        self.iconphoto(True, *self._app_icons)
        if platform.system() == "Windows":
            icon_path = str(APP_DIR / "assets" / "studio.ico")
            self.iconbitmap(default=icon_path)
            self.iconbitmap(icon_path)

        self._configure_visual_system()

        self.pm = ProjectManager()

        container = tb.Frame(self)
        container.pack(fill="both", expand=True)
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        self.launcher = LauncherFrame(container, self)
        self.workspace = WorkspaceFrame(container, self)
        self.launcher.grid(row=0, column=0, sticky="nsew")
        self.workspace.grid(row=0, column=0, sticky="nsew")

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.show_launcher()
        self.after(80, self._style_native_window)

    def _style_native_window(self):
        """Ask Windows 11 to use dark chrome and rounded native window corners."""
        if platform.system() != "Windows":
            return
        try:
            import ctypes

            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            dwm = ctypes.windll.dwmapi.DwmSetWindowAttribute
            dark = ctypes.c_int(1)
            corner = ctypes.c_int(2)  # DWMWCP_ROUND
            caption = ctypes.c_int(0x00191919)  # COLORREF is BGR
            border = ctypes.c_int(0x00383838)
            text = ctypes.c_int(0x00F1F1F1)
            dwm(hwnd, 20, ctypes.byref(dark), ctypes.sizeof(dark))
            dwm(hwnd, 33, ctypes.byref(corner), ctypes.sizeof(corner))
            dwm(hwnd, 34, ctypes.byref(border), ctypes.sizeof(border))
            dwm(hwnd, 35, ctypes.byref(caption), ctypes.sizeof(caption))
            dwm(hwnd, 36, ctypes.byref(text), ctypes.sizeof(text))
        except Exception:
            pass

    def _configure_visual_system(self):
        """Apply the visual language before constructing the application views."""
        global UI_FONT, DISPLAY_FONT
        self.configure(bg=COLORS["bg"])
        s = self.style
        available = set(tkfont.families(self))
        font = next((name for name in ("Inter", "Aptos", "Segoe UI Variable Text", "Segoe UI", "Helvetica")
                     if name in available), "TkDefaultFont")
        display_font = next((name for name in ("Inter Display", "Inter", "Aptos Display", "Segoe UI Variable Display")
                             if name in available), font)
        UI_FONT, DISPLAY_FONT = font, display_font
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont"):
            tkfont.nametofont(name).configure(family=font, size=10)
        self.option_add("*Font", (font, 10))
        s.configure(".", font=(font, 10), background=COLORS["bg"], foreground=COLORS["text"])
        s.configure("App.TFrame", background=COLORS["bg"])
        s.configure("Sidebar.TFrame", background=COLORS["sidebar"])
        s.configure("Surface.TFrame", background=COLORS["surface"])
        s.configure("Topbar.TFrame", background=COLORS["surface"])
        s.configure("TFrame", background=COLORS["sidebar"])
        s.configure("TLabel", background=COLORS["sidebar"], foreground=COLORS["text"], font=(font, 10))
        s.configure("TButton", font=(font, 10))
        s.configure("Editor.TNotebook", background=COLORS["sidebar"], borderwidth=0, tabmargins=(0, 0, 0, 14))
        s.configure("Editor.TNotebook.Tab", background=COLORS["sidebar"], foreground=COLORS["muted"],
                    padding=(20, 10), font=(font, 10), borderwidth=0)
        s.map("Editor.TNotebook.Tab", background=[("selected", COLORS["surface_hover"]), ("active", COLORS["surface"])],
              foreground=[("selected", COLORS["text"]), ("active", COLORS["text"])])
        s.configure("Editor.Treeview", background=COLORS["surface"], fieldbackground=COLORS["surface"],
                    foreground=COLORS["text"], rowheight=34, borderwidth=0, font=(font, 10))
        s.configure("Editor.Treeview.Heading", background=COLORS["sidebar"], foreground=COLORS["muted"],
                    font=(font, 9), padding=8, relief="flat")
        s.map("Editor.Treeview", background=[("selected", COLORS["selection"])], foreground=[("selected", COLORS["text"])])
        s.configure("Hero.TLabel", background=COLORS["bg"], foreground=COLORS["text"], font=(display_font, 27, "bold"))
        s.configure("Brand.TLabel", background=COLORS["sidebar"], foreground=COLORS["text"], font=(display_font, 13, "bold"))
        s.configure("Section.TLabel", background=COLORS["sidebar"], foreground=COLORS["text"], font=(display_font, 14, "bold"))
        s.configure("Eyebrow.TLabel", background=COLORS["sidebar"], foreground=COLORS["muted"], font=(font, 8, "bold"))
        s.configure("EyebrowTop.TLabel", background=COLORS["surface"], foreground=COLORS["muted"], font=(font, 10, "bold"))
        s.configure("ProjectTitle.TLabel", background=COLORS["surface"], foreground=COLORS["text"], font=(font, 11, "bold"))
        s.configure("Muted.TLabel", background=COLORS["bg"], foreground=COLORS["muted"])
        s.configure("MutedSidebar.TLabel", background=COLORS["sidebar"], foreground=COLORS["muted"])
        s.configure("MutedTop.TLabel", background=COLORS["surface"], foreground=COLORS["muted"])
        s.configure("ToolbarMuted.TLabel", background=COLORS["surface"], foreground=COLORS["muted"])
        s.configure("Path.TLabel", background=COLORS["sidebar"], foreground="#6f6d68", font=("Consolas", 8))
        s.configure("Accent.TButton", background=COLORS["accent"], foreground="#171717", borderwidth=0, padding=(15, 8), font=(font, 10, "bold"))
        s.map(
            "Accent.TButton",
            background=[("active", COLORS["accent_hover"]), ("disabled", "#555555")],
            foreground=[("disabled", "#909090")],
        )
        s.configure("Quiet.TButton", background=COLORS["surface"], foreground=COLORS["muted"], borderwidth=0, padding=(12, 7))
        s.map("Quiet.TButton", background=[("active", COLORS["surface_hover"])], foreground=[("active", COLORS["text"])])
        s.configure("TLabelframe", background=COLORS["sidebar"], bordercolor=COLORS["sidebar"], relief="flat", borderwidth=0)
        s.configure("TLabelframe.Label", background=COLORS["sidebar"], foreground=COLORS["muted"], font=(font, 8, "bold"))
        s.configure("Card.TLabelframe", background=COLORS["surface"], bordercolor=COLORS["border"], relief="solid")
        s.configure("Card.TLabelframe.Label", background=COLORS["bg"], foreground=COLORS["muted"], font=(font, 9, "bold"))
        s.configure("TEntry", fieldbackground=COLORS["surface_2"], foreground=COLORS["text"], bordercolor=COLORS["border"], insertcolor=COLORS["text"], padding=7)
        s.configure("TCombobox", fieldbackground=COLORS["surface_2"], foreground=COLORS["text"], bordercolor=COLORS["border"], padding=6)
        for control in ("TEntry", "TCombobox", "TSpinbox"):
            s.configure(control, font=(font, 10), fieldbackground=COLORS["surface_2"],
                        foreground=COLORS["text"], bordercolor=COLORS["border"], lightcolor=COLORS["border"],
                        darkcolor=COLORS["border"], insertcolor=COLORS["text"])
            s.map(control, bordercolor=[("focus", "#777777"), ("!focus", COLORS["border"])],
                  fieldbackground=[("readonly", COLORS["surface_2"]), ("disabled", COLORS["sidebar"])],
                  foreground=[("disabled", COLORS["muted"]), ("readonly", COLORS["text"])])
        s.configure("Projects.Treeview", background=COLORS["bg"], fieldbackground=COLORS["bg"], foreground=COLORS["text"],
                    font=(font, 11), rowheight=62, borderwidth=0, relief="flat")
        s.configure("Projects.Treeview.Heading", background=COLORS["bg"], foreground=COLORS["muted"], relief="flat",
                    borderwidth=0, padding=(0, 14), font=(font, 9))
        s.map("Projects.Treeview.Heading", background=[("active", COLORS["bg"])], relief=[("active", "flat")])
        s.map("Projects.Treeview", background=[("selected", COLORS["selection"])], foreground=[("selected", COLORS["text"])])
        s.configure("Accent.Horizontal.TProgressbar", troughcolor=COLORS["surface_2"], background=COLORS["accent"], borderwidth=0)

    def show_launcher(self):
        self.launcher.refresh()
        self.launcher.tkraise()

    def open_project(self, project):
        self.workspace.load_project(project)
        self.workspace.tkraise()

    def _on_close(self):
        try:
            self.workspace.save_project(silent=True)
        except Exception:
            pass
        self.destroy()


def main():
    app = CertificateStudioApp()
    app.mainloop()


if __name__ == "__main__":
    main()
