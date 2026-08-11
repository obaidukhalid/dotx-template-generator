"""
app.py
------
Local web GUI for building .dotx Word templates.

Run:
    python app.py

Then open http://127.0.0.1:5000 in a browser. The browser opens automatically.
"""

import copy
import json
import os
import platform
import subprocess
import sys
import threading
import webbrowser

from flask import Flask, jsonify, request, send_file, render_template

from template_builder import build_template

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
PRESETS_PATH = os.path.join(APP_DIR, "presets.json")

app = Flask(__name__)
app.config["JSON_SORT_KEYS"] = False


# ----------------------------------------------------------------------------
# Field schema. This single structure drives the entire form.
# Add a field here and it appears in the GUI with no other changes.
# ----------------------------------------------------------------------------

FONTS = [
    "Calibri", "Arial", "Helvetica", "Segoe UI", "Verdana", "Tahoma",
    "Trebuchet MS", "Times New Roman", "Georgia", "Cambria", "Garamond",
    "Book Antiqua", "Courier New", "Consolas",
]

BULLETS = [
    ["\u2022", "\u2022  filled circle"],
    ["\u25e6", "\u25e6  open circle"],
    ["\u25aa", "\u25aa  small square"],
    ["\u25ab", "\u25ab  open square"],
    ["\u2013", "\u2013  en dash"],
    ["\u2014", "\u2014  em dash"],
    ["\u25ba", "\u25ba  triangle"],
    ["\u2023", "\u2023  small triangle"],
    ["\u2043", "\u2043  hyphen bullet"],
    ["\u2192", "\u2192  arrow"],
]


def _text_style_fields(base, include_line_spacing=True, include_caps=False,
                       include_indent=False):
    """Build the standard set of fields for a paragraph style."""
    fields = [
        {"path": f"{base}.font", "label": "Font", "type": "select",
         "options": FONTS},
        {"path": f"{base}.fontSize", "label": "Size", "type": "number",
         "min": 6, "max": 96, "step": 0.5, "unit": "pt"},
        {"path": f"{base}.color", "label": "Colour", "type": "color"},
        {"path": f"{base}.bold", "label": "Bold", "type": "bool"},
        {"path": f"{base}.italic", "label": "Italic", "type": "bool"},
    ]
    if include_caps:
        fields.append(
            {"path": f"{base}.all_caps", "label": "All caps", "type": "bool"}
        )
    fields += [
        {"path": f"{base}.spacing_before", "label": "Space before",
         "type": "number", "min": 0, "max": 4000, "step": 20, "unit": "twips"},
        {"path": f"{base}.spacing_after", "label": "Space after",
         "type": "number", "min": 0, "max": 4000, "step": 20, "unit": "twips"},
    ]
    if include_line_spacing:
        fields.append(
            {"path": f"{base}.line_spacing", "label": "Line spacing",
             "type": "number", "min": 0.8, "max": 3, "step": 0.05,
             "unit": "lines"}
        )
    if include_indent:
        fields.append(
            {"path": f"{base}.indent", "label": "Left indent", "type": "number",
             "min": 0, "max": 4000, "step": 60, "unit": "twips"}
        )
    return fields


SCHEMA = [
    {
        "id": "template",
        "title": "Template details",
        "hint": "Metadata stored inside the file.",
        "fields": [
            {"path": "template.name", "label": "Template name", "type": "text"},
            {"path": "template.description", "label": "Description",
             "type": "text"},
            {"path": "template.author", "label": "Author", "type": "text"},
        ],
    },
    {
        "id": "document",
        "title": "Document structure",
        "hint": "What the generated template contains.",
        "fields": [
            {"path": "document.include_cover", "label": "Include cover page",
             "type": "bool"},
            {"path": "document.include_toc", "label": "Include contents page",
             "type": "bool"},
            {"path": "document.include_style_guide",
             "label": "Include style guide", "type": "bool"},
            {"path": "document.number_headings", "label": "Number the headings",
             "type": "bool",
             "help": "Word keeps the numbers in sequence by itself."},
            {"path": "document.number_format", "label": "Number format",
             "type": "select",
             "options": [
                 ["decimal", "1, 1.1, 1.1.1"],
                 ["decimal_dot", "1., 1.1., 1.1.1."],
                 ["legal", "1.0, 1.1, 1.1.1"],
                 ["chapter", "Chapter 1, 1.1, 1.1.1"],
                 ["section", "Section 1, 1.1, 1.1.1"],
                 ["outline", "I., A., 1., a."],
             ]},
            {"path": "document.number_levels", "label": "Levels numbered",
             "type": "select",
             "options": [[1, "Heading 1 only"], [2, "Heading 1 to 2"],
                         [3, "Heading 1 to 3"], [4, "Heading 1 to 4"]]},
            {"path": "document.number_suffix", "label": "Gap after number",
             "type": "select",
             "options": [["tab", "Tab"], ["space", "Space"],
                         ["nothing", "None"]]},
            {"path": "document.number_indent", "label": "Indent by level",
             "type": "bool",
             "help": "Off keeps every heading flush with the left margin."},
            {"path": "document.outline_type", "label": "Starter outline",
             "type": "select",
             "options": [
                 ["none", "None"],
                 ["proposal", "Proposal"],
                 ["report", "Report"],
                 ["guide", "Software usage guide"],
                 ["study_guide", "Study guide"],
                 ["specification", "Technical specification"],
             ],
             "help": "Adds empty Heading 1 sections for the chosen type."},
            {"path": "document.title_placeholder", "label": "Cover title",
             "type": "text"},
            {"path": "document.subtitle_placeholder", "label": "Cover subtitle",
             "type": "text"},
            {"path": "document.author_placeholder", "label": "Cover author",
             "type": "text"},
            {"path": "document.date_placeholder", "label": "Cover date",
             "type": "text"},
            {"path": "document.reference_placeholder",
             "label": "Cover reference", "type": "text"},
            {"path": "document.header_text", "label": "Page header text",
             "type": "text", "help": "Leave empty for no header."},
            {"path": "document.page_numbers",
             "label": "Page number in footer", "type": "bool"},
        ],
    },
    {
        "id": "heading_1",
        "title": "Heading 1",
        "fields": _text_style_fields(
            "styles.headings.heading_1", include_line_spacing=False,
            include_caps=True,
        ),
    },
    {
        "id": "heading_2",
        "title": "Heading 2",
        "fields": _text_style_fields(
            "styles.headings.heading_2", include_line_spacing=False,
            include_caps=True,
        ),
    },
    {
        "id": "heading_3",
        "title": "Heading 3",
        "fields": _text_style_fields(
            "styles.headings.heading_3", include_line_spacing=False,
            include_caps=True,
        ),
    },
    {
        "id": "heading_4",
        "title": "Heading 4",
        "fields": _text_style_fields(
            "styles.headings.heading_4", include_line_spacing=False,
            include_caps=True,
        ),
    },
    {
        "id": "body",
        "title": "Body text",
        "hint": "The default paragraph style. Everything else inherits from it.",
        "fields": _text_style_fields("styles.paragraph.body"),
    },
    {
        "id": "caption",
        "title": "Caption",
        "hint": "For figure and table captions.",
        "fields": _text_style_fields(
            "styles.paragraph.caption", include_line_spacing=False
        ),
    },
    {
        "id": "quote",
        "title": "Quote block",
        "fields": _text_style_fields(
            "styles.paragraph.quote", include_indent=True
        ),
    },
    {
        "id": "text_runs",
        "title": "Inline text styles",
        "hint": "Applied to words inside a paragraph. Leave colour empty to "
                "inherit the surrounding text colour.",
        "fields": [
            {"path": "styles.text.strong.bold", "label": "Strong is bold",
             "type": "bool"},
            {"path": "styles.text.strong.color", "label": "Strong colour",
             "type": "color", "allow_empty": True},
            {"path": "styles.text.emphasis.italic",
             "label": "Emphasis is italic", "type": "bool"},
            {"path": "styles.text.emphasis.color", "label": "Emphasis colour",
             "type": "color", "allow_empty": True},
            {"path": "styles.text.highlight.bold", "label": "Highlight is bold",
             "type": "bool"},
            {"path": "styles.text.highlight.color", "label": "Highlight colour",
             "type": "color", "allow_empty": True},
        ],
    },
    {
        "id": "bullet_1",
        "title": "Bullets level 1",
        "fields": [
            {"path": "styles.bullets.level_1.bullet_char", "label": "Bullet",
             "type": "select", "options": BULLETS},
            {"path": "styles.bullets.level_1.font", "label": "Font",
             "type": "select", "options": FONTS},
            {"path": "styles.bullets.level_1.fontSize", "label": "Size",
             "type": "number", "min": 6, "max": 48, "step": 0.5, "unit": "pt"},
            {"path": "styles.bullets.level_1.color", "label": "Colour",
             "type": "color"},
            {"path": "styles.bullets.level_1.indent", "label": "Indent",
             "type": "number", "min": 0, "max": 4000, "step": 60,
             "unit": "twips"},
            {"path": "styles.bullets.level_1.spacing_after",
             "label": "Space after", "type": "number", "min": 0, "max": 2000,
             "step": 20, "unit": "twips"},
            {"path": "styles.bullets.level_1.line_spacing",
             "label": "Line spacing", "type": "number", "min": 0.8, "max": 3,
             "step": 0.05, "unit": "lines"},
        ],
    },
    {
        "id": "bullet_2",
        "title": "Bullets level 2",
        "fields": [
            {"path": "styles.bullets.level_2.bullet_char", "label": "Bullet",
             "type": "select", "options": BULLETS},
            {"path": "styles.bullets.level_2.font", "label": "Font",
             "type": "select", "options": FONTS},
            {"path": "styles.bullets.level_2.fontSize", "label": "Size",
             "type": "number", "min": 6, "max": 48, "step": 0.5, "unit": "pt"},
            {"path": "styles.bullets.level_2.color", "label": "Colour",
             "type": "color"},
            {"path": "styles.bullets.level_2.indent", "label": "Indent",
             "type": "number", "min": 0, "max": 4000, "step": 60,
             "unit": "twips"},
            {"path": "styles.bullets.level_2.spacing_after",
             "label": "Space after", "type": "number", "min": 0, "max": 2000,
             "step": 20, "unit": "twips"},
            {"path": "styles.bullets.level_2.line_spacing",
             "label": "Line spacing", "type": "number", "min": 0.8, "max": 3,
             "step": 0.05, "unit": "lines"},
        ],
    },
    {
        "id": "toc",
        "title": "Table of contents",
        "hint": "The contents list is a live Word field. It fills in when the "
                "field is updated.",
        "fields": [
            {"path": "styles.toc.title", "label": "Heading text",
             "type": "text"},
            {"path": "styles.toc.levels", "label": "Levels shown",
             "type": "select",
             "options": [["1-1", "1 only"], ["1-2", "1 to 2"],
                         ["1-3", "1 to 3"], ["1-4", "1 to 4"]]},
            {"path": "styles.toc.title_font", "label": "Heading font",
             "type": "select", "options": FONTS},
            {"path": "styles.toc.title_fontSize", "label": "Heading size",
             "type": "number", "min": 8, "max": 60, "step": 0.5, "unit": "pt"},
            {"path": "styles.toc.title_bold", "label": "Heading bold",
             "type": "bool"},
            {"path": "styles.toc.title_color", "label": "Heading colour",
             "type": "color"},
            {"path": "styles.toc.spacing_after",
             "label": "Space after heading", "type": "number", "min": 0,
             "max": 2000, "step": 20, "unit": "twips"},
            {"path": "styles.toc.entry_font", "label": "Entry font",
             "type": "select", "options": FONTS},
            {"path": "styles.toc.entry_fontSize", "label": "Entry size",
             "type": "number", "min": 6, "max": 24, "step": 0.5, "unit": "pt"},
            {"path": "styles.toc.entry_color", "label": "Entry colour",
             "type": "color"},
            {"path": "styles.toc.entry_bold_level1",
             "label": "Level 1 entries bold", "type": "bool"},
            {"path": "styles.toc.entry_spacing_after",
             "label": "Space between entries", "type": "number", "min": 0,
             "max": 1000, "step": 10, "unit": "twips"},
        ],
    },
    {
        "id": "page",
        "title": "Page setup",
        "hint": "1 inch is 1440 twips. 1 cm is 567 twips.",
        "fields": [
            {"path": "page_setup.page_size", "label": "Page size",
             "type": "select",
             "options": ["A4", "Letter", "Legal", "A5"]},
            {"path": "page_setup.orientation", "label": "Orientation",
             "type": "select",
             "options": [["portrait", "Portrait"], ["landscape", "Landscape"]]},
            {"path": "page_setup.default_font", "label": "Default font",
             "type": "select", "options": FONTS},
            {"path": "page_setup.default_font_size", "label": "Default size",
             "type": "number", "min": 6, "max": 24, "step": 0.5, "unit": "pt"},
            {"path": "page_setup.margins.top", "label": "Top margin",
             "type": "number", "min": 0, "max": 5000, "step": 60,
             "unit": "twips"},
            {"path": "page_setup.margins.bottom", "label": "Bottom margin",
             "type": "number", "min": 0, "max": 5000, "step": 60,
             "unit": "twips"},
            {"path": "page_setup.margins.left", "label": "Left margin",
             "type": "number", "min": 0, "max": 5000, "step": 60,
             "unit": "twips"},
            {"path": "page_setup.margins.right", "label": "Right margin",
             "type": "number", "min": 0, "max": 5000, "step": 60,
             "unit": "twips"},
        ],
    },
]


# ----------------------------------------------------------------------------
# Config helpers
# ----------------------------------------------------------------------------

def load_config(path=CONFIG_PATH):
    with open(path, "r", encoding="utf-8") as handle:
        return json.load(handle)


def save_config(config, path=CONFIG_PATH):
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(config, handle, indent=2, ensure_ascii=False)


def load_presets():
    if not os.path.exists(PRESETS_PATH):
        return {}
    with open(PRESETS_PATH, "r", encoding="utf-8") as handle:
        return json.load(handle)


def deep_merge(base, overlay):
    """Merge overlay into a copy of base, recursing into dictionaries."""
    result = copy.deepcopy(base)
    for key, value in overlay.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def default_output_dir():
    for candidate in ("Desktop", "Documents"):
        path = os.path.join(os.path.expanduser("~"), candidate)
        if os.path.isdir(path):
            return path
    return os.path.expanduser("~")


# ----------------------------------------------------------------------------
# Native folder dialog, run in a separate process so it never blocks Flask
# ----------------------------------------------------------------------------

FOLDER_PICKER = (
    "import tkinter as tk\n"
    "from tkinter import filedialog\n"
    "root = tk.Tk()\n"
    "root.withdraw()\n"
    "root.attributes('-topmost', True)\n"
    "path = filedialog.askdirectory(title='Choose where to save the template')\n"
    "root.destroy()\n"
    "print(path or '')\n"
)


def pick_folder():
    """Open a native folder dialog. Returns the path, or None if unavailable."""
    try:
        result = subprocess.run(
            [sys.executable, "-c", FOLDER_PICKER],
            capture_output=True, text=True, timeout=300,
        )
    except Exception:
        return None
    if result.returncode != 0:
        return None
    path = result.stdout.strip()
    return path or None


def reveal_in_file_manager(path):
    """Open the containing folder and select the file where supported."""
    system = platform.system()
    folder = path if os.path.isdir(path) else os.path.dirname(path)

    if system == "Windows":
        if os.path.isfile(path):
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        else:
            os.startfile(folder)  # noqa: S606
    elif system == "Darwin":
        if os.path.isfile(path):
            subprocess.Popen(["open", "-R", path])
        else:
            subprocess.Popen(["open", folder])
    else:
        subprocess.Popen(["xdg-open", folder])


# ----------------------------------------------------------------------------
# Routes
# ----------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/bootstrap")
def api_bootstrap():
    return jsonify({
        "schema": SCHEMA,
        "config": load_config(),
        "presets": load_presets(),
        "output_dir": default_output_dir(),
        "tk_available": True,
    })


@app.route("/api/browse", methods=["POST"])
def api_browse():
    path = pick_folder()
    if not path:
        return jsonify({
            "ok": False,
            "error": "The folder dialog is not available. Type the path "
                     "into the box instead.",
        }), 200
    return jsonify({"ok": True, "path": path})


@app.route("/api/generate", methods=["POST"])
def api_generate():
    payload = request.get_json(force=True)

    # Merge over the saved defaults so a partial config can never crash the build.
    try:
        config = deep_merge(load_config(), payload.get("config") or {})
    except (OSError, ValueError):
        config = payload.get("config") or {}

    output_dir = (payload.get("output_dir") or "").strip()
    filename = (payload.get("filename") or "template").strip()

    if not filename.lower().endswith(".dotx"):
        filename = os.path.splitext(filename)[0] + ".dotx"
    filename = os.path.basename(filename)

    if not output_dir:
        output_dir = default_output_dir()
    output_dir = os.path.abspath(os.path.expanduser(output_dir))

    try:
        os.makedirs(output_dir, exist_ok=True)
    except OSError as error:
        return jsonify({"ok": False, "error": f"Cannot use that folder: {error}"})

    output_path = os.path.join(output_dir, filename)

    try:
        build_template(config, output_path)
    except Exception as error:  # surfaced to the user in the GUI
        return jsonify({"ok": False, "error": f"{type(error).__name__}: {error}"})

    if payload.get("save_config", True):
        try:
            save_config(config)
        except OSError:
            pass

    return jsonify({
        "ok": True,
        "path": output_path,
        "size_kb": round(os.path.getsize(output_path) / 1024, 1),
    })


@app.route("/api/open-folder", methods=["POST"])
def api_open_folder():
    payload = request.get_json(force=True)
    path = payload.get("path") or default_output_dir()
    if not os.path.exists(path):
        return jsonify({"ok": False, "error": "That path no longer exists."})
    try:
        reveal_in_file_manager(path)
    except Exception as error:
        return jsonify({"ok": False, "error": str(error)})
    return jsonify({"ok": True})


@app.route("/api/download", methods=["POST"])
def api_download():
    payload = request.get_json(force=True)
    path = payload.get("path")
    if not path or not os.path.isfile(path):
        return jsonify({"ok": False, "error": "File not found."}), 404
    return jsonify({"ok": True, "url": "/api/file?path=" + path})


@app.route("/api/file")
def api_file():
    path = request.args.get("path", "")
    if not os.path.isfile(path):
        return "Not found", 404
    return send_file(path, as_attachment=True)


@app.route("/api/config", methods=["POST"])
def api_save_config():
    payload = request.get_json(force=True)
    config = payload.get("config") or {}
    try:
        save_config(config)
    except OSError as error:
        return jsonify({"ok": False, "error": str(error)})
    return jsonify({"ok": True, "path": CONFIG_PATH})


@app.route("/api/preset", methods=["POST"])
def api_preset():
    payload = request.get_json(force=True)
    name = payload.get("name")
    current = payload.get("config") or load_config()
    presets = load_presets()
    if name not in presets:
        return jsonify({"ok": False, "error": "Unknown preset."})
    return jsonify({"ok": True, "config": deep_merge(current, presets[name])})


def open_browser():
    webbrowser.open("http://127.0.0.1:5000")


if __name__ == "__main__":
    if os.environ.get("WERKZEUG_RUN_MAIN") != "true":
        threading.Timer(1.2, open_browser).start()
    print("\n  Template Studio running at http://127.0.0.1:5000")
    print("  Press Ctrl+C to stop.\n")
    app.run(host="127.0.0.1", port=5000, debug=False)
