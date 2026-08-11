"""
app.py
------
Local web GUI for building .dotx Word templates.

Run:
    python app.py

Then open http://127.0.0.1:5000 in a browser. The browser opens automatically.
"""

import copy
import hmac
import json
import os
import platform
import secrets
import subprocess
import sys
import threading
import webbrowser
from urllib.parse import quote

from flask import Flask, jsonify, request, send_file, render_template

from template_builder import build_template

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
PRESETS_PATH = os.path.join(APP_DIR, "presets.json")

HOST = "127.0.0.1"
PORT = 5000

app = Flask(__name__)
app.config["JSON_SORT_KEYS"] = False


# ----------------------------------------------------------------------------
# Security
#
# The server binds to loopback, but that alone protects nothing: any page you
# happen to have open in the same browser can send requests to 127.0.0.1, and a
# hostile DNS record can point an attacker's domain at loopback so their scripts
# read the responses too. Three checks close that off.
#
#   1. Host header must be one we recognise, which is what stops DNS rebinding:
#      the browser sends the attacker's hostname, not ours.
#   2. Origin, when the browser sends one, must be this app.
#   3. Every mutating request must carry a token that is only ever handed to the
#      real page. A cross-origin caller cannot read it.
#
# Requests are also parsed as strict JSON, so a form POST — the one shape that
# crosses origins without a preflight — is rejected before any handler runs.
# ----------------------------------------------------------------------------

CSRF_TOKEN = secrets.token_urlsafe(32)

ALLOWED_HOSTS = frozenset({
    f"127.0.0.1:{PORT}", f"localhost:{PORT}", f"[::1]:{PORT}",
})
ALLOWED_ORIGINS = frozenset(f"http://{host}" for host in ALLOWED_HOSTS)

# Serving a file by path is a read primitive, so it is limited to files this
# process actually generated rather than anything on disk.
GENERATED_FILES = set()
GENERATED_LOCK = threading.Lock()


def remember_generated(path):
    with GENERATED_LOCK:
        GENERATED_FILES.add(os.path.abspath(path))


def is_generated(path):
    return os.path.abspath(path) in GENERATED_FILES


def deny(message, code=403):
    return jsonify({"ok": False, "error": message}), code


@app.before_request
def guard_request():
    if request.host not in ALLOWED_HOSTS:
        return deny("Unrecognised Host header.")

    if request.method not in ("GET", "HEAD", "OPTIONS"):
        origin = request.headers.get("Origin")
        if origin is not None and origin not in ALLOWED_ORIGINS:
            return deny("Cross-origin request refused.")
        if not hmac.compare_digest(
            request.headers.get("X-CSRF-Token", ""), CSRF_TOKEN
        ):
            return deny("Missing or invalid CSRF token.")
    return None


def read_json():
    """Parse a JSON body strictly. Never force: the Content-Type requirement is
    what makes a cross-origin form POST impossible without a preflight."""
    payload = request.get_json(silent=True)
    return payload if isinstance(payload, dict) else {}


def conform_to(candidate, template):
    """Return candidate reshaped to template: unknown keys dropped, values kept
    only when their type matches. Config is written straight to disk, so it is
    never trusted to define its own shape."""
    if not isinstance(candidate, dict):
        return copy.deepcopy(template)
    result = copy.deepcopy(template)
    for key, reference in template.items():
        if key not in candidate:
            continue
        value = candidate[key]
        if isinstance(reference, dict):
            result[key] = conform_to(value, reference)
        elif isinstance(reference, bool):
            if isinstance(value, bool):
                result[key] = value
        elif isinstance(reference, (int, float)):
            # bool is an int subclass; reject it for a numeric field.
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                result[key] = value
        elif isinstance(reference, str):
            if isinstance(value, str):
                result[key] = value
        elif isinstance(reference, list):
            if isinstance(value, list):
                result[key] = value
    return result


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
    return render_template("index.html", csrf_token=CSRF_TOKEN)


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
    payload = read_json()

    # Merge over the saved defaults so a partial config can never crash the
    # build, and conform it so the same shape rules apply here as in
    # /api/config — this path writes config.json too.
    try:
        config = conform_to(payload.get("config") or {}, load_config())
    except (OSError, ValueError) as error:
        return jsonify({"ok": False, "error": f"Cannot read config: {error}"})

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

    remember_generated(output_path)

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
    payload = read_json()
    path = payload.get("path") or default_output_dir()
    # Only reveal somewhere this app has actually written, or the default
    # output folder. Handing an arbitrary path to the file manager is not a
    # thing the GUI ever needs.
    allowed = (
        is_generated(path)
        or any(os.path.dirname(f) == os.path.abspath(path)
               for f in GENERATED_FILES)
        or os.path.abspath(path) == os.path.abspath(default_output_dir())
    )
    if not allowed:
        return deny("That folder is not one this app has written to.")
    if not os.path.exists(path):
        return jsonify({"ok": False, "error": "That path no longer exists."})
    try:
        reveal_in_file_manager(path)
    except Exception as error:
        return jsonify({"ok": False, "error": str(error)})
    return jsonify({"ok": True})


@app.route("/api/download", methods=["POST"])
def api_download():
    payload = read_json()
    path = payload.get("path")
    if not path or not is_generated(path) or not os.path.isfile(path):
        return jsonify({"ok": False, "error": "File not found."}), 404
    return jsonify({
        "ok": True,
        "url": "/api/file?path=" + quote(os.path.abspath(path), safe=""),
    })


@app.route("/api/file")
def api_file():
    # Reached by a top-level navigation, which cannot carry the CSRF header, so
    # the allowlist is what protects it: only templates generated by this
    # process are readable, never an arbitrary path off disk.
    path = request.args.get("path", "")
    if not path or not is_generated(path) or not os.path.isfile(path):
        return "Not found", 404
    return send_file(os.path.abspath(path), as_attachment=True)


@app.route("/api/config", methods=["POST"])
def api_save_config():
    payload = read_json()
    try:
        current = load_config()
    except (OSError, ValueError) as error:
        return jsonify({"ok": False, "error": f"Cannot read config: {error}"})
    config = conform_to(payload.get("config") or {}, current)
    try:
        save_config(config)
    except OSError as error:
        return jsonify({"ok": False, "error": str(error)})
    return jsonify({"ok": True, "path": CONFIG_PATH})


@app.route("/api/preset", methods=["POST"])
def api_preset():
    payload = read_json()
    name = payload.get("name")
    current = payload.get("config") or load_config()
    presets = load_presets()
    if name not in presets:
        return jsonify({"ok": False, "error": "Unknown preset."})
    return jsonify({"ok": True, "config": deep_merge(current, presets[name])})


APP_URL = f"http://{HOST}:{PORT}"


def open_browser():
    webbrowser.open(APP_URL)


if __name__ == "__main__":
    if os.environ.get("WERKZEUG_RUN_MAIN") != "true":
        threading.Timer(1.2, open_browser).start()
    print(f"\n  Template Studio running at {APP_URL}")
    print("  Press Ctrl+C to stop.\n")
    app.run(host=HOST, port=PORT, debug=False)
