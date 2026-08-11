# Template Studio

A local web GUI that turns styling settings into a real Word template (`.dotx`).

Set your fonts, sizes, colours, spacing and bullets in the browser, watch the live
preview update, pick a folder, click Generate. The template is written straight to
disk and you can open the containing folder from the same screen.

---

## Setup

### Windows

Double click `run.bat`. It checks for Python, installs the two dependencies on
first run, starts the server and opens your browser.

### macOS and Linux

```bash
./run.sh
```

On first run it creates a virtual environment in `.venv` inside the project
folder, installs the dependencies there and starts the server. Nothing is
installed into your system Python. Later runs reuse the same environment and
start immediately.

Delete the `.venv` folder to force a clean reinstall.

### Manual

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

The browser opens at `http://127.0.0.1:5000`. Nothing leaves your machine and
nothing is sent over the internet.

Requires Python 3.9 or newer.

---

## Using it

1. **Pick a preset** from the dropdown at the top right, or start from the
   current settings.
2. **Work through the sections** using the navigation on the left. Every control
   updates the live preview on the right immediately.
3. **Set the save folder** in the bar at the bottom. Click `Browse` for a native
   folder dialog, or type a path directly.
4. **Set the file name**. The `.dotx` extension is added automatically.
5. **Click Generate template.** The status line shows the full path.
6. **Click Open folder** to reveal the file in Explorer or Finder.

Your settings are saved to `config.json` each time you generate, so the next
launch picks up where you left off. `Reload saved config` discards unsaved
changes on screen.

---

## What you can control

| Section | Controls |
|---|---|
| Template details | Name, description and author stored in the file properties |
| Document structure | Cover page, contents page, style guide, heading numbering, starter outline, cover placeholders, header text, page numbers |
| Heading 1 to 4 | Font, size, colour, bold, italic, all caps, space before, space after |
| Body text | Font, size, colour, bold, italic, spacing, line spacing |
| Caption | Same set, for figure and table captions |
| Quote block | Same set plus left indent |
| Inline text styles | Bold, italic and highlight colours used inside paragraphs |
| Bullets level 1 and 2 | Bullet glyph, font, size, colour, indent, spacing |
| Table of contents | Heading text and styling, levels shown, entry font, size, colour and spacing |
| Page setup | Page size, orientation, default font and size, four margins |

### Heading numbering

Turn on **Number the headings** and Word numbers Heading 1 to 4 automatically
using a multilevel list attached to the styles. You never type a number. Word
renumbers everything as you add, move or delete sections, and the numbers carry
through into the contents list.

Available formats:

| Setting | Result |
|---|---|
| `1, 1.1, 1.1.1` | 1, then 1.1, then 1.1.1, then 1.1.1.1 |
| `1., 1.1., 1.1.1.` | Same with a trailing dot at every level |
| `1.0, 1.1, 1.1.1` | Top level shows as 1.0, common in specifications |
| `Chapter 1, 1.1, 1.1.1` | Word "Chapter" before the top level number |
| `Section 1, 1.1, 1.1.1` | Word "Section" before the top level number |
| `I., A., 1., a.` | Classic outline, each level shows only its own counter |

**Levels numbered** stops numbering below a chosen level, so you can number
Heading 1 and 2 while leaving Heading 3 and 4 plain.

**Gap after number** is the separator between the number and the heading text,
either a tab, a single space, or nothing.

**Indent by level** steps each level in by a quarter inch with a hanging indent.
Leave it off to keep every heading flush with the left margin.

### Starter outlines

Choosing an outline type adds empty Heading 1 sections so you can start writing
straight away:

- **Proposal** — executive summary through to commercial terms
- **Report** — introduction, method, results, discussion, conclusions
- **Software usage guide** — installation, getting started, feature reference,
  troubleshooting
- **Study guide** — objectives, key concepts, worked examples, practice questions
- **Technical specification** — scope, definitions, functional and non functional
  requirements, interfaces, verification

---

## Units

| Unit | Where | Conversion |
|---|---|---|
| Points (pt) | Font sizes | 1 pt = 1/72 inch |
| Twips | Spacing, indents, margins | 1440 twips = 1 inch, 567 twips = 1 cm, 20 twips = 1 pt |
| Lines | Line spacing | 1.15 is the common default, 1.5 for double spaced feel |

Useful spacing values: 120 twips is a small gap, 240 is one line at 12 pt, 360 is
generous.

---

## Using the generated template in Word

Double clicking a `.dotx` opens a **new document based on it** rather than
editing the template. That is the point of a template. To edit the template
itself, open Word first, then File, Open, and select the `.dotx`.

To install it so it shows up under File, New:

- **Windows:** copy it to
  `%APPDATA%\Microsoft\Templates`
- **macOS:** copy it to
  `~/Library/Group Containers/UBF8T346G9.Office/User Content/Templates`

### The table of contents

The contents page holds a live Word field. Word is told to update fields on
open, so most of the time it fills in by itself. If it does not, right click the
contents list and choose Update Field, then Update entire table.

The list picks up anything styled Heading 1 to Heading 4.

### Applying styles

Open the Styles panel with `Ctrl+Alt+Shift+S` on Windows or
`Cmd+Alt+Shift+S` on macOS. The styles you configured appear as
Heading 1 to 4, Body Text, Caption, Quote, List Bullet and List Bullet 2.

Keyboard shortcuts: `Ctrl+Alt+1`, `Ctrl+Alt+2`, `Ctrl+Alt+3` for the first three
heading levels.

---

## File map

```
dotx_studio/
├── app.py                 Flask server, field schema, folder dialog, routes
├── template_builder.py    Config to .dotx engine
├── config.json            Current settings, rewritten on each generate
├── presets.json           Named style sets shown in the preset dropdown
├── templates/
│   └── index.html         The GUI, one file, no build step
├── requirements.txt
├── run.bat                Windows launcher
├── run.sh                 macOS and Linux launcher, sets up .venv
└── .venv/                 Created on first run by run.sh, safe to delete
```

---

## Extending it

### Add a preset

Add an entry to `presets.json`. It only needs the keys you want to override, and
they are merged over the current settings:

```json
"My Brand": {
  "styles": {
    "headings": {
      "heading_1": {"font": "Arial", "fontSize": 24, "color": "8B0000"}
    }
  }
}
```

It appears in the dropdown the next time you reload the page.

### Add a control

Every control in the GUI comes from the `SCHEMA` list in `app.py`. Add a field
and it appears in the form, wired to the config and to the preview. For example,
to expose Heading 1 underlining:

```python
{"path": "styles.headings.heading_1.underline", "label": "Underline",
 "type": "bool"},
```

Field types are `text`, `number`, `color`, `bool` and `select`. Number fields
accept `min`, `max`, `step` and `unit`. Select fields take `options` as a list of
strings, or a list of `[value, label]` pairs.

Then read the new key in `template_builder.py`.

### Add a starter outline

Add a list to the `outlines` dictionary in `_add_outline`, then add the matching
option to the `document.outline_type` field in `app.py`.

---

## Generating without the GUI

The engine works on its own, which is useful in scripts and build pipelines:

```python
import json
from template_builder import build_template

config = json.load(open("config.json"))
config["styles"]["headings"]["heading_1"]["color"] = "8B0000"

build_template(config, "output/Report_Template.dotx")
```

---

## Notes

- The generated file is a true template. The main document part declares the
  Word template content type, so Word treats it as a template rather than a
  document with a renamed extension.
- Built in Word styles bind their fonts and colours to the document theme. The
  builder strips those theme references before writing your values, otherwise
  the theme would override them.
- Fonts must be installed on the machine where the document is opened. Sticking
  to fonts that ship with Office avoids substitution.

---

## Troubleshooting

**The Browse button says the dialog is not available.**
The folder dialog needs `tkinter`. On Debian and Ubuntu install it with
`sudo apt install python3-tk`. On Windows and macOS it ships with Python. You
can always type the path into the box instead.

**Port 5000 is already in use.**
On macOS this is usually AirPlay Receiver. Turn it off in System Settings, or
change the port on the last line of `app.py`.

**Colours look wrong in Word.**
Check the hex value is six characters with no `#`. The colour swatch keeps this
correct for you.

**A bullet shows as a hollow box.**
The chosen font does not contain that glyph. Pick a different bullet, or set the
bullet font to one that has it.
