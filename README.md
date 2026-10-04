# Template Studio

**v0.2**

A local web GUI that turns styling settings into a real Word template (`.dotx`).

Set your fonts, sizes, colours, spacing and bullets in the browser, watch the live
preview update, pick a folder, click Generate. The template is written straight to
disk and you can open the containing folder from the same screen.

New in v0.2: load an existing `.dotx` or `.docx` to read its styling back into
the form and restyle it, and a Feedback page that writes a report you can email
back. See [CHANGELOG.md](CHANGELOG.md) for the detail, and
[HOW_TO_USE.txt](HOW_TO_USE.txt) for the version written for people who are not
going to read this file.

---

## Getting it running

A Linux or macOS machine with Python 3.9 or newer and git. Nothing else, and
nothing is installed system-wide.

```bash
git clone git@github.com:obaidukhalid/dotx-template-generator.git
cd dotx-template-generator
git checkout develop
./run.sh
```

`develop` is where the work is. `master` holds only the README.

`run.sh` does everything: it creates a virtual environment in `.venv` inside the
project folder, installs the pinned dependencies into it, starts the server and
opens your browser at `http://127.0.0.1:5000`. Later runs reuse the same
environment and start straight away.

If the script is not executable after the clone, `chmod +x run.sh` first.

To stop the server, press `Ctrl+C` in the terminal it is running in.

Nothing leaves your machine. The server binds to `127.0.0.1` and the app makes
no outbound requests of any kind.

### If you would rather drive it yourself

`run.sh` prefers [uv](https://docs.astral.sh/uv/getting-started/installation/)
and falls back to venv and pip. Both land in the same place. To do it by hand,
pick whichever you have.

**With uv**

```bash
uv run python app.py
```

That is the whole thing. `uv run` reads `pyproject.toml`, installs the exact
versions recorded in `uv.lock` into `.venv`, creating it if needed, then starts
the server. No separate install step, no environment to activate. Use `uv sync`
to set the environment up without starting the server.

**With plain Python**

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

`requirements.txt` is exported from `uv.lock`, so it pins the same versions.
Dependencies are declared in `pyproject.toml` — that is the file to edit when
adding one. Regenerate the export afterwards with:

```bash
uv export --format requirements-txt --no-hashes --no-dev --no-emit-project -o requirements.txt
```

Delete `.venv` to force a clean reinstall.

### Checking it works

With the server running, in another terminal:

```bash
curl -s -o /dev/null -w '%{http_code}\n' http://127.0.0.1:5000/
```

`200` means the app is up. Then generate a template from the GUI and open the
result in Word or LibreOffice. The styles, the heading numbers and the contents
field are the three things worth looking at.

A quick build without the GUI at all:

```bash
.venv/bin/python -c "
import json
from template_builder import build_template
build_template(json.load(open('config.json')), '/tmp/test.dotx')
print('built /tmp/test.dotx')
"
```

### Missing tkinter

The `Browse` button opens a native folder dialog, which needs `tkinter`. It is
not installed with Python on most Linux distributions:

```bash
sudo apt install python3-tk        # Debian, Ubuntu
sudo dnf install python3-tkinter   # Fedora
```

Without it everything else still works — type the save path into the box, or
use the `Download` button instead.

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
6. **Click Open folder** to reveal the file in your file manager.

Your settings are saved to `config.json` each time you generate, so the next
launch picks up where you left off. `Reload saved config` discards unsaved
changes on screen.

---

## Updating a template you already have

`Load existing template…` in the header takes a `.dotx` or `.docx`, reads the
styling out of it and fills in every control on the page. Adjust what you want
and generate. The file name is pre-filled as `<name>_restyled.dotx`, and the
file you loaded is never written to.

The import panel that appears lists what was read out of your file and what was
left at the settings already on screen. Read it — it is the only honest account
of what came across.

### What carries over

Heading 1 to 4, body text, caption, quote block, both bullet levels (including
the glyph and indent), the contents list heading and entries, page size,
orientation, margins, the default font, heading numbering (scheme, number of
levels, the gap after the number, whether levels indent), whether the template
has a cover page, contents page or style guide, which starter outline it used,
the cover placeholder text, the page header text, whether the footer carries a
page number, and the template name, description and author.

### What does not

Anything the config has no field for: custom styles, Heading 5 and beyond,
table styles, numbered-list styles other than the two bullet levels, images,
and the document's own text.

This is a deliberate limit, not an oversight. `Generate` always builds a fresh
template from the settings on screen through the same `build_template` path as
any other template; it does not edit the package you uploaded, so the parts it
cannot describe are not preserved. If you need those, keep working from the
original file in Word.

Inline text styles are also not read back, because the builder applies them as
direct run formatting in the style guide rather than as character styles, so
there is nothing in the file to read them from.

The old binary `.doc` and `.dot` formats cannot be read at all. Open them in
Word and save as `.dotx` first.

---

## Feedback

`Feedback` in the header opens a short form: six ratings, one question about
which part needs the most work, and a free text box.

`Save feedback` writes `feedback/feedback_<timestamp>.txt` in the app folder.
Nothing is transmitted by the app. `Open email` then hands the file's contents
to your mail program, addressed to `contact@v-embed.com` with the subject
`Feedback - Dotx Studio v0.2`; `Copy to clipboard` and `Open feedback folder`
are there for when `mailto` is not wired up on the machine.

The questions are one list, `FEEDBACK_QUESTIONS` in `app.py`, and drive both
the page and the saved file. Add an entry and it appears in both.

`feedback/` is tracked so the app has somewhere to write; its contents are
gitignored.

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
├── app.py                 Flask server, field schema, feedback questions, routes
├── template_builder.py    Config to .dotx engine
├── template_reader.py     .dotx back to config, the inverse within limits
├── config.json            Current settings, rewritten on each generate
├── presets.json           Named style sets shown in the preset dropdown
├── templates/
│   ├── index.html         The builder GUI, one file, no build step
│   └── feedback.html      The feedback page, rendered from FEEDBACK_QUESTIONS
├── feedback/              Where saved feedback lands, contents gitignored
├── HOW_TO_USE.txt         Plain text instructions, for testers who want them
├── CHANGELOG.md           What changed in each version
├── pyproject.toml         Project metadata and dependency declarations
├── uv.lock                Exact resolved versions, committed for reproducibility
├── requirements.txt       Generated from uv.lock for the non-uv path
├── run.sh                 The launcher: sets up .venv and starts the server
├── run.bat                The same for Windows, unmaintained
└── .venv/                 Created on first run, safe to delete
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

### Why the server is fussy about who is calling

The server listens on loopback only, but that on its own guards nothing: any
page open in the same browser can send requests to `127.0.0.1`, and a hostile
DNS record can point a domain at loopback so the attacker's scripts read the
replies too. Three checks close that off, and they are worth knowing about if
you ever call the API yourself.

- Every request that changes something must carry an `X-CSRF-Token` header. The
  token is generated per run and handed only to the real page, so another site
  cannot read it. The GUI attaches it automatically.
- The `Host` header has to match the address the app is served on. This is what
  stops DNS rebinding, since the browser sends the attacker's hostname there.
- Bodies are parsed as strict JSON. A plain form POST — the one request shape
  that crosses origins without a preflight — is refused before it reaches a
  handler, so `Content-Type: application/json` is required.

`/api/import` is the single exception to the JSON rule, because a file upload
has to be multipart. The CSRF check still covers it: a cross-origin form post
cannot set `X-CSRF-Token`, and setting it is what forces a preflight this server
never answers. Uploads are capped at 16 MB.

Two consequences for the file endpoints. `/api/file` only serves templates that
this run of the app generated, never an arbitrary path, so it cannot be turned
into a way to read your disk. `/api/open-folder` likewise only opens folders the
app has written to, or the default output folder.

None of this authenticates *local* programs — anything already running as you
can read the page and take the token. It is a defence against other websites,
not against software you have already installed.

---

## Troubleshooting

**The Browse button says the dialog is not available.**
`tkinter` is missing — see [Missing tkinter](#missing-tkinter) above. Type the
path into the box instead, or use `Download`.

**Port 5000 is already in use.**
Find what has it with `ss -ltnp | grep :5000` and stop that, or change the
`PORT` constant near the top of `app.py`. The server checks incoming requests
against that value, so changing it in one place is enough. On macOS the culprit
is usually AirPlay Receiver, in System Settings.

**Colours look wrong in Word.**
Check the hex value is six characters with no `#`. The colour swatch keeps this
correct for you.

**A bullet shows as a hollow box.**
The chosen font does not contain that glyph. Pick a different bullet, or set the
bullet font to one that has it.

**Loading a template left most settings unchanged.**
Check the import panel's second column. A file whose styles bind their fonts
and colours to the document theme rather than setting them explicitly has
nothing for the reader to pick up, so those fields keep the values that were
already loaded. Files this tool generated itself round trip completely, because
the builder writes explicit values and strips the theme references.

**Loading a template says it cannot be read.**
Only `.dotx` and `.docx` work. The binary `.doc` and `.dot` formats are a
different file format entirely — open them in Word and save as `.dotx`.

**The Open email button does nothing.**
Nothing is registered to handle `mailto:` on that machine. The feedback is
already saved as a file; use `Open feedback folder` and attach it to an email
yourself, to `contact@v-embed.com` with the subject
`Feedback - Dotx Studio v0.2`.
