# Template Studio v0.2

Builds Word template files (`.dotx`). You pick fonts, sizes, colours and spacing
in your browser and it writes the template to disk.

Everything runs on your own machine. Nothing is uploaded.

---

## 1. Requirements

- Linux or macOS
- Python 3.9 or newer
- git
- A browser (Firefox, Chrome or Edge)

---

## 2. Run it

```bash
git clone git@github.com:obaidukhalid/dotx-template-generator.git
cd dotx-template-generator
./run.sh
```

That is all. `run.sh` installs what it needs into a `.venv` folder inside the
project, starts the server, and opens `http://127.0.0.1:5000`.

- Later runs: `./run.sh` again. Starts straight away.
- Stop it: `Ctrl+C`.
- Permission denied? Run `chmod +x run.sh` once.
- No `tkinter`? Install it: `sudo apt install python3-tk` (Debian, Ubuntu) or
  `sudo dnf install python3-tkinter` (Fedora). Only the `Browse` button needs
  it.

### Without run.sh

With [uv](https://docs.astral.sh/uv/getting-started/installation/):

```bash
uv run python app.py
```

With pip:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

---

## 2b. Build the Windows .exe

Only on this branch (`feature/windows-exe`). `master` and `develop` do not
carry it.

**Must be run on Windows.** PyInstaller bundles the interpreter of the machine
it runs on, so a Linux clone cannot produce a Windows exe.

On a Windows machine with Python installed:

```powershell
git clone git@github.com:obaidukhalid/dotx-template-generator.git
cd dotx-template-generator
git checkout feature/windows-exe
.\build_windows_exe.ps1
```

Result: `dist\TemplateStudio.exe`, about 18 MB. It needs no Python on whatever
machine runs it.

- The script makes its own `.build-venv`, so your `.venv` is untouched.
- PowerShell may block the script. Run it as
  `powershell -File .\build_windows_exe.ps1`.
- Put the exe in a writable folder. `config.json`, `presets.json` and
  `feedback/` are written **next to the exe**.
- It is not code-signed, so SmartScreen will warn: More info → Run anyway.

`TemplateStudio.spec` is the build definition. The `FROZEN` section in `app.py`
is what makes the exe find its files and keep its settings.

---

## 3. Make a template

1. Choose a preset from the dropdown, or start from the current settings.
2. Edit the settings. Use the list on the left to jump between sections.
3. Watch the preview on the right update as you type.
4. Set the save folder and file name at the bottom.
5. Click **Generate template**.
6. Click **Open folder** or **Download** to get the file.

Your settings are saved to `config.json` when you generate, so the next run
picks up where you left off. **Reload saved config** throws away unsaved edits.

---

## 4. Restyle a template you already have

1. Click **Load existing template…** at the top.
2. Pick a `.dotx` or `.docx`. Every setting fills in from that file.
3. Check the green panel. It lists what was read and what was not.
4. Edit what you want.
5. Click **Generate template**.

Your original file is never changed. The new file is named
`<name>_restyled.dotx` by default.

### What you can change

Heading 1 to 4, body text, caption, quote block, both bullet levels, the
contents list, page size, orientation, margins, the default font, heading
numbering, cover and header text, and the file's name, description and author.

### What you cannot

Anything the tool has no setting for is dropped:

- Custom styles you made yourself
- Heading 5 and below
- Table styles
- Images
- The document's text

**Generate** always builds a fresh template from the settings on screen. It does
not edit your file. Keep the original until you are happy with the new one.

`.doc` and `.dot` cannot be read at all. Open them in Word and save as `.dotx`
first.

---

## 5. Settings reference

| Section | What it controls |
|---|---|
| Template details | Name, description and author stored in the file |
| Document structure | Cover page, contents page, style guide, heading numbering, starter outline, header text, page numbers |
| Heading 1 to 4 | Font, size, colour, bold, italic, all caps, spacing |
| Body text | Font, size, colour, bold, italic, spacing, line spacing |
| Caption | Same, for figure and table captions |
| Quote block | Same, plus left indent |
| Inline text styles | Bold, italic and highlight colours used inside paragraphs |
| Bullets level 1 and 2 | Glyph, font, size, colour, indent, spacing |
| Table of contents | Heading text and style, levels shown, entry style |
| Page setup | Page size, orientation, default font and size, margins |

### Heading numbering

Turn on **Number the headings** and Word numbers Heading 1 to 4 for you. You
never type a number. Word renumbers as you add, move or delete sections, and the
numbers appear in the contents list.

| Format | Looks like |
|---|---|
| `1, 1.1, 1.1.1` | 1, 1.1, 1.1.1, 1.1.1.1 |
| `1., 1.1., 1.1.1.` | Same, with a trailing dot |
| `1.0, 1.1, 1.1.1` | Top level shows as 1.0 |
| `Chapter 1, 1.1, 1.1.1` | "Chapter" before the top number |
| `Section 1, 1.1, 1.1.1` | "Section" before the top number |
| `I., A., 1., a.` | Classic outline |

- **Levels numbered** — stop numbering below a level, e.g. number Heading 1 and
  2 only.
- **Gap after number** — tab, space or nothing.
- **Indent by level** — step each level in. Off keeps headings flush left.

### Starter outlines

Adds empty Heading 1 sections so you can start writing:

- **Proposal** — executive summary to commercial terms
- **Report** — introduction, method, results, discussion, conclusions
- **Software usage guide** — installation, getting started, feature reference
- **Study guide** — objectives, key concepts, examples, practice questions
- **Technical specification** — scope, definitions, requirements, interfaces

### Units

| Unit | Used for | Conversion |
|---|---|---|
| Points (pt) | Font sizes | 1 pt = 1/72 inch |
| Twips | Spacing, indents, margins | 1440 = 1 inch, 567 = 1 cm, 20 = 1 pt |
| Lines | Line spacing | 1.15 is a common default |

Handy twips values: 120 small gap, 240 one line at 12 pt, 360 generous.

---

## 6. Send feedback

1. Click **Feedback** at the top of the page.
2. Answer the questions.
3. Click **Save feedback**. It writes a text file to `feedback/`.
4. Click **Open email** and send it.

Nothing is sent until you send it. If `mailto:` does not work on your machine,
use **Copy to clipboard** or **Open feedback folder** and email the file to
`contact@v-embed.com` with the subject `Feedback - Dotx Studio v0.2`.

---

## 7. Use the template in Word

- Double click the `.dotx` — Word opens a **new document** based on it. That is
  what a template does.
- To edit the template itself: open Word first, then File → Open.
- To get it under File → New, copy it to:
  - Windows: `%APPDATA%\Microsoft\Templates`
  - macOS: `~/Library/Group Containers/UBF8T346G9.Office/User Content/Templates`

**Contents page.** It is a live Word field and usually fills in on open. If not,
right click it → Update Field → Update entire table. It picks up Heading 1 to 4.

**Styles panel.** `Ctrl+Alt+Shift+S` (Windows) or `Cmd+Alt+Shift+S` (macOS).
Your styles appear as Heading 1 to 4, Body Text, Caption, Quote, List Bullet and
List Bullet 2. Shortcuts `Ctrl+Alt+1/2/3` apply the first three heading levels.

---

## 8. Use it from a script

The engine runs without the GUI, for build pipelines or batch jobs:

```python
import json
from template_builder import build_template

config = json.load(open("config.json"))
config["styles"]["headings"]["heading_1"]["color"] = "8B0000"

build_template(config, "output/Report_Template.dotx")
```

---

## 9. Customising

**Add a preset** — add an entry to `presets.json`. Only the keys you want to
override; they merge over the current settings.

```json
"My Brand": {
  "styles": {
    "headings": {
      "heading_1": {"font": "Arial", "fontSize": 24, "color": "8B0000"}
    }
  }
}
```

**Add a control** — add a field to the `SCHEMA` list in `app.py`. It appears in
the form, wired to the config and the preview.

```python
{"path": "styles.headings.heading_1.underline", "label": "Underline",
 "type": "bool"},
```

Types: `text`, `number`, `color`, `bool`, `select`. Then read the new key in
`template_builder.py`.

**Add a starter outline** — add a list to `outlines` in `_add_outline`, then add
the matching option to `document.outline_type` in `app.py`.

**Add a feedback question** — add an entry to `FEEDBACK_QUESTIONS` in `app.py`.
It appears on the page and in the saved file.

---

## 10. Files

```
app.py                 Server, settings schema, feedback questions, routes
template_builder.py    Settings  ->  .dotx
template_reader.py     .dotx  ->  settings
config.json            Current settings, rewritten on each generate
presets.json           Named style sets for the dropdown
templates/index.html   The main page
templates/feedback.html The feedback page
feedback/              Saved feedback, contents gitignored
run.sh                 Sets up .venv and starts the server
pyproject.toml         Dependencies
uv.lock                Pinned versions
requirements.txt       Exported from uv.lock for the pip path
HOW_TO_USE.txt         Plain text instructions for testers
CHANGELOG.md           What changed in each version
```

---

## 11. Troubleshooting

| Problem | Fix |
|---|---|
| Browse button says the dialog is unavailable | `tkinter` is missing. Install it, or type the path in the box. |
| Port 5000 already in use | `ss -ltnp \| grep :5000` to find it. Or change `PORT` in `app.py`. On macOS it is usually AirPlay Receiver. |
| Browser does not open | Open `http://127.0.0.1:5000` yourself. |
| Colours look wrong in Word | Hex must be six characters, no `#`. |
| A bullet shows as a hollow box | That font lacks the glyph. Pick another bullet or font. |
| Loading a template changed almost nothing | Its styles use the Word theme instead of explicit values, so there is nothing to read. Check the green panel. |
| Loading a template fails | Only `.dotx` and `.docx` work. Save `.doc`/`.dot` as `.dotx` in Word first. |
| Open email button does nothing | No `mailto:` handler. Attach the saved file from `feedback/` yourself. |
| Buttons stop working | Reload the page. The page gets a one-time token at load. |

---

## 12. Notes

- The output is a real template. The file declares the Word template content
  type, not just a renamed `.docx`.
- Word's built-in styles tie fonts and colours to the document theme. The
  builder strips those references first, or the theme would win.
- Fonts must be installed on whatever machine opens the document.
