# Changelog

## v0.2

### Load an existing template and restyle it

A new **Load existing template…** button at the top of the builder takes a
`.dotx` or `.docx` and reads its styling back into every setting on the page,
so an existing template can be adjusted rather than retyped from scratch.

A new module, `template_reader.py`, does the reading. It is the inverse of
`template_builder.py` within the limits of the config: it recovers Heading 1
to 4, body text, caption, quote, both bullet levels, the contents list heading
and entries, page size, orientation, margins, the default font, the heading
numbering scheme, which optional blocks the template contains, the cover
placeholder text, the page header, the footer page number and the file
metadata.

It does **not** recover anything the config has no slot for — custom styles,
Heading 5 and beyond, table styles, images, or the document's own text.
Generate therefore writes a fresh template from the settings on screen rather
than editing the uploaded file, and the uploaded file is never modified. The
import panel lists, field group by field group, what came out of the file and
what was left at the settings already loaded, so there is no guessing about
what carried over.

Supporting changes:

- `POST /api/import` reads an uploaded package. It is the one handler that
  takes a multipart body rather than strict JSON; the CSRF token check still
  covers it, since a cross-origin form post cannot set the header. Uploads are
  capped at 16 MB and a 413 is answered in the JSON shape the GUI expects.
- A `.dotx` declares a template content type that python-docx refuses to open,
  so the reader relabels a temporary copy before parsing it.
- Select fields now keep a value that is not on their option list — a font or
  a bullet glyph out of a loaded file is preserved and marked `(from file)`
  instead of being snapped to the first option.
- The old binary `.doc` and `.dot` formats are rejected with an explanation
  rather than a traceback, as are files that are not Word packages at all.

### Feedback page

A new **Feedback** link in the header opens `/feedback`: six rating questions
about setup, the interface, the live preview, the generated template and the
new import, one question about which part needs the most work, and a free text
box.

- **Save feedback** writes a plain text file to `feedback/` in the app folder.
  Nothing is sent anywhere by the app itself.
- **Open email** then opens the user's mail program addressed to
  `contact@v-embed.com` with the subject `Feedback - Dotx Studio v0.2` and the
  feedback in the body. Long feedback is left out of the `mailto` link, which
  mail programs truncate, and the page says to attach the saved file instead.
- **Copy to clipboard** and **Open feedback folder** cover the cases where
  `mailto` does not work.
- `POST /api/feedback` validates every answer against the question definitions
  before anything is written: a rating must be one of the offered numbers, a
  choice one of the offered values, and free text is stripped of control
  characters and capped.
- The questions live in one `FEEDBACK_QUESTIONS` list in `app.py` and drive
  both the page and the saved file, the same way `SCHEMA` drives the builder.
- `feedback/` is tracked but its contents are not. Feedback is the user's.

### Runs from a clone

Clone the repo and run `./run.sh`. That is the whole setup. Linux and macOS,
Python 3.9 or newer.

Windows support was dropped: `run.bat` and the executable build tooling are
gone from the repo, and `app.py` no longer carries the path handling a frozen
build needed.

A port clash now prints an explanation instead of a traceback.

### Documentation

- `HOW_TO_USE.txt` in the project root: how to start the tool, how to build a
  template, what the import does and does not carry over, how to use the
  result in Word, how to send feedback to `contact@v-embed.com` with the
  subject `Feedback - Dotx Studio v0.2`, and what to do when something fails.
- This changelog.
- The version is shown in the app header, printed on startup, returned from
  `/api/bootstrap` and set in `pyproject.toml`.

### Repository

`develop` was created as the permanent working line, branched so that `master`
is a genuine ancestor of it. `master` stays the README-only production line.

`tests-initial` was deleted, locally and on the remote, once every commit on it
was confirmed present in `develop`.

The Windows build tooling is gitignored, not tracked, so a clone holds only
what a Python environment needs. `master` carries the running code, since that
is the branch testers clone.

## v0.1

First version. A local Flask GUI driven by a field schema, a live preview, and
`template_builder.py` writing a real `.dotx`: styles for Heading 1 to 4, body,
caption, quote, two bullet levels and the contents list; automatic multilevel
heading numbering in six schemes; page setup; an optional cover page, contents
page, self-documenting style guide and starter outline for five document
types; presets; and a loopback server closed off with a Host allowlist, an
Origin check, a per-session CSRF token and an allowlist on the file-serving
routes.
