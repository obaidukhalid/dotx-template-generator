"""
template_reader.py
------------------
Reads an existing Word template (.dotx) or document (.docx) back into a styling
configuration dictionary of the same shape that template_builder.py consumes.

This is the inverse of build_template, within limits. The config describes a
fixed set of styles, so only those are recovered: Heading 1 to 4, body text,
caption, quote, two bullet levels, the contents list, page setup, heading
numbering and the file metadata. Anything else the source file happens to
contain — custom styles, Heading 5 and beyond, table styles, the body text
itself — is not represented in the config and is therefore dropped. Loading a
file and generating from it produces a fresh template in this tool's own shape,
not a copy of the original.

Public entry point:
    read_template(source, defaults) -> (config, report)
"""

import copy
import os
import re
import shutil
import tempfile
import zipfile

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from template_builder import (
    CUSTOM_FIELD_KEYS,
    STUDIO_META_KEYS,
    DOCX_MAIN_CT,
    DOTX_MAIN_CT,
    HEADING_KEYS,
    HEADING_NUMBER_FORMATS,
    PAGE_SIZES,
    PLACEABLE_KEYS,
    STUDIO_NS,
)

# Sections of a cover page, in the order _add_cover writes them.
COVER_FIELDS = [
    "title_placeholder",
    "subtitle_placeholder",
    "author_placeholder",
    "date_placeholder",
    "reference_placeholder",
]

# Heading 1 texts the starter outlines use, so an outline can be recognised.
# Keep in step with the outlines dictionary in template_builder._add_outline.
OUTLINE_MARKERS = {
    "proposal": "Scope of Work",
    "report": "Conclusions and Recommendations",
    "guide": "Feature Reference",
    "study_guide": "Learning Objectives",
    "specification": "Verification and Validation",
}

STYLE_GUIDE_MARKER = "Style Guide"

# Page dimensions are stored in EMU and a round trip through Word can shift
# them by a hair, so sizes are matched with a tolerance rather than exactly.
# 10000 EMU is about 0.01 mm.
PAGE_SIZE_TOLERANCE = 10000


class TemplateReadError(Exception):
    """Raised when the file is not a Word package this tool can read."""


# ----------------------------------------------------------------------------
# Package handling
# ----------------------------------------------------------------------------

def _as_readable_docx(source):
    """
    Write `source` to a temporary .docx and return its path.

    python-docx refuses a package whose main part declares the template content
    type, so a .dotx has to be relabelled before it can be opened. The copy is
    temporary and the caller deletes it; the original file is never touched.
    """
    temp_fd, temp_path = tempfile.mkstemp(suffix=".docx")
    os.close(temp_fd)

    if isinstance(source, (bytes, bytearray)):
        with open(temp_path, "wb") as handle:
            handle.write(source)
    else:
        shutil.copyfile(source, temp_path)

    try:
        with zipfile.ZipFile(temp_path, "r") as archive:
            items = archive.infolist()
            if not any(i.filename == "word/document.xml" for i in items):
                raise TemplateReadError(
                    "That file is not a Word template. Choose a .dotx or "
                    ".docx file."
                )
            relabelled = tempfile.mkstemp(suffix=".docx")
            os.close(relabelled[0])
            with zipfile.ZipFile(relabelled[1], "w", zipfile.ZIP_DEFLATED) as out:
                for item in items:
                    data = archive.read(item.filename)
                    if item.filename == "[Content_Types].xml":
                        text = data.decode("utf-8")
                        text = text.replace(DOTX_MAIN_CT, DOCX_MAIN_CT)
                        data = text.encode("utf-8")
                    out.writestr(item, data)
    except zipfile.BadZipFile:
        os.unlink(temp_path)
        raise TemplateReadError(
            "That file is not a Word template. The old binary .doc and .dot "
            "formats cannot be read; save as .dotx first."
        )
    except TemplateReadError:
        os.unlink(temp_path)
        raise

    os.unlink(temp_path)
    shutil.move(relabelled[1], temp_path)
    return temp_path


# ----------------------------------------------------------------------------
# Reading individual values off a style
# ----------------------------------------------------------------------------

def _style(doc, name):
    """Return a style by display name, or None when the file has no such style."""
    try:
        return doc.styles[name]
    except KeyError:
        return None


def _run_props(style):
    return style.element.find(qn("w:rPr"))


def _font_of(style):
    """The explicit ascii font of a style, ignoring theme references."""
    rpr = _run_props(style)
    if rpr is None:
        return None
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        return None
    return rfonts.get(qn("w:ascii")) or rfonts.get(qn("w:hAnsi"))


def _color_of(style):
    """The explicit colour of a style as six uppercase hex digits."""
    rpr = _run_props(style)
    if rpr is None:
        return None
    color = rpr.find(qn("w:color"))
    if color is None:
        return None
    value = color.get(qn("w:val"))
    if not value or value.lower() == "auto":
        return None
    value = value.upper().lstrip("#")
    return value if re.fullmatch(r"[0-9A-F]{6}", value) else None


def _size_of(style):
    """Font size in points, rounded to the half point the GUI steps in."""
    size = style.font.size
    if size is None:
        return None
    return round(size.pt * 2) / 2


def _twips_of(length):
    return None if length is None else int(round(length.twips))


def _line_spacing_of(style):
    """
    Line spacing as a multiple.

    Word stores it either as a multiple or as an exact height in twips. Only
    the multiple maps onto the config field, so an exact height is reported as
    unreadable and the saved default is kept.
    """
    spacing = style.paragraph_format.line_spacing
    if spacing is None:
        return None
    if isinstance(spacing, float):
        return round(spacing, 2)
    return None


# ----------------------------------------------------------------------------
# Report
#
# Every field either comes out of the file or stays at the value already in the
# config. The GUI shows both lists so it is clear what the file actually
# carried and what is simply left over from before the import.
# ----------------------------------------------------------------------------

class Report:
    def __init__(self):
        self.imported = []
        self.defaulted = []

    def record(self, label, found):
        (self.imported if found else self.defaulted).append(label)
        return found

    def as_dict(self):
        return {"imported": self.imported, "defaulted": self.defaulted}


def _read_text_style(doc, style_name, target, label, report,
                     line_spacing=True, caps=False, indent=False):
    """
    Copy the standard paragraph style fields out of `style_name` into `target`.

    `target` arrives holding the current config values, so a field the style
    does not set explicitly keeps the value it already had.
    """
    style = _style(doc, style_name)
    if style is None:
        report.defaulted.append(label)
        return False

    readers = [
        ("font", _font_of(style)),
        ("fontSize", _size_of(style)),
        ("color", _color_of(style)),
        ("bold", style.font.bold),
        ("italic", style.font.italic),
        ("spacing_before", _twips_of(style.paragraph_format.space_before)),
        ("spacing_after", _twips_of(style.paragraph_format.space_after)),
    ]
    if caps:
        readers.append(("all_caps", style.font.all_caps))
    if line_spacing:
        readers.append(("line_spacing", _line_spacing_of(style)))
    if indent:
        readers.append(("indent", _twips_of(style.paragraph_format.left_indent)))

    found_any = False
    for key, value in readers:
        if value is None or key not in target:
            continue
        target[key] = value
        found_any = True

    report.record(label, found_any)
    return found_any


# ----------------------------------------------------------------------------
# Numbering
# ----------------------------------------------------------------------------

def _numbering_element(doc):
    try:
        return doc.part.numbering_part.element
    except (AttributeError, KeyError, NotImplementedError):
        return None


def _abstract_for_style(doc, style_name):
    """Resolve the abstract numbering definition a paragraph style points at."""
    numbering = _numbering_element(doc)
    if numbering is None:
        return None

    style = _style(doc, style_name)
    if style is None:
        return None

    ppr = style.element.find(qn("w:pPr"))
    if ppr is None:
        return None
    num_pr = ppr.find(qn("w:numPr"))
    if num_pr is None:
        return None
    num_id_el = num_pr.find(qn("w:numId"))
    if num_id_el is None:
        return None
    num_id = num_id_el.get(qn("w:val"))

    for num in numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) != num_id:
            continue
        ref = num.find(qn("w:abstractNumId"))
        if ref is None:
            return None
        abstract_id = ref.get(qn("w:val"))
        for abstract in numbering.findall(qn("w:abstractNum")):
            if abstract.get(qn("w:abstractNumId")) == abstract_id:
                return abstract
    return None


def _read_bullet(doc, style_name, target, label, report):
    """Read a bullet level: the paragraph style plus its glyph and indent."""
    style = _style(doc, style_name)
    if style is None:
        report.defaulted.append(label)
        return False

    found_any = False
    for key, value in (
        ("font", _font_of(style)),
        ("fontSize", _size_of(style)),
        ("color", _color_of(style)),
        ("spacing_after", _twips_of(style.paragraph_format.space_after)),
        ("line_spacing", _line_spacing_of(style)),
    ):
        if value is not None and key in target:
            target[key] = value
            found_any = True

    abstract = _abstract_for_style(doc, style_name)
    if abstract is not None:
        lvl = abstract.find(qn("w:lvl"))
        if lvl is not None:
            lvl_text = lvl.find(qn("w:lvlText"))
            if lvl_text is not None and lvl_text.get(qn("w:val")):
                target["bullet_char"] = lvl_text.get(qn("w:val"))
                found_any = True
            ppr = lvl.find(qn("w:pPr"))
            ind = None if ppr is None else ppr.find(qn("w:ind"))
            if ind is not None and ind.get(qn("w:left")) is not None:
                # _set_bullet writes left = indent + 360 for the hanging glyph,
                # so take that 360 back off to recover the configured indent.
                left = int(ind.get(qn("w:left")))
                target["indent"] = max(0, left - 360)
                found_any = True

    report.record(label, found_any)
    return found_any


def _read_heading_numbering(doc, document_cfg, report):
    """Recover the heading numbering scheme from the multilevel list."""
    abstract = _abstract_for_style(doc, "Heading 1")
    if abstract is None:
        document_cfg["number_headings"] = False
        report.record("Heading numbering", False)
        return

    levels = abstract.findall(qn("w:lvl"))[:4]
    if not levels:
        document_cfg["number_headings"] = False
        report.record("Heading numbering", False)
        return

    texts, formats, numbered = [], [], 0
    for lvl in levels:
        fmt_el = lvl.find(qn("w:numFmt"))
        text_el = lvl.find(qn("w:lvlText"))
        fmt = fmt_el.get(qn("w:val")) if fmt_el is not None else "none"
        text = text_el.get(qn("w:val")) if text_el is not None else ""
        formats.append(fmt)
        texts.append(text or "")
        if fmt != "none" and text:
            numbered += 1

    if numbered == 0:
        document_cfg["number_headings"] = False
        report.record("Heading numbering", False)
        return

    document_cfg["number_headings"] = True
    document_cfg["number_levels"] = numbered

    # Match the level texts against the known schemes. Only the levels that
    # are actually numbered take part, so "Heading 1 to 2" still identifies
    # the scheme it was cut down from.
    for name, scheme in HEADING_NUMBER_FORMATS.items():
        if all(
            texts[i] == scheme["texts"][i] and formats[i] == scheme["formats"][i]
            for i in range(numbered)
        ):
            document_cfg["number_format"] = name
            break

    suff = levels[0].find(qn("w:suff"))
    if suff is not None and suff.get(qn("w:val")) in ("tab", "space", "nothing"):
        document_cfg["number_suffix"] = suff.get(qn("w:val"))

    ppr = levels[0].find(qn("w:pPr"))
    ind = None if ppr is None else ppr.find(qn("w:ind"))
    if ind is not None:
        document_cfg["number_indent"] = ind.get(qn("w:left")) not in (None, "0")

    report.record("Heading numbering", True)


# ----------------------------------------------------------------------------
# Page setup
# ----------------------------------------------------------------------------

def _read_page_setup(doc, page_cfg, report):
    if not doc.sections:
        report.defaulted.append("Page setup")
        return
    section = doc.sections[0]

    width, height = section.page_width, section.page_height
    if width is not None and height is not None:
        landscape = width > height
        short, long = (height, width) if landscape else (width, height)
        for name, (size_w, size_h) in PAGE_SIZES.items():
            if (abs(int(short) - int(size_w)) <= PAGE_SIZE_TOLERANCE
                    and abs(int(long) - int(size_h)) <= PAGE_SIZE_TOLERANCE):
                page_cfg["page_size"] = name
                break
        page_cfg["orientation"] = "landscape" if landscape else "portrait"
        report.record("Page size and orientation", True)
    else:
        report.defaulted.append("Page size and orientation")

    margins = page_cfg.setdefault("margins", {})
    found = False
    for key, value in (
        ("top", section.top_margin),
        ("bottom", section.bottom_margin),
        ("left", section.left_margin),
        ("right", section.right_margin),
    ):
        twips = _twips_of(value)
        if twips is not None:
            margins[key] = twips
            found = True
    report.record("Margins", found)

    normal = _style(doc, "Normal")
    found = False
    if normal is not None:
        font = _font_of(normal)
        size = _size_of(normal)
        if font:
            page_cfg["default_font"] = font
            found = True
        if size:
            page_cfg["default_font_size"] = size
            found = True
    report.record("Default font", found)


# ----------------------------------------------------------------------------
# Document body: what the template contains
# ----------------------------------------------------------------------------

def _instr_texts(doc):
    """Every field instruction in the body, header and footer."""
    texts = []
    parts = [doc.element.body]
    for section in doc.sections:
        for container in (section.header, section.footer):
            try:
                parts.append(container._element)
            except AttributeError:
                continue
    for part in parts:
        for el in part.iter(qn("w:instrText")):
            if el.text:
                texts.append(el.text.strip())
    return texts


def _paragraph_text(paragraph):
    """
    All the text in a paragraph, including text inside content controls.

    python-docx's Paragraph.text only walks direct w:r children, so a cover
    line that is a bound control reads as empty and every placeholder after it
    shifts up one.
    """
    parts = []
    for node in paragraph._p.iter(qn("w:t")):
        if node.text:
            parts.append(node.text)
    return "".join(parts).strip()


def _has_custom_binding(paragraph):
    """True when the paragraph carries a control bound to one of our fields."""
    for binding in paragraph._p.iter(qn("w:dataBinding")):
        if "ts:custom_" in (binding.get(qn("w:xpath")) or ""):
            return True
    return False


def _read_document_structure(doc, document_cfg, toc_cfg, report):
    """
    Work out which of the optional blocks the template was built with.

    None of this is stored in the file as a setting, so it is read back off the
    content: a TOC field means there is a contents page, a "Style Guide"
    heading means the style guide is present, and the centred run of
    paragraphs at the top is the cover.
    """
    instructions = _instr_texts(doc)

    toc_field = next((t for t in instructions if t.upper().startswith("TOC")), None)
    document_cfg["include_toc"] = toc_field is not None
    if toc_field:
        match = re.search(r'\\o\s*"(\d-\d)"', toc_field)
        if match:
            toc_cfg["levels"] = match.group(1)
    report.record("Contents page", toc_field is not None)

    heading_1_texts = [
        p.text.strip() for p in doc.paragraphs
        if p.style is not None and p.style.name == "Heading 1" and p.text.strip()
    ]

    has_guide = any(t == STYLE_GUIDE_MARKER for t in heading_1_texts)
    document_cfg["include_style_guide"] = has_guide
    report.record("Style guide section", has_guide)

    document_cfg["outline_type"] = "none"
    for name, marker in OUTLINE_MARKERS.items():
        if marker in heading_1_texts:
            document_cfg["outline_type"] = name
            break
    report.record(
        "Starter outline", document_cfg["outline_type"] != "none"
    )

    # The cover is the opening run of centred paragraphs, before any heading.
    # The lines for custom fields sit at the end of it and are read from the
    # header instead, so they stop the run rather than shifting every
    # placeholder down by one.
    cover_lines = []
    for paragraph in doc.paragraphs:
        name = paragraph.style.name if paragraph.style is not None else ""
        if name.startswith("Heading") or name.startswith("TOC"):
            break
        if _has_custom_binding(paragraph):
            break
        text = _paragraph_text(paragraph)
        # alignment 1 is centred; python-docx exposes it as an enum that
        # compares equal to its integer value.
        if text and paragraph.alignment is not None and int(paragraph.alignment) == 1:
            cover_lines.append(text)
        if len(cover_lines) >= len(COVER_FIELDS):
            break

    document_cfg["include_cover"] = bool(cover_lines)
    for key, text in zip(COVER_FIELDS, cover_lines):
        document_cfg[key] = text
    report.record("Cover page", bool(cover_lines))




def _slot_items(paragraph):
    """
    Walk a header or footer paragraph and say what sits at each position.

    Position is decided by how many tabs have been passed: none means left,
    one means centre, two means right. That is the same rule the builder lays
    the paragraph out by.
    """
    items = {"left": None, "center": None, "right": None}
    sides = ["left", "center", "right"]
    index = 0
    in_field = False

    for child in paragraph._p:
        tag = child.tag

        if tag == qn("w:sdt"):
            binding = child.find(".//" + qn("w:dataBinding"))
            alias = child.find(".//" + qn("w:alias"))
            if binding is None or index >= len(sides):
                continue
            xpath = binding.get(qn("w:xpath")) or ""
            name = alias.get(qn("w:val")) if alias is not None else ""
            items[sides[index]] = ("bound", xpath, name)
            continue

        if tag != qn("w:r"):
            continue

        if child.find(qn("w:tab")) is not None:
            index += 1
            continue

        char = child.find(qn("w:fldChar"))
        if char is not None:
            kind = char.get(qn("w:fldCharType"))
            if kind == "begin":
                in_field = True
            elif kind == "end":
                in_field = False
            continue

        instruction = child.find(qn("w:instrText"))
        if instruction is not None:
            if (instruction.text or "").strip().upper().startswith("PAGE") \
                    and index < len(sides):
                items[sides[index]] = ("page", "", "")
            continue

        node = child.find(qn("w:t"))
        if node is not None and not in_field and index < len(sides):
            text = (node.text or "").strip()
            if text:
                items[sides[index]] = ("text", "", text)

    return items


def _read_header_footer(doc, header_footer, report):
    """Recover which field sits in each of the six positions."""
    if not doc.sections:
        report.defaulted.append("Header and footer")
        return
    section = doc.sections[0]

    header_footer["different_first_page"] = bool(
        section.different_first_page_header_footer
    )

    # Start from nothing placed, so a position the file leaves empty does not
    # keep whatever the previous config had in it.
    for key in PLACEABLE_KEYS:
        field = header_footer.setdefault(key, {})
        field["enabled"] = False
        field["slot"] = ""

    found, custom_found = False, False
    custom_slots = {}

    for area, container in (("header", section.header),
                            ("footer", section.footer)):
        if not container.paragraphs:
            continue
        for side, item in _slot_items(container.paragraphs[0]).items():
            if item is None:
                continue
            kind, xpath, name = item
            slot = "%s_%s" % (area, side)

            if kind == "page":
                key = "page_number"
            elif kind == "text":
                key = "static_text"
                header_footer[key]["text"] = name
            elif "title" in xpath:
                key = "doc_title"
            elif "creator" in xpath:
                key = "author"
            else:
                match = re.search(r"custom_(\d)", xpath)
                if not match:
                    continue
                key = "custom_%s" % match.group(1)
                if key not in CUSTOM_FIELD_KEYS:
                    continue
                header_footer[key]["name"] = name
                custom_slots[key] = slot
                custom_found = True

            header_footer[key]["enabled"] = True
            header_footer[key]["slot"] = slot
            found = True

    # Custom field starting values live in the custom XML part, not the header.
    for key, value in _custom_xml_values(doc).items():
        if key in CUSTOM_FIELD_KEYS and isinstance(
                header_footer.get(key), dict):
            header_footer[key]["value"] = value

    report.record("Header and footer positions", found)
    report.record("Custom fields", custom_found)


def _custom_xml_values(doc):
    """Read the starting values out of our own custom XML part, if present."""
    values = {}
    try:
        parts = doc.part.package.iter_parts()
    except AttributeError:
        return values
    for part in parts:
        if "customXml/item" not in str(part.partname):
            continue
        try:
            root = etree.fromstring(part.blob)
        except Exception:
            continue
        if not str(root.tag).startswith("{%s}" % STUDIO_NS):
            continue
        for element in root:
            key = etree.QName(element).localname
            if key in CUSTOM_FIELD_KEYS or key in STUDIO_META_KEYS:
                values[key] = element.text or ""
    return values


def _read_metadata(doc, template_cfg, document_cfg, header_footer, report):
    """
    Read the file properties into the template details.

    Mapping the document title or the author into a header changes what those
    two properties mean: they stop being file metadata and become the live
    value the user types on the cover. So when a field is mapped, its property
    is read into the cover placeholder and the template detail is left alone,
    or importing a file would overwrite the template name with a cover
    placeholder.
    """
    props = doc.core_properties
    found = False

    if props.comments:
        template_cfg["description"] = props.comments
        found = True

    # Written by this tool, and the only place the template's own name and
    # author survive once those core properties are bound to live fields.
    stored = _custom_xml_values(doc)
    for stored_key, detail_key in (
        ("template_name", "name"), ("template_author", "author"),
    ):
        if stored.get(stored_key):
            template_cfg[detail_key] = stored[stored_key]
            found = True

    for prop, mapped_key, detail_key, cover_key in (
        (props.title, "doc_title", "name", "title_placeholder"),
        (props.author, "author", "author", "author_placeholder"),
    ):
        if not prop:
            continue
        found = True
        if (header_footer.get(mapped_key) or {}).get("enabled"):
            document_cfg[cover_key] = prop
        elif not stored.get("template_name" if detail_key == "name"
                            else "template_author"):
            template_cfg[detail_key] = prop

    report.record("Template details", found)


def _read_toc_styles(doc, toc_cfg, report):
    heading = _style(doc, "TOC Heading")
    found = False
    if heading is not None:
        for key, value in (
            ("title_font", _font_of(heading)),
            ("title_fontSize", _size_of(heading)),
            ("title_bold", heading.font.bold),
            ("title_color", _color_of(heading)),
            ("spacing_after", _twips_of(heading.paragraph_format.space_after)),
        ):
            if value is not None:
                toc_cfg[key] = value
                found = True
    report.record("Contents heading style", found)

    entry = _style(doc, "TOC 1")
    found = False
    if entry is not None:
        for key, value in (
            ("entry_font", _font_of(entry)),
            ("entry_fontSize", _size_of(entry)),
            ("entry_color", _color_of(entry)),
            ("entry_bold_level1", entry.font.bold),
            ("entry_spacing_after",
             _twips_of(entry.paragraph_format.space_after)),
        ):
            if value is not None:
                toc_cfg[key] = value
                found = True
    report.record("Contents entry style", found)

    # The contents heading text is the paragraph that carries TOC Heading, not
    # anything stored in the style itself.
    for paragraph in doc.paragraphs:
        if (paragraph.style is not None
                and paragraph.style.name == "TOC Heading"
                and paragraph.text.strip()):
            toc_cfg["title"] = paragraph.text.strip()
            break


# ----------------------------------------------------------------------------
# Public entry point
# ----------------------------------------------------------------------------

def read_template(source, defaults):
    """
    Read a .dotx or .docx into a styling config.

    `source` is a path or the file bytes. `defaults` is the config to fall back
    on field by field, normally the one currently loaded in the GUI, so a style
    the file does not define keeps the value the user already had.

    Returns (config, report). The report lists which groups of settings came
    out of the file and which were left at their existing values.
    """
    config = copy.deepcopy(defaults)
    report = Report()

    temp_path = _as_readable_docx(source)
    try:
        try:
            doc = Document(temp_path)
        except Exception as error:
            raise TemplateReadError(
                f"Word could not be read from that file: {error}"
            )

        styles = config["styles"]

        for index, key in enumerate(HEADING_KEYS, start=1):
            _read_text_style(
                doc, f"Heading {index}", styles["headings"][key],
                f"Heading {index}", report,
                line_spacing=False, caps=True,
            )

        _read_text_style(
            doc, "Body Text", styles["paragraph"]["body"], "Body text", report,
        )
        _read_text_style(
            doc, "Caption", styles["paragraph"]["caption"], "Caption", report,
            line_spacing=False,
        )
        if styles["paragraph"].get("quote") is not None:
            _read_text_style(
                doc, "Quote", styles["paragraph"]["quote"], "Quote block",
                report, indent=True,
            )

        for level_key, style_name in (
            ("level_1", "List Bullet"), ("level_2", "List Bullet 2"),
        ):
            if styles["bullets"].get(level_key) is not None:
                _read_bullet(
                    doc, style_name, styles["bullets"][level_key],
                    f"Bullets {level_key.replace('_', ' ')}", report,
                )

        _read_toc_styles(doc, styles["toc"], report)
        _read_page_setup(doc, config["page_setup"], report)
        _read_heading_numbering(doc, config["document"], report)
        _read_document_structure(
            doc, config["document"], styles["toc"], report,
        )
        _read_header_footer(
            doc, config.setdefault("header_footer", {}), report,
        )
        _read_metadata(
            doc, config["template"], config["document"],
            config.get("header_footer", {}), report,
        )

        # Inline run styles are applied as direct formatting rather than as
        # character styles, so there is nothing in the file to read them from.
        report.defaulted.append("Inline text styles")
    finally:
        try:
            os.unlink(temp_path)
        except OSError:
            pass

    return config, report.as_dict()
