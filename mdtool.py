#!/usr/bin/env python3
"""
mdtool - manual CLI utility for working with markdown stored in a local
GitHub repository using Microsoft Word on Windows.

Version 1 supports:
    - configuration mode (--config, or auto-triggered when config.ini is absent)
    - --fromgit: bootstrap .docx working copies from .md files

Out of scope for v1: --togit, git/GitHub integration, recursive traversal,
filesystem monitoring, image import.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import re
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

try:
    from markdown_it import MarkdownIt
    from markdown_it.token import Token
except ImportError:
    sys.stderr.write(
        "ERROR: markdown-it-py is required.\n"
        "Install with: pip install markdown-it-py\n"
    )
    sys.exit(2)

try:
    from docx import Document
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt
    from docx.table import Table as DocxTable
    from docx.text.paragraph import Paragraph as DocxParagraph
except ImportError:
    sys.stderr.write(
        "ERROR: python-docx is required.\n"
        "Install with: pip install python-docx\n"
    )
    sys.exit(2)

try:
    from PIL import Image
except ImportError:
    sys.stderr.write(
        "ERROR: Pillow is required (for --togit image rendering).\n"
        "Install with: pip install Pillow\n"
    )
    sys.exit(2)


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

CONFIG_FILENAME = "config.ini"
CONFIG_KEYS: Tuple[str, ...] = ("GitDir", "WorkDir")

CODE_FONT = "Consolas"
CODE_BLOCK_STYLE_NAME = "CodeBlock"
INLINE_CODE_STYLE_NAME = "InlineCode"
TABLE_STYLE_NAME = "Table Grid"
QUOTE_STYLE_NAME = "Quote"

CODE_SHADING_FILL = "F2F2F2"
HR_BORDER_COLOR = "808080"

# --togit
ASSETS_DIR_NAME = "assets"
MANIFEST_FILENAME = ".mdwordtool-manifest.json"
MANIFEST_VERSION = 1
STATE_FILENAME = ".mdwordtool-state.json"
STATE_VERSION = 1
IMAGE_DPI = 150  # Render images at this DPI of Word display size.
SUPPORTED_IMAGE_CONTENT_TYPES = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/jpg": "jpg",
}
LIST_INDENT_STEP_CM = 0.75  # Matches --fromgit indent step.


# ---------------------------------------------------------------------------
# Error handling
# ---------------------------------------------------------------------------

def fail(message: str) -> None:
    """Print error to stderr and exit non-zero. Fail-fast per spec."""
    sys.stderr.write(f"ERROR: {message}\n")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Config handling
# ---------------------------------------------------------------------------

def config_path() -> Path:
    return Path.cwd() / CONFIG_FILENAME


def load_config(path: Path) -> Optional[Dict[str, str]]:
    """Load config from key=value INI-style file. Returns None if file absent."""
    if not path.exists():
        return None
    cfg: Dict[str, str] = {}
    try:
        with open(path, "r", encoding="utf-8") as fp:
            for raw in fp:
                line = raw.strip()
                if not line or line.startswith("#") or line.startswith(";"):
                    continue
                if "=" not in line:
                    continue
                k, _, v = line.partition("=")
                cfg[k.strip()] = v.strip()
    except OSError as exc:
        fail(f"Failed to read {path}: {exc}")
    return cfg


def write_config(path: Path, cfg: Dict[str, str]) -> None:
    try:
        with open(path, "w", encoding="utf-8") as fp:
            for key in CONFIG_KEYS:
                fp.write(f"{key} = {cfg[key]}\n")
    except OSError as exc:
        fail(f"Failed to write {path}: {exc}")


def run_config_mode(existing: Optional[Dict[str, str]]) -> None:
    print("Entering configuration mode.")
    new_cfg: Dict[str, str] = {}

    for key in CONFIG_KEYS:
        default = (existing or {}).get(key, "")
        prompt = key
        if default:
            prompt += f" [{default}]"
        prompt += ": "
        try:
            value = input(prompt).strip()
        except EOFError:
            fail("Input ended unexpectedly during configuration.")
            return  # unreachable
        if not value:
            value = default
        if not value:
            fail(f"{key} must not be empty.")
        new_cfg[key] = value

    print("\nProposed configuration:")
    for key in CONFIG_KEYS:
        print(f"  {key} = {new_cfg[key]}")

    try:
        confirm = input("Confirm and write config? [y/n]: ").strip().lower()
    except EOFError:
        fail("Input ended unexpectedly during confirmation.")
        return  # unreachable

    if confirm != "y":
        print("Configuration cancelled.")
        sys.exit(0)

    # Validate that all configured paths exist as directories.
    for key in CONFIG_KEYS:
        p = Path(new_cfg[key])
        if not p.is_dir():
            fail(
                f"Path for {key} does not exist or is not a directory: "
                f"{new_cfg[key]}"
            )

    write_config(config_path(), new_cfg)
    print(f"Configuration written to {config_path()}.")
    sys.exit(0)


# ---------------------------------------------------------------------------
# OOXML helpers
# ---------------------------------------------------------------------------

def _shading(fill_hex: str) -> "OxmlElement":
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    return shd


def _add_paragraph_bottom_border(paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    pbdr = p_pr.find(qn("w:pBdr"))
    if pbdr is None:
        pbdr = OxmlElement("w:pBdr")
        p_pr.append(pbdr)
    border = OxmlElement("w:bottom")
    border.set(qn("w:val"), "single")
    border.set(qn("w:sz"), "6")
    border.set(qn("w:space"), "1")
    border.set(qn("w:color"), HR_BORDER_COLOR)
    pbdr.append(border)


def _mark_quick_format(style) -> None:
    """Add <w:qFormat/> so the style appears in Word's Styles gallery."""
    el = style.element
    # qFormat lives after w:name/w:basedOn/w:next/w:link/w:uiPriority and
    # before run/paragraph property blocks. Appending works in practice
    # because Word tolerates trailing flag elements before rPr/pPr; to be
    # safe, insert immediately after w:name if present.
    qfmt = OxmlElement("w:qFormat")
    name = el.find(qn("w:name"))
    if name is not None:
        name.addnext(qfmt)
    else:
        el.append(qfmt)


def _ensure_styles(doc) -> None:
    """Create the minimal custom style set used for round-trippable output."""
    styles = doc.styles
    existing = {s.name for s in styles}

    # Inline code character style.
    if INLINE_CODE_STYLE_NAME not in existing:
        s = styles.add_style(INLINE_CODE_STYLE_NAME, WD_STYLE_TYPE.CHARACTER)
        s.font.name = CODE_FONT
        s.font.size = Pt(10)
        r_pr = s.element.get_or_add_rPr()
        r_pr.append(_shading(CODE_SHADING_FILL))
        # Force fixed-width font for all script ranges.
        rfonts = OxmlElement("w:rFonts")
        rfonts.set(qn("w:ascii"), CODE_FONT)
        rfonts.set(qn("w:hAnsi"), CODE_FONT)
        rfonts.set(qn("w:cs"), CODE_FONT)
        r_pr.append(rfonts)
        _mark_quick_format(s)

    # Code block paragraph style.
    if CODE_BLOCK_STYLE_NAME not in existing:
        s = styles.add_style(CODE_BLOCK_STYLE_NAME, WD_STYLE_TYPE.PARAGRAPH)
        s.font.name = CODE_FONT
        s.font.size = Pt(10)
        s.paragraph_format.space_before = Pt(0)
        s.paragraph_format.space_after = Pt(0)
        p_pr = s.element.get_or_add_pPr()
        p_pr.append(_shading(CODE_SHADING_FILL))
        _mark_quick_format(s)


# ---------------------------------------------------------------------------
# Markdown -> DOCX conversion
# ---------------------------------------------------------------------------

class InlineFormat:
    """Mutable bag of inline formatting flags as we walk inline tokens."""

    __slots__ = ("bold", "italic", "strike", "code")

    def __init__(self) -> None:
        self.bold = False
        self.italic = False
        self.strike = False
        self.code = False

    def copy(self) -> "InlineFormat":
        c = InlineFormat()
        c.bold = self.bold
        c.italic = self.italic
        c.strike = self.strike
        c.code = self.code
        return c


def _apply_run_formatting(run, fmt: InlineFormat,
                          inline_code_style) -> None:
    if fmt.code:
        run.style = inline_code_style
    if fmt.bold:
        run.bold = True
    if fmt.italic:
        run.italic = True
    if fmt.strike:
        run.font.strike = True


def _wrap_runs_in_hyperlink(paragraph, runs, url: str) -> None:
    """Move trailing w:r elements into a new w:hyperlink with relationship."""
    if not runs:
        return
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    first_r = runs[0]._element
    first_r.addprevious(hyperlink)
    for r in runs:
        paragraph._p.remove(r._element)
        hyperlink.append(r._element)

    # Apply Hyperlink character style to each run.
    for r in runs:
        r_pr = r._element.get_or_add_rPr()
        rstyle = OxmlElement("w:rStyle")
        rstyle.set(qn("w:val"), "Hyperlink")
        r_pr.insert(0, rstyle)


def _render_inline(doc, paragraph, tokens: List[Token],
                   fmt: InlineFormat, inline_code_style) -> None:
    """Render an inline token stream into the given paragraph."""
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        t = tok.type

        if t == "text":
            run = paragraph.add_run(tok.content)
            _apply_run_formatting(run, fmt, inline_code_style)

        elif t == "code_inline":
            sub = fmt.copy()
            sub.code = True
            run = paragraph.add_run(tok.content)
            _apply_run_formatting(run, sub, inline_code_style)

        elif t == "softbreak":
            run = paragraph.add_run(" ")
            _apply_run_formatting(run, fmt, inline_code_style)

        elif t == "hardbreak":
            run = paragraph.add_run()
            run.add_break()

        elif t == "strong_open":
            fmt.bold = True
        elif t == "strong_close":
            fmt.bold = False

        elif t == "em_open":
            fmt.italic = True
        elif t == "em_close":
            fmt.italic = False

        elif t == "s_open":
            fmt.strike = True
        elif t == "s_close":
            fmt.strike = False

        elif t == "link_open":
            href = tok.attrGet("href") or ""
            # Find matching link_close, recurse on inner tokens, then wrap.
            depth = 1
            j = i + 1
            while j < len(tokens) and depth > 0:
                if tokens[j].type == "link_open":
                    depth += 1
                elif tokens[j].type == "link_close":
                    depth -= 1
                if depth == 0:
                    break
                j += 1
            inner_tokens = tokens[i + 1:j]

            # Snapshot existing runs; render inner; identify new runs; wrap.
            existing_run_elements = set(id(r._element) for r in paragraph.runs)
            _render_inline(doc, paragraph, inner_tokens, fmt, inline_code_style)
            new_runs = [r for r in paragraph.runs
                        if id(r._element) not in existing_run_elements]
            if href:
                _wrap_runs_in_hyperlink(paragraph, new_runs, href)
            i = j  # skip past link_close

        elif t == "link_close":
            # Already consumed by link_open branch.
            pass

        elif t == "image":
            # v1: images are out of scope; emit alt text as plain run.
            alt = tok.content or tok.attrGet("alt") or ""
            if alt:
                run = paragraph.add_run(alt)
                _apply_run_formatting(run, fmt, inline_code_style)

        elif t == "html_inline":
            # v1: raw HTML is out of scope; pass through as literal text so
            # nothing is silently dropped.
            run = paragraph.add_run(tok.content)
            _apply_run_formatting(run, fmt, inline_code_style)

        else:
            # Unknown inline token. Fail fast per spec.
            fail(f"Unsupported inline markdown token: {t}")

        i += 1


def _cell_alignment(style_attr: str) -> Optional[int]:
    if not style_attr:
        return None
    if "text-align:right" in style_attr:
        return WD_PARAGRAPH_ALIGNMENT.RIGHT
    if "text-align:center" in style_attr:
        return WD_PARAGRAPH_ALIGNMENT.CENTER
    if "text-align:left" in style_attr:
        return WD_PARAGRAPH_ALIGNMENT.LEFT
    return None


class BlockRenderer:
    """Walks a markdown-it block token stream and emits DOCX content."""

    def __init__(self, doc) -> None:
        self.doc = doc
        self.inline_code_style = doc.styles[INLINE_CODE_STYLE_NAME]
        # Each entry: {"ordered": bool, "level": int}
        self.list_stack: List[Dict[str, object]] = []
        # Set when a list_item_open is processed; consumed by the next
        # paragraph created. Cleared after first use so subsequent paragraphs
        # inside the same item become continuation paragraphs.
        self._pending_list_item = False
        # Depth of enclosing blockquotes; non-zero means paragraphs not
        # otherwise styled (e.g. by list state) get the Quote style.
        self._blockquote_depth = 0

    # -- entry point --

    def render(self, tokens: List[Token]) -> None:
        i = 0
        while i < len(tokens):
            tok = tokens[i]
            i = self._render_block(tokens, i, tok)

    # -- helpers --

    def _find_matching_close(self, tokens: List[Token], start: int,
                             open_type: str, close_type: str) -> int:
        depth = 1
        j = start + 1
        while j < len(tokens):
            if tokens[j].type == open_type:
                depth += 1
            elif tokens[j].type == close_type:
                depth -= 1
                if depth == 0:
                    return j
            j += 1
        fail(f"Malformed markdown: missing {close_type}")
        return -1  # unreachable

    def _inline_tokens_of(self, tokens: List[Token], idx: int) -> List[Token]:
        """If tokens[idx] is an inline token, return its children list."""
        if idx < len(tokens) and tokens[idx].type == "inline":
            return tokens[idx].children or []
        return []

    # -- block dispatch --

    def _render_block(self, tokens: List[Token], i: int, tok: Token) -> int:
        t = tok.type

        if t == "heading_open":
            level = int(tok.tag[1])  # 'h1' -> 1
            close_idx = self._find_matching_close(
                tokens, i, "heading_open", "heading_close"
            )
            inline_idx = i + 1
            p = self.doc.add_paragraph(style=f"Heading {level}")
            fmt = InlineFormat()
            _render_inline(
                self.doc, p, self._inline_tokens_of(tokens, inline_idx),
                fmt, self.inline_code_style,
            )
            return close_idx + 1

        if t == "paragraph_open":
            close_idx = self._find_matching_close(
                tokens, i, "paragraph_open", "paragraph_close"
            )
            inline_idx = i + 1
            p = self._new_paragraph()
            fmt = InlineFormat()
            _render_inline(
                self.doc, p, self._inline_tokens_of(tokens, inline_idx),
                fmt, self.inline_code_style,
            )
            return close_idx + 1

        if t == "bullet_list_open":
            close_idx = self._find_matching_close(
                tokens, i, "bullet_list_open", "bullet_list_close"
            )
            self.list_stack.append({
                "ordered": False,
                "level": len(self.list_stack),
            })
            self._render_list_body(tokens, i + 1, close_idx)
            self.list_stack.pop()
            return close_idx + 1

        if t == "ordered_list_open":
            close_idx = self._find_matching_close(
                tokens, i, "ordered_list_open", "ordered_list_close"
            )
            self.list_stack.append({
                "ordered": True,
                "level": len(self.list_stack),
            })
            self._render_list_body(tokens, i + 1, close_idx)
            self.list_stack.pop()
            return close_idx + 1

        if t == "blockquote_open":
            close_idx = self._find_matching_close(
                tokens, i, "blockquote_open", "blockquote_close"
            )
            self._render_blockquote(tokens, i + 1, close_idx)
            return close_idx + 1

        if t == "fence" or t == "code_block":
            self._render_code_block(tok.content)
            return i + 1

        if t == "hr":
            self._render_hr()
            return i + 1

        if t == "table_open":
            close_idx = self._find_matching_close(
                tokens, i, "table_open", "table_close"
            )
            self._render_table(tokens, i + 1, close_idx)
            return close_idx + 1

        if t == "html_block":
            # v1: raw HTML blocks are unsupported. Spec marks them as such.
            # Emit content as a plain paragraph so nothing is silently lost,
            # but flag this loudly.
            sys.stderr.write(
                "WARNING: raw HTML block encountered; emitting as plain text.\n"
            )
            p = self.doc.add_paragraph(tok.content.rstrip("\n"))
            return i + 1

        if t in ("paragraph_close", "heading_close", "bullet_list_close",
                 "ordered_list_close", "list_item_close", "blockquote_close",
                 "table_close", "thead_close", "tbody_close", "tr_close",
                 "th_close", "td_close"):
            return i + 1

        # Unknown block token. Fail fast.
        fail(f"Unsupported block markdown token: {t}")
        return i + 1  # unreachable

    # -- block helpers --

    def _new_paragraph(self):
        """Create a paragraph, applying the correct style for the current
        block context (list item, blockquote, or plain)."""
        # Highest priority: a list item just opened; consume the slot.
        if self._pending_list_item and self.list_stack:
            self._pending_list_item = False
            top = self.list_stack[-1]
            style_name = "List Number" if top["ordered"] else "List Bullet"
            p = self.doc.add_paragraph(style=style_name)
            level = int(top["level"])
            if level > 0:
                # Built-in template does not guarantee "List Bullet 2" etc.,
                # so we keep the base style and apply manual indentation.
                p.paragraph_format.left_indent = Cm(0.75 + 0.75 * level)
            return p

        # Continuation paragraph inside an already-bulleted list item.
        if self.list_stack:
            p = self.doc.add_paragraph()
            level = len(self.list_stack)
            p.paragraph_format.left_indent = Cm(0.75 + 0.75 * max(level - 1, 0) + 0.6)
            return p

        # Blockquote with no enclosing list.
        if self._blockquote_depth > 0:
            try:
                return self.doc.add_paragraph(style=QUOTE_STYLE_NAME)
            except KeyError:
                return self.doc.add_paragraph()

        return self.doc.add_paragraph()

    def _render_list_body(self, tokens: List[Token],
                          start: int, end_exclusive: int) -> None:
        i = start
        while i < end_exclusive:
            tok = tokens[i]
            if tok.type == "list_item_open":
                close_idx = self._find_matching_close(
                    tokens, i, "list_item_open", "list_item_close"
                )
                self._render_list_item(tokens, i + 1, close_idx)
                i = close_idx + 1
            else:
                i += 1

    def _render_list_item(self, tokens: List[Token],
                          start: int, end_exclusive: int) -> None:
        """Render one list item. The first block-level paragraph emitted
        consumes the list bullet; subsequent paragraphs in the same item
        become continuation paragraphs (indented, no bullet)."""
        self._pending_list_item = True
        try:
            i = start
            while i < end_exclusive:
                i = self._render_block(tokens, i, tokens[i])
        finally:
            # If the item had no paragraph content at all (rare), clear flag
            # so it does not leak into a sibling.
            self._pending_list_item = False

    def _render_blockquote(self, tokens: List[Token],
                           start: int, end_exclusive: int) -> None:
        self._blockquote_depth += 1
        try:
            i = start
            while i < end_exclusive:
                i = self._render_block(tokens, i, tokens[i])
        finally:
            self._blockquote_depth -= 1

    def _render_code_block(self, content: str) -> None:
        # Strip a single trailing newline; preserve internal newlines as
        # separate paragraphs all sharing the code block style.
        text = content.rstrip("\n")
        if text == "":
            self.doc.add_paragraph(style=CODE_BLOCK_STYLE_NAME)
            return
        for line in text.split("\n"):
            p = self.doc.add_paragraph(style=CODE_BLOCK_STYLE_NAME)
            if line:
                p.add_run(line)

    def _render_hr(self) -> None:
        p = self.doc.add_paragraph()
        _add_paragraph_bottom_border(p)

    def _render_table(self, tokens: List[Token],
                      start: int, end_exclusive: int) -> None:
        # First pass: collect rows of cells with alignment metadata.
        rows: List[List[Tuple[List[Token], Optional[int], bool]]] = []
        # cell = (inline_tokens, alignment, is_header)

        i = start
        while i < end_exclusive:
            tok = tokens[i]
            if tok.type == "tr_open":
                tr_close = self._find_matching_close(
                    tokens, i, "tr_open", "tr_close"
                )
                row: List[Tuple[List[Token], Optional[int], bool]] = []
                j = i + 1
                while j < tr_close:
                    ctok = tokens[j]
                    if ctok.type in ("th_open", "td_open"):
                        is_header = (ctok.type == "th_open")
                        cell_close = self._find_matching_close(
                            tokens, j,
                            "th_open" if is_header else "td_open",
                            "th_close" if is_header else "td_close",
                        )
                        style_attr = ctok.attrGet("style") or ""
                        alignment = _cell_alignment(style_attr)
                        inline_tokens: List[Token] = []
                        k = j + 1
                        while k < cell_close:
                            if tokens[k].type == "inline":
                                inline_tokens = tokens[k].children or []
                                break
                            k += 1
                        row.append((inline_tokens, alignment, is_header))
                        j = cell_close + 1
                    else:
                        j += 1
                rows.append(row)
                i = tr_close + 1
            else:
                i += 1

        if not rows:
            return

        ncols = max(len(r) for r in rows)
        table = self.doc.add_table(rows=len(rows), cols=ncols)
        try:
            table.style = TABLE_STYLE_NAME
        except KeyError:
            pass

        for ri, row in enumerate(rows):
            for ci in range(ncols):
                cell = table.rows[ri].cells[ci]
                # Clear the default empty paragraph python-docx inserts.
                # We will write our content into the first paragraph directly.
                target_p = cell.paragraphs[0]
                target_p.text = ""
                if ci < len(row):
                    inline_tokens, alignment, is_header = row[ci]
                    fmt = InlineFormat()
                    if is_header:
                        fmt.bold = True
                    _render_inline(
                        self.doc, target_p, inline_tokens,
                        fmt, self.inline_code_style,
                    )
                    if alignment is not None:
                        target_p.alignment = alignment


def convert_markdown_to_docx(md_path: Path, docx_path: Path) -> None:
    """Convert one markdown file to a docx working copy. Fail fast on error."""
    try:
        with open(md_path, "r", encoding="utf-8") as fp:
            md_text = fp.read()
    except OSError as exc:
        fail(f"Failed to read {md_path}: {exc}")
        return  # unreachable
    except UnicodeDecodeError as exc:
        fail(f"{md_path} is not valid UTF-8: {exc}")
        return  # unreachable

    md = MarkdownIt("commonmark", {"html": False, "linkify": False, "breaks": False})
    md.enable(["table", "strikethrough"])

    try:
        tokens = md.parse(md_text)
    except Exception as exc:  # markdown-it rarely raises, but be defensive.
        fail(f"Failed to parse {md_path}: {exc}")
        return  # unreachable

    doc = Document()
    _ensure_styles(doc)

    renderer = BlockRenderer(doc)
    renderer.render(tokens)

    try:
        doc.save(str(docx_path))
    except OSError as exc:
        fail(f"Failed to write {docx_path}: {exc}")


# ---------------------------------------------------------------------------
# --togit: DOCX -> Markdown
# ---------------------------------------------------------------------------
#
# Design notes:
#
# * Style detection is anchored to the styles --fromgit produces. The user is
#   expected to edit content within that style vocabulary; styles outside the
#   supported set are mapped to plain paragraphs (per spec: "do not preserve
#   arbitrary Word formatting outside the supported markdown-compatible style
#   set").
# * Block order: walk document.element.body children directly so paragraphs
#   and tables appear in their actual document order (python-docx exposes
#   them as separate lists which loses interleaving).
# * Tracked changes: detected via <w:ins>/<w:del> presence in the document
#   XML. python-docx does not cleanly expose an accepted-view, so per spec
#   ("if tracked changes cannot be resolved safely and deterministically,
#   fail fast"), refuse to convert until the user accepts/rejects them.
# * Images: rendered at IMAGE_DPI of their Word display size, never upscaled.
#   Hashed by their rendered bytes. Filenames are stable across runs when
#   content is unchanged; changed content gets a new filename.

# --- OOXML namespace shortcuts used in this section ---
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}


# ---------------------------------------------------------------------------
# Manifest
# ---------------------------------------------------------------------------

def _manifest_path(git_dir: Path) -> Path:
    return git_dir / ASSETS_DIR_NAME / MANIFEST_FILENAME


def load_manifest(git_dir: Path) -> Dict:
    path = _manifest_path(git_dir)
    if not path.exists():
        return {"version": MANIFEST_VERSION, "documents": {}}
    try:
        with open(path, "r", encoding="utf-8") as fp:
            data = json.load(fp)
    except (OSError, json.JSONDecodeError) as exc:
        fail(f"Manifest is corrupt or unreadable ({path}): {exc}")
        return {}  # unreachable
    if not isinstance(data, dict) or "documents" not in data:
        fail(f"Manifest has unexpected shape: {path}")
    data.setdefault("version", MANIFEST_VERSION)
    data.setdefault("documents", {})
    return data


def save_manifest(git_dir: Path, manifest: Dict) -> None:
    path = _manifest_path(git_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    try:
        with open(tmp, "w", encoding="utf-8") as fp:
            json.dump(manifest, fp, indent=2, sort_keys=True)
            fp.write("\n")
        os.replace(tmp, path)
    except OSError as exc:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass
        fail(f"Failed to update manifest {path}: {exc}")


# ---------------------------------------------------------------------------
# Sync state (per-WorkDir change detection)
# ---------------------------------------------------------------------------
#
# Records mtime + size of each .docx at the time of its last successful
# conversion. A subsequent --togit run skips files whose mtime AND size
# match the recorded values. mtime+size combined is cheap, deterministic,
# and good enough in practice; a content hash would catch the (rare) case
# of an edit that preserves both, at the cost of hashing every file each
# run.
#
# Deleting this file forces a full resync on the next run.

def _state_path(work_dir: Path) -> Path:
    return work_dir / STATE_FILENAME


def load_state(work_dir: Path) -> Dict:
    path = _state_path(work_dir)
    if not path.exists():
        return {"version": STATE_VERSION, "files": {}}
    try:
        with open(path, "r", encoding="utf-8") as fp:
            data = json.load(fp)
    except (OSError, json.JSONDecodeError) as exc:
        fail(f"Sync state file is corrupt or unreadable ({path}): {exc}")
        return {}  # unreachable
    if not isinstance(data, dict) or "files" not in data:
        fail(f"Sync state file has unexpected shape: {path}")
    data.setdefault("version", STATE_VERSION)
    data.setdefault("files", {})
    return data


def save_state(work_dir: Path, state: Dict) -> None:
    path = _state_path(work_dir)
    tmp = path.with_suffix(path.suffix + ".tmp")
    try:
        with open(tmp, "w", encoding="utf-8") as fp:
            json.dump(state, fp, indent=2, sort_keys=True)
            fp.write("\n")
        os.replace(tmp, path)
    except OSError as exc:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass
        fail(f"Failed to update sync state {path}: {exc}")


def _docx_unchanged_since_last_sync(docx_path: Path, state: Dict) -> bool:
    entry = state["files"].get(docx_path.name)
    if not entry:
        return False
    try:
        stat = docx_path.stat()
    except OSError:
        return False
    return (
        stat.st_mtime == entry.get("mtime")
        and stat.st_size == entry.get("size")
    )


def _record_sync(docx_path: Path, state: Dict) -> None:
    stat = docx_path.stat()
    state["files"][docx_path.name] = {
        "mtime": stat.st_mtime,
        "size": stat.st_size,
        "last_synced": datetime.now(timezone.utc).strftime(
            "%Y-%m-%dT%H:%M:%SZ"
        ),
    }


# ---------------------------------------------------------------------------
# Tracked changes
# ---------------------------------------------------------------------------

def _check_tracked_changes(doc, docx_path: Path) -> None:
    body_xml = doc.element.body
    # findall returns lists; use simple xpath for any descendant.
    ins = body_xml.find(f".//{{{NS['w']}}}ins")
    dele = body_xml.find(f".//{{{NS['w']}}}del")
    if ins is not None or dele is not None:
        fail(
            f"{docx_path.name} contains unresolved tracked changes. "
            f"Accept or reject all changes in Word before running --togit."
        )


# ---------------------------------------------------------------------------
# Image extraction
# ---------------------------------------------------------------------------

class ExtractedImage:
    __slots__ = ("rendered_bytes", "ext", "hash_hex", "alt_text",
                 "width_px", "height_px")

    def __init__(self, rendered_bytes: bytes, ext: str, alt_text: str,
                 width_px: int, height_px: int) -> None:
        self.rendered_bytes = rendered_bytes
        self.ext = ext
        self.alt_text = alt_text
        self.width_px = width_px
        self.height_px = height_px
        self.hash_hex = hashlib.sha256(rendered_bytes).hexdigest()


def _emu_to_px(emu: int, dpi: int = IMAGE_DPI) -> int:
    # 914400 EMU per inch.
    return max(1, int(round(emu * dpi / 914400)))


def _extract_inline_image(drawing_el, doc, docx_path: Path) -> ExtractedImage:
    """Render a single inline image found inside a <w:drawing> element.

    Anchored/floating images are rejected per spec.
    """
    # Reject anchored/floating images.
    if drawing_el.find(f"{{{NS['wp']}}}anchor") is not None:
        fail(
            f"{docx_path.name} contains a floating or anchored image. "
            f"Only inline images are supported."
        )

    inline = drawing_el.find(f"{{{NS['wp']}}}inline")
    if inline is None:
        fail(
            f"{docx_path.name} contains a drawing element with no inline "
            f"layout; only inline images are supported."
        )

    # Extent (display size in EMU).
    extent = inline.find(f"{{{NS['wp']}}}extent")
    if extent is None:
        fail(f"{docx_path.name} has an inline image with no display extent.")
    try:
        cx = int(extent.get("cx", "0"))
        cy = int(extent.get("cy", "0"))
    except ValueError:
        fail(f"{docx_path.name} has an inline image with non-numeric extent.")
        cx = cy = 0  # unreachable

    # Alt text from docPr/@descr (preferred) or @title.
    doc_pr = inline.find(f"{{{NS['wp']}}}docPr")
    alt_text = ""
    if doc_pr is not None:
        alt_text = (doc_pr.get("descr") or doc_pr.get("title") or "").strip()

    # Find the blip (image reference) and optional srcRect (cropping).
    blip = inline.find(
        f".//{{{NS['a']}}}blip"
    )
    if blip is None:
        fail(f"{docx_path.name} has an inline image with no blip reference.")
    r_embed = blip.get(f"{{{NS['r']}}}embed")
    if not r_embed:
        fail(
            f"{docx_path.name} has an inline image with no r:embed "
            f"relationship."
        )

    src_rect = inline.find(
        f".//{{{NS['a']}}}srcRect"
    )

    # Resolve relationship to image part.
    try:
        image_part = doc.part.related_parts[r_embed]
    except KeyError:
        fail(
            f"{docx_path.name} references missing image relationship "
            f"{r_embed!r}."
        )
        return None  # unreachable

    content_type = image_part.content_type
    if content_type not in SUPPORTED_IMAGE_CONTENT_TYPES:
        fail(
            f"{docx_path.name} contains an unsupported image type "
            f"{content_type!r}. Only PNG and JPEG are supported."
        )
    out_ext = SUPPORTED_IMAGE_CONTENT_TYPES[content_type]

    # Render via Pillow.
    try:
        src_img = Image.open(io.BytesIO(image_part.blob))
        src_img.load()
    except Exception as exc:
        fail(f"Failed to decode an image in {docx_path.name}: {exc}")
        return None  # unreachable

    # Apply cropping if specified. srcRect values are in 1/100000 (i.e.
    # thousandths of a percent). Positive values trim from that edge.
    if src_rect is not None:
        try:
            l = int(src_rect.get("l", "0")) / 100000.0
            t = int(src_rect.get("t", "0")) / 100000.0
            r = int(src_rect.get("r", "0")) / 100000.0
            b = int(src_rect.get("b", "0")) / 100000.0
        except ValueError:
            fail(f"{docx_path.name} has non-numeric image crop metadata.")
            return None  # unreachable
        w, h = src_img.size
        left = max(0, int(round(w * l)))
        top = max(0, int(round(h * t)))
        right = min(w, int(round(w * (1.0 - r))))
        bottom = min(h, int(round(h * (1.0 - b))))
        if right <= left or bottom <= top:
            fail(
                f"{docx_path.name} has an image crop that collapses to "
                f"zero area; cannot render reliably."
            )
        src_img = src_img.crop((left, top, right, bottom))

    # Resize to Word's display size (downscale only, never upscale).
    target_w = _emu_to_px(cx)
    target_h = _emu_to_px(cy)
    if target_w < src_img.width or target_h < src_img.height:
        src_img.thumbnail((target_w, target_h), Image.LANCZOS)

    # Encode bytes deterministically. PNG uses optimize; JPEG uses quality=90.
    buf = io.BytesIO()
    try:
        if out_ext == "png":
            # Normalise to RGBA->RGB if palette/transparency would otherwise
            # produce inconsistent output across PIL versions.
            save_img = src_img
            if save_img.mode == "P":
                save_img = save_img.convert("RGBA")
            save_img.save(buf, format="PNG", optimize=True)
        else:
            save_img = src_img.convert("RGB") if src_img.mode != "RGB" else src_img
            save_img.save(buf, format="JPEG", quality=90, optimize=True)
    except Exception as exc:
        fail(f"Failed to render an image in {docx_path.name}: {exc}")
        return None  # unreachable

    return ExtractedImage(
        rendered_bytes=buf.getvalue(),
        ext=out_ext,
        alt_text=alt_text,
        width_px=src_img.width,
        height_px=src_img.height,
    )


# ---------------------------------------------------------------------------
# Markdown emission helpers
# ---------------------------------------------------------------------------

def _escape_md_text(text: str, *, in_table: bool = False) -> str:
    """Escape markdown-significant characters in literal text.

    Conservative but not paranoid; deliberately does not escape every
    character, only those likely to be misinterpreted by a GFM parser.
    """
    if not text:
        return text
    # Escape backslash first to avoid double-escaping.
    out = text.replace("\\", "\\\\")
    # Escape backticks, brackets, asterisks, underscores, pipes (in tables).
    for ch in "`*_[]<>":
        out = out.replace(ch, "\\" + ch)
    if in_table:
        out = out.replace("|", "\\|")
        # Newlines in a cell would break the row; collapse to spaces.
        out = out.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    return out


# ---------------------------------------------------------------------------
# DOCX -> Markdown converter
# ---------------------------------------------------------------------------

class _RunPiece:
    """One contiguous text piece with formatting flags and optional link."""

    __slots__ = ("text", "bold", "italic", "code", "strike", "link",
                 "is_image", "image_index")

    def __init__(self, text: str = "", *, bold: bool = False,
                 italic: bool = False, code: bool = False,
                 strike: bool = False, link: Optional[str] = None,
                 is_image: bool = False,
                 image_index: Optional[int] = None) -> None:
        self.text = text
        self.bold = bold
        self.italic = italic
        self.code = code
        self.strike = strike
        self.link = link
        self.is_image = is_image
        self.image_index = image_index

    def fmt_key(self) -> Tuple:
        return (self.bold, self.italic, self.code, self.strike, self.link)


class DocxToMarkdown:
    def __init__(self, doc, docx_path: Path,
                 image_filenames: List[str]) -> None:
        self.doc = doc
        self.docx_path = docx_path
        # image_filenames[i] is the final filename (no path) for the i-th
        # inline image encountered in document order.
        self.image_filenames = image_filenames
        self._image_counter = 0  # next image index to consume

    # --- public ---

    def render(self) -> str:
        blocks: List[str] = []
        body = self.doc.element.body
        items = list(body.iterchildren())

        i = 0
        while i < len(items):
            el = items[i]
            tag = el.tag

            if tag == qn("w:p"):
                para = DocxParagraph(el, self.doc)
                style_name = (para.style.name if para.style else "Normal") or "Normal"

                # Group consecutive CodeBlock paragraphs into one fenced block.
                if style_name == CODE_BLOCK_STYLE_NAME:
                    code_lines: List[str] = []
                    while i < len(items) and items[i].tag == qn("w:p"):
                        p2 = DocxParagraph(items[i], self.doc)
                        s2 = (p2.style.name if p2.style else "Normal") or "Normal"
                        if s2 != CODE_BLOCK_STYLE_NAME:
                            break
                        code_lines.append(self._raw_text(p2))
                        i += 1
                    blocks.append("```\n" + "\n".join(code_lines) + "\n```")
                    continue

                # Group consecutive Quote paragraphs into one blockquote.
                if style_name == QUOTE_STYLE_NAME:
                    quote_paras: List[str] = []
                    while i < len(items) and items[i].tag == qn("w:p"):
                        p2 = DocxParagraph(items[i], self.doc)
                        s2 = (p2.style.name if p2.style else "Normal") or "Normal"
                        if s2 != QUOTE_STYLE_NAME:
                            break
                        rendered = self._render_inline_paragraph(p2).strip()
                        quote_paras.append(rendered)
                        i += 1
                    # Emit: each paragraph prefixed with "> "; insert a bare
                    # ">" line between paragraphs so GFM treats them as
                    # separate paragraphs inside the blockquote.
                    quote_lines: List[str] = []
                    for idx_q, qp in enumerate(quote_paras):
                        if idx_q > 0:
                            quote_lines.append(">")
                        if qp == "":
                            quote_lines.append(">")
                        else:
                            for sub in qp.split("\n"):
                                quote_lines.append(
                                    ("> " + sub) if sub else ">"
                                )
                    blocks.append("\n".join(quote_lines))
                    continue

                # Group consecutive list paragraphs into one list block.
                if self._is_list_paragraph(para):
                    list_lines, consumed = self._render_list(items, i)
                    blocks.append("\n".join(list_lines))
                    i += consumed
                    continue

                # Headings.
                m = re.match(r"^Heading (\d)$", style_name)
                if m:
                    level = int(m.group(1))
                    inline = self._render_inline_paragraph(para)
                    blocks.append(("#" * level) + " " + inline.strip())
                    i += 1
                    continue

                # Horizontal rule: empty paragraph with bottom border.
                if self._is_hr(para):
                    blocks.append("---")
                    i += 1
                    continue

                # Default: normal paragraph.
                inline = self._render_inline_paragraph(para)
                if inline.strip() == "":
                    # Skip stray empty paragraphs to avoid double blank lines.
                    i += 1
                    continue
                blocks.append(inline)
                i += 1

            elif tag == qn("w:tbl"):
                table = DocxTable(el, self.doc)
                blocks.append(self._render_table(table))
                i += 1

            else:
                # Section properties, bookmarks etc - ignored.
                i += 1

        # All inline images should have been consumed.
        if self._image_counter != len(self.image_filenames):
            fail(
                f"{self.docx_path.name}: rendered {self._image_counter} of "
                f"{len(self.image_filenames)} inline images. Document "
                f"structure may contain unsupported image placement."
            )

        # Join with blank lines between blocks. Trailing newline.
        return "\n\n".join(blocks).rstrip() + "\n"

    # --- list rendering ---

    def _is_list_paragraph(self, para: DocxParagraph) -> bool:
        style_name = (para.style.name if para.style else "") or ""
        if style_name in ("List Bullet", "List Number", "List Paragraph"):
            return True
        # User-added Word lists may use numPr without our style names.
        return self._numpr(para) is not None

    def _numpr(self, para: DocxParagraph):
        pPr = para._p.find(qn("w:pPr"))
        if pPr is None:
            return None
        return pPr.find(qn("w:numPr"))

    def _list_level(self, para: DocxParagraph) -> int:
        numPr = self._numpr(para)
        if numPr is not None:
            ilvl_el = numPr.find(qn("w:ilvl"))
            if ilvl_el is not None:
                try:
                    return max(0, int(ilvl_el.get(qn("w:val"), "0")))
                except ValueError:
                    pass
        # Fall back to left_indent (set by --fromgit for nested levels).
        indent = para.paragraph_format.left_indent
        if indent is None:
            return 0
        # indent is in EMU (twentieths of a point); python-docx returns Emu.
        try:
            indent_cm = indent.cm
        except AttributeError:
            indent_cm = float(indent) / 360000.0
        if indent_cm <= LIST_INDENT_STEP_CM + 0.01:
            return 0
        return max(0, int(round((indent_cm - LIST_INDENT_STEP_CM) /
                                LIST_INDENT_STEP_CM)))

    def _is_ordered_list(self, para: DocxParagraph) -> bool:
        style_name = (para.style.name if para.style else "") or ""
        if style_name == "List Number":
            return True
        if style_name == "List Bullet":
            return False
        # User-added list via Word UI: inspect numbering definition. As a
        # heuristic, treat anything with numPr but unknown style as bullet
        # unless we can determine otherwise. Determining ordered/bullet from
        # numId requires walking word/numbering.xml; for v1, default to
        # bullet which matches the common Word UI default.
        return False

    def _render_list(self, items: List, start: int) -> Tuple[List[str], int]:
        lines: List[str] = []
        # Track ordering counter per level so ordered list markers count up.
        counters: Dict[int, int] = {}
        last_level = -1
        consumed = 0
        # Each ordered list at each level should restart at 1 when we move
        # to a new list. We reset a level's counter when we leave it (move
        # to a shallower level) or when its parent type changes.

        i = start
        while i < len(items) and items[i].tag == qn("w:p"):
            para = DocxParagraph(items[i], self.doc)
            if not self._is_list_paragraph(para):
                break
            level = self._list_level(para)
            ordered = self._is_ordered_list(para)
            # Reset counters for any level deeper than current (we're moving
            # back up the tree).
            for deeper in [lvl for lvl in counters if lvl > level]:
                del counters[deeper]
            if ordered:
                counters[level] = counters.get(level, 0) + 1
                marker = f"{counters[level]}. "
            else:
                # Bullets do not need counting; clear any prior ordered count
                # at this level so a subsequent ordered list restarts at 1.
                counters.pop(level, None)
                marker = "- "
            indent = "  " * level
            content = self._render_inline_paragraph(para).strip()
            # Detect task list pattern. After inline rendering, brackets have
            # been escaped to "\[" and "\]"; rewrite the leading marker so it
            # round-trips as a real GFM task-list item.
            task_match = re.match(
                r"^\\\[(?P<state>[ xX])\\\] ?(?P<rest>.*)$",
                content,
            )
            if task_match:
                state = task_match.group("state").lower()
                if state == " ":
                    state = " "
                content = f"[{state}] {task_match.group('rest')}"
            lines.append(f"{indent}{marker}{content}")
            i += 1
            consumed += 1
            last_level = level
        return lines, consumed

    # --- table rendering ---

    def _render_table(self, table: DocxTable) -> str:
        # Validate: no merged cells, no nested tables, no images in cells.
        rows = table.rows
        if not rows:
            return ""
        for row in rows:
            for cell in row.cells:
                tc = cell._tc
                tcPr = tc.find(qn("w:tcPr"))
                if tcPr is not None:
                    if tcPr.find(qn("w:gridSpan")) is not None:
                        fail(
                            f"{self.docx_path.name} contains a table with "
                            f"horizontally merged cells (gridSpan), which "
                            f"GFM pipe tables cannot represent."
                        )
                    if tcPr.find(qn("w:vMerge")) is not None:
                        fail(
                            f"{self.docx_path.name} contains a table with "
                            f"vertically merged cells (vMerge), which GFM "
                            f"pipe tables cannot represent."
                        )
                if tc.find(f".//{{{NS['w']}}}tbl") is not None:
                    fail(
                        f"{self.docx_path.name} contains a nested table, "
                        f"which is unsupported."
                    )
                if tc.find(f".//{{{NS['w']}}}drawing") is not None:
                    fail(
                        f"{self.docx_path.name} contains an image inside a "
                        f"table cell, which is unsupported."
                    )

        ncols = len(rows[0].cells)
        # Detect column count mismatch across rows: also unsupported.
        for r in rows[1:]:
            if len(r.cells) != ncols:
                fail(
                    f"{self.docx_path.name} contains a table with ragged "
                    f"row widths; cannot map to a GFM pipe table."
                )

        # First row is the header. Build cell text and per-column alignment.
        header_cells: List[str] = []
        col_aligns: List[str] = []
        for cell in rows[0].cells:
            header_cells.append(self._render_cell(cell, in_header=True))
            align = self._cell_alignment(cell)
            col_aligns.append(align)

        body_rows: List[List[str]] = []
        for row in rows[1:]:
            body_rows.append([self._render_cell(c) for c in row.cells])

        # Compose pipe table.
        out = ["| " + " | ".join(header_cells) + " |"]
        sep_parts = []
        for align in col_aligns:
            if align == "center":
                sep_parts.append(":---:")
            elif align == "right":
                sep_parts.append("---:")
            elif align == "left":
                sep_parts.append(":---")
            else:
                sep_parts.append("---")
        out.append("| " + " | ".join(sep_parts) + " |")
        for body in body_rows:
            out.append("| " + " | ".join(body) + " |")
        return "\n".join(out)

    def _render_cell(self, cell, *, in_header: bool = False) -> str:
        pieces: List[str] = []
        for p in cell.paragraphs:
            inline = self._render_inline_paragraph(
                p, in_table=True, ignore_bold=in_header,
            )
            if inline.strip():
                pieces.append(inline.strip())
        return " ".join(pieces) if pieces else ""

    def _cell_alignment(self, cell) -> Optional[str]:
        for p in cell.paragraphs:
            jc = p._p.find(f".//{qn('w:jc')}")
            if jc is not None:
                val = jc.get(qn("w:val"), "")
                if val in ("center", "right", "left"):
                    return val
            if p.alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
                return "center"
            if p.alignment == WD_PARAGRAPH_ALIGNMENT.RIGHT:
                return "right"
            if p.alignment == WD_PARAGRAPH_ALIGNMENT.LEFT:
                return "left"
        return None

    # --- HR and miscellaneous ---

    def _is_hr(self, para: DocxParagraph) -> bool:
        if para.text.strip() != "":
            return False
        pPr = para._p.find(qn("w:pPr"))
        if pPr is None:
            return False
        pBdr = pPr.find(qn("w:pBdr"))
        if pBdr is None:
            return False
        return pBdr.find(qn("w:bottom")) is not None

    def _raw_text(self, para: DocxParagraph) -> str:
        """Return paragraph text verbatim (used for code blocks)."""
        # Walk all w:t and w:tab/w:br nodes in order.
        parts: List[str] = []
        for el in para._p.iter():
            t = el.tag
            if t == qn("w:t"):
                parts.append(el.text or "")
            elif t == qn("w:tab"):
                parts.append("\t")
            elif t == qn("w:br"):
                parts.append("\n")
        return "".join(parts)

    # --- inline paragraph rendering (bold/italic/code/links/images) ---

    def _render_inline_paragraph(self, para: DocxParagraph, *,
                                 in_table: bool = False,
                                 ignore_bold: bool = False) -> str:
        pieces = self._collect_pieces(para)
        if ignore_bold:
            for piece in pieces:
                piece.bold = False
        return self._emit_pieces(pieces, in_table=in_table)

    def _collect_pieces(self, para: DocxParagraph) -> List[_RunPiece]:
        """Walk paragraph XML and produce a flat list of run pieces."""
        pieces: List[_RunPiece] = []

        def walk(node, link_url: Optional[str]) -> None:
            for child in node:
                tag = child.tag
                if tag == qn("w:hyperlink"):
                    # Resolve r:id to URL via paragraph's part rels.
                    r_id = child.get(qn("r:id"))
                    url = None
                    if r_id:
                        try:
                            url = para.part.rels[r_id].target_ref
                        except KeyError:
                            url = None
                    walk(child, url)
                elif tag == qn("w:r"):
                    self._process_run(child, pieces, link_url)
                elif tag == qn("w:smartTag") or tag == qn("w:ins"):
                    # Tracked changes already rejected earlier; smartTag
                    # is a wrapper around runs.
                    walk(child, link_url)
                # other children (bookmarks, comment refs, etc) ignored

        walk(para._p, None)
        return pieces

    def _process_run(self, r_el, pieces: List[_RunPiece],
                     link_url: Optional[str]) -> None:
        rPr = r_el.find(qn("w:rPr"))
        bold = False
        italic = False
        code = False
        strike = False
        if rPr is not None:
            # Style reference for inline code.
            rStyle = rPr.find(qn("w:rStyle"))
            if rStyle is not None:
                if rStyle.get(qn("w:val"), "") == INLINE_CODE_STYLE_NAME:
                    code = True
            # Bold: <w:b/> or <w:b w:val="true"/>; <w:b w:val="false"/> means off.
            b = rPr.find(qn("w:b"))
            if b is not None and b.get(qn("w:val"), "true").lower() not in ("0", "false"):
                bold = True
            i = rPr.find(qn("w:i"))
            if i is not None and i.get(qn("w:val"), "true").lower() not in ("0", "false"):
                italic = True
            s = rPr.find(qn("w:strike"))
            if s is not None and s.get(qn("w:val"), "true").lower() not in ("0", "false"):
                strike = True

        # Walk run children in order for text, breaks, drawings.
        for child in r_el:
            tag = child.tag
            if tag == qn("w:t"):
                if child.text:
                    pieces.append(_RunPiece(
                        text=child.text,
                        bold=bold, italic=italic, code=code, strike=strike,
                        link=link_url,
                    ))
            elif tag == qn("w:tab"):
                pieces.append(_RunPiece(
                    text="\t",
                    bold=bold, italic=italic, code=code, strike=strike,
                    link=link_url,
                ))
            elif tag == qn("w:br"):
                # Markdown line break: two trailing spaces + newline. But
                # for v1 we collapse to a space to keep paragraph flow
                # readable; explicit hard breaks remain rare in practice.
                pieces.append(_RunPiece(
                    text="  \n",
                    bold=False, italic=False, code=False, strike=False,
                    link=None,
                ))
            elif tag == qn("w:drawing"):
                # Consume an image slot.
                if self._image_counter >= len(self.image_filenames):
                    fail(
                        f"{self.docx_path.name}: encountered more inline "
                        f"images than were extracted."
                    )
                idx = self._image_counter
                pieces.append(_RunPiece(is_image=True, image_index=idx,
                                        link=link_url))
                self._image_counter += 1

    def _emit_pieces(self, pieces: List[_RunPiece], *,
                     in_table: bool) -> str:
        # First, coalesce adjacent same-format text pieces (excluding images).
        out: List[str] = []
        i = 0
        while i < len(pieces):
            p = pieces[i]
            if p.is_image:
                out.append(self._format_image(p))
                i += 1
                continue
            # Hard break tokens are preserved verbatim.
            if p.text == "  \n":
                out.append(p.text)
                i += 1
                continue
            # Coalesce same-format run.
            j = i
            buf = []
            key = p.fmt_key()
            while j < len(pieces):
                q = pieces[j]
                if q.is_image or q.text == "  \n" or q.fmt_key() != key:
                    break
                buf.append(q.text)
                j += 1
            text = "".join(buf)
            out.append(self._format_text_run(text, key, in_table=in_table))
            i = j
        return "".join(out)

    def _format_text_run(self, text: str, fmt_key: Tuple,
                         *, in_table: bool) -> str:
        bold, italic, code, strike, link = fmt_key
        if not text:
            return ""
        if code:
            # Inline code: escape backticks by widening fence.
            fence = "`"
            while fence in text:
                fence += "`"
            # Pad if text starts/ends with backtick.
            pad = " " if text.startswith("`") or text.endswith("`") else ""
            rendered = f"{fence}{pad}{text}{pad}{fence}"
        else:
            rendered = _escape_md_text(text, in_table=in_table)
            if bold and italic:
                rendered = f"***{rendered}***"
            elif bold:
                rendered = f"**{rendered}**"
            elif italic:
                rendered = f"*{rendered}*"
            if strike:
                rendered = f"~~{rendered}~~"
        if link:
            # Hyperlink wraps the rendered content.
            rendered = f"[{rendered}]({link})"
        return rendered

    def _format_image(self, piece: _RunPiece) -> str:
        from urllib.parse import quote
        idx = piece.image_index
        filename = self.image_filenames[idx]
        alt = self._image_alts[idx] if idx < len(self._image_alts) else ""
        # URL-encode the path so spaces and other special characters survive
        # rendering. Keep "/" unencoded so the assets directory separator
        # remains readable.
        rel = f"{ASSETS_DIR_NAME}/{quote(filename, safe='-_.')}"
        md = f"![{_escape_md_text(alt)}]({rel})"
        if piece.link:
            md = f"[{md}]({piece.link})"
        return md


# ---------------------------------------------------------------------------
# Conversion orchestration
# ---------------------------------------------------------------------------

def _walk_inline_drawings(doc):
    """Yield <w:drawing> elements in document body order."""
    body = doc.element.body
    for drawing in body.iter(qn("w:drawing")):
        yield drawing


def _existing_assets_for_doc(git_dir: Path, md_name: str,
                             manifest_entry: Optional[Dict]) -> List[Path]:
    assets_dir = git_dir / ASSETS_DIR_NAME
    if not assets_dir.is_dir():
        return []
    known: List[Path] = []
    if manifest_entry:
        for img in manifest_entry.get("images", []):
            p = assets_dir / img["filename"]
            if p.exists():
                known.append(p)
    return known


def _allocate_image_filename(prefix: str, ext: str,
                             used: set, manifest_entry: Optional[Dict],
                             assets_dir: Path) -> str:
    """Pick a new image filename for the document, skipping any number
    currently in use either in the manifest or as a file on disk."""
    taken_numbers: set = set()
    # From the manifest (this document's previous images).
    if manifest_entry:
        for img in manifest_entry.get("images", []):
            m = re.match(rf"^{re.escape(prefix)}-image(\d+)\.[A-Za-z0-9]+$",
                         img["filename"])
            if m:
                taken_numbers.add(int(m.group(1)))
    # From disk (any file matching the prefix).
    if assets_dir.is_dir():
        for p in assets_dir.iterdir():
            m = re.match(
                rf"^{re.escape(prefix)}-image(\d+)\.[A-Za-z0-9]+$",
                p.name,
            )
            if m:
                taken_numbers.add(int(m.group(1)))
    # From the set we have already allocated in this run.
    for name in used:
        m = re.match(rf"^{re.escape(prefix)}-image(\d+)\.[A-Za-z0-9]+$", name)
        if m:
            taken_numbers.add(int(m.group(1)))
    n = 1
    while n in taken_numbers:
        n += 1
    return f"{prefix}-image{n}.{ext}"


def convert_docx_to_markdown(docx_path: Path, md_path: Path, git_dir: Path,
                             manifest: Dict) -> Dict[str, int]:
    """Convert a single .docx to .md, manage assets, update manifest.

    Returns counters: extracted, reused, removed.
    """
    counters = {"extracted": 0, "reused": 0, "removed": 0}

    # Open document.
    try:
        doc = Document(str(docx_path))
    except Exception as exc:
        # python-docx raises a variety of errors; treat all as fatal.
        fail(
            f"Failed to read {docx_path.name}: {exc}. "
            f"If the file is open in Word, save and try again."
        )
        return counters  # unreachable

    # Tracked changes: refuse to convert.
    _check_tracked_changes(doc, docx_path)

    # Extract and render all inline images in document order.
    extracted: List[ExtractedImage] = []
    for drawing_el in _walk_inline_drawings(doc):
        # Reject if the drawing is anchored (floating). Anchored elements
        # do not live under w:r/w:drawing inline; but be defensive.
        if drawing_el.getparent() is None:
            continue
        img = _extract_inline_image(drawing_el, doc, docx_path)
        extracted.append(img)

    # Resolve filenames for each extracted image by hash, reusing existing
    # files where the content matches.
    md_name = md_path.name
    md_stem = md_path.stem
    manifest_entry = manifest["documents"].get(md_name)
    prev_by_hash: Dict[str, str] = {}
    prev_filenames: set = set()
    if manifest_entry:
        for entry in manifest_entry.get("images", []):
            prev_by_hash[entry["hash"]] = entry["filename"]
            prev_filenames.add(entry["filename"])

    assets_dir = git_dir / ASSETS_DIR_NAME

    final_filenames: List[str] = []
    final_alts: List[str] = []
    new_image_records: List[Dict] = []
    files_to_write: List[Tuple[Path, bytes]] = []
    used_in_this_run: set = set()

    for img in extracted:
        if img.hash_hex in prev_by_hash:
            # Reuse existing filename. Verify the file still exists; if it
            # was deleted out-of-band, re-create it.
            fname = prev_by_hash[img.hash_hex]
            target = assets_dir / fname
            if target.exists():
                counters["reused"] += 1
            else:
                files_to_write.append((target, img.rendered_bytes))
                counters["extracted"] += 1
        else:
            fname = _allocate_image_filename(
                md_stem, img.ext, used_in_this_run, manifest_entry, assets_dir
            )
            target = assets_dir / fname
            files_to_write.append((target, img.rendered_bytes))
            counters["extracted"] += 1

        used_in_this_run.add(fname)
        final_filenames.append(fname)
        final_alts.append(img.alt_text)
        new_image_records.append({
            "filename": fname,
            "hash": img.hash_hex,
            "width": img.width_px,
            "height": img.height_px,
        })

    # Render markdown.
    converter = DocxToMarkdown(doc, docx_path, final_filenames)
    # Stash alts on the converter for emission.
    converter._image_alts = final_alts  # type: ignore[attr-defined]
    md_text = converter.render()

    # Validate that every referenced image will exist on disk after writes.
    referenced = set(final_filenames)
    referenced_on_disk = set()
    for fname in referenced:
        target = assets_dir / fname
        if target.exists() or any(p == target for p, _ in files_to_write):
            referenced_on_disk.add(fname)
    missing = referenced - referenced_on_disk
    if missing:
        fail(
            f"{docx_path.name}: markdown references images that would not "
            f"exist after write: {sorted(missing)}"
        )

    # Write images first (so the markdown points at real files).
    assets_dir.mkdir(parents=True, exist_ok=True)
    for target, blob in files_to_write:
        try:
            tmp = target.with_suffix(target.suffix + ".tmp")
            with open(tmp, "wb") as fp:
                fp.write(blob)
            os.replace(tmp, target)
        except OSError as exc:
            fail(f"Failed to write image {target}: {exc}")

    # Write markdown atomically.
    try:
        tmp_md = md_path.with_suffix(md_path.suffix + ".tmp")
        with open(tmp_md, "w", encoding="utf-8", newline="\n") as fp:
            fp.write(md_text)
        os.replace(tmp_md, md_path)
    except OSError as exc:
        # Clean up tmp on failure.
        try:
            if tmp_md.exists():
                tmp_md.unlink()
        except OSError:
            pass
        fail(f"Failed to write {md_path}: {exc}")

    # Now safe to remove stale image files for THIS document only.
    stale = prev_filenames - used_in_this_run
    for stale_name in stale:
        stale_path = assets_dir / stale_name
        if stale_path.exists():
            try:
                stale_path.unlink()
                counters["removed"] += 1
            except OSError as exc:
                fail(f"Failed to remove stale asset {stale_path}: {exc}")

    # Update manifest entry.
    manifest["documents"][md_name] = {
        "images": new_image_records,
        "updated": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    }

    return counters


# ---------------------------------------------------------------------------
# Per-run orchestration
# ---------------------------------------------------------------------------

def _is_word_temp(filename: str) -> bool:
    return filename.startswith("~$")


def _list_docx_files(work_dir: Path) -> List[Path]:
    return sorted(
        p for p in work_dir.iterdir()
        if p.is_file()
        and p.suffix.lower() == ".docx"
        and not _is_word_temp(p.name)
    )


def _list_md_files(git_dir: Path) -> List[Path]:
    return sorted(
        p for p in git_dir.iterdir()
        if p.is_file() and p.suffix.lower() == ".md"
    )


def _build_case_insensitive_md_map(md_files: List[Path]) -> Dict[str, Path]:
    """Map stem.lower() -> md Path, failing on ambiguous duplicates."""
    result: Dict[str, Path] = {}
    seen: Dict[str, List[Path]] = {}
    for p in md_files:
        key = p.stem.lower()
        seen.setdefault(key, []).append(p)
    for key, ps in seen.items():
        if len(ps) > 1:
            names = ", ".join(p.name for p in ps)
            fail(
                f"Ambiguous .md files differing only in case: {names}. "
                f"Rename so stems are unique case-insensitively."
            )
        result[key] = ps[0]
    return result


def _check_docx_ambiguity(docx_files: List[Path]) -> None:
    seen: Dict[str, List[Path]] = {}
    for p in docx_files:
        seen.setdefault(p.stem.lower(), []).append(p)
    for key, ps in seen.items():
        if len(ps) > 1:
            names = ", ".join(p.name for p in ps)
            fail(
                f"Ambiguous .docx files differing only in case: {names}. "
                f"Rename so stems are unique case-insensitively."
            )


def run_togit(cfg: Dict[str, str]) -> None:
    git_dir = Path(cfg["GitDir"])
    work_dir = Path(cfg["WorkDir"])

    if not git_dir.is_dir():
        fail(f"GitDir does not exist or is not a directory: {git_dir}")
    if not work_dir.is_dir():
        fail(f"WorkDir does not exist or is not a directory: {work_dir}")

    docx_files = _list_docx_files(work_dir)
    md_files = _list_md_files(git_dir)

    _check_docx_ambiguity(docx_files)
    md_map = _build_case_insensitive_md_map(md_files)

    if not docx_files:
        print(f"No .docx files found in {work_dir}.")
        return

    manifest = load_manifest(git_dir)
    state = load_state(work_dir)

    totals = {
        "converted": 0,
        "unchanged": 0,
        "skipped": 0,
        "extracted": 0,
        "reused": 0,
        "removed": 0,
        "errors": 0,
    }

    for docx in docx_files:
        key = docx.stem.lower()
        md_match = md_map.get(key)
        if md_match is None:
            print(f"Skipping (no matching .md): {docx.name}")
            totals["skipped"] += 1
            continue

        if _docx_unchanged_since_last_sync(docx, state):
            print(f"Unchanged, skipping: {docx.name}")
            totals["unchanged"] += 1
            continue

        print(f"Converting: {docx.name} -> {md_match.name}")
        counters = convert_docx_to_markdown(docx, md_match, git_dir, manifest)
        totals["converted"] += 1
        totals["extracted"] += counters["extracted"]
        totals["reused"] += counters["reused"]
        totals["removed"] += counters["removed"]

        # Record this file's sync state and persist BOTH manifest and state
        # immediately. If a later file in the batch fails fast, the work
        # done up to that point is preserved on disk: re-running --togit
        # will skip already-synced files and retry the failed one.
        _record_sync(docx, state)
        save_manifest(git_dir, manifest)
        save_state(work_dir, state)

        print(f"Converted successfully: {md_match.name}")

    print()
    print(f"Converted:        {totals['converted']}")
    print(f"Unchanged:        {totals['unchanged']}")
    print(f"Skipped:          {totals['skipped']}")
    print(f"Images extracted: {totals['extracted']}")
    print(f"Images reused:    {totals['reused']}")
    print(f"Images removed:   {totals['removed']}")
    print(f"Errors:           {totals['errors']}")


# ---------------------------------------------------------------------------
# --fromgit operation
# ---------------------------------------------------------------------------

def run_fromgit(cfg: Dict[str, str]) -> None:
    git_dir = Path(cfg["GitDir"])
    work_dir = Path(cfg["WorkDir"])

    if not git_dir.is_dir():
        fail(f"GitDir does not exist or is not a directory: {git_dir}")
    if not work_dir.is_dir():
        fail(f"WorkDir does not exist or is not a directory: {work_dir}")

    # Root-level .md files only. Case-insensitive match for the extension on
    # Windows (which is case-insensitive anyway), but we accept any case here.
    md_files: List[Path] = sorted(
        p for p in git_dir.iterdir()
        if p.is_file() and p.suffix.lower() == ".md"
    )

    if not md_files:
        print(f"No .md files found in {git_dir}.")
        return

    created = 0
    skipped = 0

    for md in md_files:
        docx_name = md.stem + ".docx"
        docx_path = work_dir / docx_name
        if docx_path.exists():
            print(f"Skipping (already exists): {docx_path.name}")
            skipped += 1
            continue
        print(f"Converting: {md.name} -> {docx_name}")
        convert_markdown_to_docx(md, docx_path)
        created += 1

    print(f"\nDone. Created: {created}, Skipped: {skipped}.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="mdtool",
        description=(
            "Manual markdown/Word working-copy utility. "
            "Operates on an already-synchronised local repository."
        ),
    )
    parser.add_argument(
        "--config",
        action="store_true",
        help="Enter interactive configuration mode.",
    )
    parser.add_argument(
        "--fromgit",
        action="store_true",
        help="Generate .docx working files from .md files in GitDir.",
    )
    parser.add_argument(
        "--togit",
        action="store_true",
        help="Update .md files in GitDir from edited .docx files in WorkDir.",
    )
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> None:
    args = parse_args(sys.argv[1:] if argv is None else argv)

    cfg_path = config_path()
    existing_cfg = load_config(cfg_path)

    # Configuration mode is triggered if config is missing OR --config flag.
    if args.config or existing_cfg is None:
        run_config_mode(existing_cfg)
        return  # run_config_mode calls sys.exit; safety net only.

    # Validate present config has both required keys.
    missing = [k for k in CONFIG_KEYS if k not in existing_cfg or not existing_cfg[k]]
    if missing:
        fail(
            f"Config {cfg_path} is missing required keys: {', '.join(missing)}. "
            f"Re-run with --config."
        )

    # Require an operational parameter when not in config mode.
    if args.fromgit and args.togit:
        fail("--fromgit and --togit are mutually exclusive.")
    if not args.fromgit and not args.togit:
        fail("No operation specified. Pass --fromgit, --togit, or --config.")

    if args.fromgit:
        run_fromgit(existing_cfg)
    else:
        run_togit(existing_cfg)


if __name__ == "__main__":
    main()