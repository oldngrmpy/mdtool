# mdtool v1.6

**Author:** [oldngrmpy](https://github.com/oldngrmpy/mdtool)
**Released:** 26 May 2026

## Description

`mdtool` is a single-file Python CLI utility that lets you manage the content of Markdown files using Microsoft Word on Windows. It converts `.md` files from a local Git repository into `.docx` working copies that can be authored, reviewed, and revised in Word, then converts those `.docx` files back into Markdown so the canonical source in Git stays as `.md`. Word is the content-management surface; Markdown remains the source of truth.

The tool operates only on already-synchronised local files. It does not push, pull, commit, or otherwise interact with Git or GitHub. The user is responsible for all version control operations.

## Setup

### Requirements

- Python 3.9 or later (uses `from __future__ import annotations` and modern typing).
- Windows (paths, line endings, and the Word workflow assume Windows; behaviour on other platforms is untested).
- Microsoft Word, for editing the generated `.docx` files.

### Installation

1. Clone or download this repository.
2. Install dependencies:

   ```
   pip install -r requirements.txt
   ```

   This installs `markdown-it-py`, `python-docx`, and `Pillow`.

3. Run the tool once to perform interactive configuration:

   ```
   python mdtool.py config
   ```

   You will be prompted for two paths:

   - **GitDir**: the local directory containing your `.md` source files (typically a Git working tree).
   - **WorkDir**: a directory where `.docx` working copies will live.

   These are written to `config.ini` in the directory from which `mdtool` is run. If `config.ini` is absent, the tool automatically enters configuration mode on first run.

## Usage

`mdtool` uses subcommands. Exactly one must be specified per invocation.

```
python mdtool.py <command> [options]
```

| Command              | Purpose                                                                            |
|----------------------|------------------------------------------------------------------------------------|
| `config`             | Re-run interactive configuration. Overwrites `config.ini`.                         |
| `fromgit`            | Generate `.docx` working copies in `WorkDir` from `.md` files in `GitDir`.        |
| `togit`              | Convert edited `.docx` files in `WorkDir` back to `.md` files in `GitDir`.        |
| `newfile <filename>` | Create a new `.md` stub in `GitDir` and matching `.docx` in `WorkDir`.            |

Examples:

```
python mdtool.py fromgit
python mdtool.py togit
python mdtool.py config
python mdtool.py newfile my-document
python mdtool.py newfile my-document.md
```

### Creating new files (`newfile`)

Use `mdtool newfile <filename>` when starting a new document through `mdtool`. It creates a `.md` stub in `GitDir` with placeholder text at all six heading levels (H1–H6), then immediately converts it to a `.docx` working copy in `WorkDir` using the same Markdown-to-DOCX conversion path as `fromgit`.

The filename argument accepts a bare stem (`my-doc`) or a name with extension (`my-doc.md` or `my-doc.docx`); the extension is stripped before use. The document title in the stub is derived from the stem with hyphens and underscores replaced by spaces and title-cased.

`newfile` fails immediately with an error if either the `.md` or the `.docx` file already exists. It will not overwrite existing content.

### Image extraction (`togit`)

`togit` extracts images at their native source resolution — the actual pixel dimensions of the image embedded in the `.docx` file, after any cropping applied in Word.

## Round-trip safety

Supported text content is designed to round-trip cleanly when edited within the styles and structures produced by `fromgit`. The following text and document-structure elements are preserved through a full `fromgit` → edit → `togit` cycle: headings (H1–H6), bold, italic, bold-italic, strikethrough, inline code, fenced code blocks, blockquotes, ordered and unordered lists (nested), GFM pipe tables with column alignment, horizontal rules, hyperlinks, and inline images (PNG and JPEG).

A few properties of the round-trip are important to understand:

**`fromgit` never overwrites an existing `.docx`.** If a working copy already exists in `WorkDir`, it is skipped entirely. This means Word edits are never silently discarded by a subsequent `fromgit` run.

**`togit` skips unchanged files.** File state is tracked in `.mdwordtool-state.json` (in `WorkDir`). A file is skipped only when its modification time and size match the recorded values from the last run.

**Asset filenames are stable across syncs.** An image manifest (`.mdwordtool-manifest.json` in `GitDir/assets/`) maps image content hashes to stable output filenames, so image references in the generated `.md` do not change between runs unless the image itself changes.

### Image handling

Images embedded in `.docx` files flow in both directions.

**`fromgit` direction (`.md` → `.docx`):** `fromgit` reads `![alt](path)` references in the Markdown source and embeds the referenced file as an inline image in the generated `.docx`. Only local PNG and JPEG files are supported. The display width is calculated from the image's embedded DPI metadata (or 96 DPI if absent), capped at 15 cm. Images narrower than 15 cm at their native DPI are placed at their native width — they are never upscaled. Alt text is preserved on the Word image object so `togit` can recover it on the return trip.

If a referenced image cannot be embedded — because the file does not exist, the path is an HTTP/HTTPS URL, or the format is unsupported — `fromgit` emits a warning to stderr and inserts a plain-text placeholder (`[Image not found: filename]` or `[Unsupported image format: filename]`) in place of the image. The conversion continues; only that image is affected.

**`togit` direction (`.docx` → `.md`):** `togit` extracts inline images from the `.docx`, writes them into `GitDir/assets/`, and emits `![alt](assets/filename)` references in the generated `.md`. Alt text set on the Word image (including alt text written by `fromgit`) is recovered and used in the markdown reference.

A few constraints apply in both directions:

- **Annotation shapes are destructively composited.** Floating geometric shapes (rectangles, ellipses, lines, arrows, text boxes) that sit over an inline image in Word are composited onto the image pixel data during `togit`. The resulting `.md` contains a flat PNG — the shapes no longer exist as editable Word objects. If you regenerate the `.docx` via `fromgit`, you will get a flat image, not the original with live shapes on top.
- **Cropping is baked in on extraction.** Word's crop is applied at `togit` time. The crop is not stored in a reversible form.
- **Only inline images are supported.** Anchored or floating pictures are rejected with an error.
- **Only PNG and JPEG are supported** in both directions. Other formats (SVG, GIF, WebP, EMF, WMF, etc.) are not handled.
- **Original git asset files are never overwritten by `togit`.** When `togit` processes a `.docx` that was generated by `fromgit`, it extracts the embedded image bytes, hashes them, and writes a new asset file with a stable hash-derived name. The original file referenced in the source `.md` is left untouched.

## General workflow

**Starting a new document:** use `mdtool newfile <name>` instead of steps 1–2. It creates the `.md` stub in `GitDir` and the `.docx` in `WorkDir` in one step, then you continue from step 3.

**Working with existing documents:**

1. **Pull or update** your Markdown repository in `GitDir` using your normal Git workflow. `mdtool` does not do this for you.
2. **Run `fromgit`** to produce `.docx` working copies of any `.md` files that do not already have one in `WorkDir`. Files that already exist in `WorkDir` are skipped, so existing edits are never overwritten. Images referenced in the `.md` files are embedded inline in the generated `.docx`.
3. **Edit** the `.docx` files in Word as normal. Save when finished. Accept or reject all tracked changes before the next step.
4. **Run `togit`** to convert the edited `.docx` files back into `.md` files in `GitDir`. Embedded images are extracted into an `assets/` subdirectory of `GitDir`. Files that have not changed since the last sync are skipped.
5. **Review the diff** in your Git client and commit. `mdtool` will not commit on your behalf.

## Limitations

- **Windows-only by design.** Paths, the Word integration, and case-insensitive filename handling assume Windows. Other platforms are not supported.
- **Root level only.** Only top-level `.md` files in `GitDir` are processed; subdirectories are ignored.
- **No Git or GitHub integration.** The tool does not run `git pull`, `git add`, `git commit`, `git push`, or any other Git command. You must manage version control yourself.
- **No filesystem monitoring.** The tool is invoked manually. It does not watch for changes.
- **Tracked changes block conversion.** `.docx` files containing unresolved tracked changes will not be converted by `togit`. Accept or reject them in Word first.
- **Floating images are not supported.** Anchored or floating pictures are rejected with an error. Only inline images are supported in both directions.
- **Annotation shapes are composited, not preserved as vector.** See [Image handling](#image-handling) above.
- **Image formats.** Only PNG and JPEG are supported in both directions. Other raster or vector formats (SVG, GIF, WebP, EMF, WMF, etc.) are not handled; `fromgit` inserts a placeholder and `togit` fails fast.
- **HTTP/HTTPS image URLs** in Markdown source are not fetched by `fromgit`. A placeholder is inserted and a warning is emitted.
- **Markdown dialect.** `togit` recognises the styles and formatting that `fromgit` produces. Hand-authored Markdown using exotic constructs, raw HTML, or non-standard extensions may not round-trip cleanly.
- **Word-side formatting drift.** Custom styles, manual formatting, or layout changes applied in Word that do not correspond to a Markdown construct will be lost on `togit`.
- **No conflict resolution.** If both the `.md` file in `GitDir` and its `.docx` working copy in `WorkDir` have been edited independently, the tool does not detect or merge the divergence. Whichever operation you run last wins.
- **No undo.** Conversions overwrite the destination file once written.

## Changelog

**1.6** *(26 May 2026)*
- `fromgit` now embeds local PNG/JPEG images inline in generated `.docx` files. Display width is derived from embedded DPI metadata (or 96 DPI fallback), capped at 15 cm with aspect ratio preserved. Images narrower than 15 cm at their native DPI are placed at native width (no upscaling).
- Alt text from `![alt](...)` is written to the Word image's alt-text field (`wp:docPr/@descr`) so `togit` recovers it on the return trip.
- Missing files, HTTP/HTTPS URLs, and unsupported formats emit a warning and insert a plain-text placeholder instead of failing the conversion.
- Image paths found by `fromgit` are recorded in the asset manifest so `togit` can allocate fresh asset names without touching the original git files.
- `togit` now reads alt text from `wp:docPr/@descr` and emits it in the `![alt](path)` reference in the generated Markdown.
- `-s` / `--scale` flag deprecated. Images always use native source resolution; the flag is a no-op and will be removed in a future version.
- Bug fix: `togit` no longer passes the removed `scale_to_display` argument to internal sync-state functions.

**1.5** *(25 May 2026)*
- Subcommand CLI: `config`, `fromgit`, `togit`, `newfile` replace the previous `--config`, `--fromgit`, `--togit` flags.
- `--scale` (`-s`) is now a subcommand option of `togit` only; passing it to any other subcommand is an error.
- `newfile <filename>` creates a `.md` stub with H1–H6 placeholder headings and converts it to a `.docx` in one step. Fails if either file already exists.
- `fromgit` now always defines all six heading styles (H1–H6) in generated `.docx` files, regardless of which heading levels appear in the source `.md`.
- Bug fix: switching between native and scaled image modes now correctly forces re-extraction on the next `togit` run.

**1.4**
- Default image output changed to native source resolution.
- Added `-s` / `--scale` flag (deprecated in 1.6).
- Sub-pixel left/right crop artefacts from Word's crop tool are snapped to zero.

**1.3**
- Floating annotation shapes composited onto host inline images during `togit`.

**1.2**
- Image asset manifest added for stable filename reuse across syncs.

**1.1**
- `togit` image extraction and compositing with Pillow.

**1.0**
- Initial release.

## Disclaimer

This tool is provided as-is, without warranty. It modifies files in both `GitDir` and `WorkDir`, and conversions may overwrite existing content.

Before running `togit`, commit or otherwise back up your Markdown source. Before running `fromgit`, ensure you do not have unsaved Word edits you need to keep.

