# mdtool v1.4

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
   python mdtool.py --config
   ```

   You will be prompted for two paths:

   - **GitDir**: the local directory containing your `.md` source files (typically a Git working tree).
   - **WorkDir**: a directory where `.docx` working copies will live.

   These are written to `config.ini` next to `mdtool.py`. If `config.ini` is absent, the tool automatically enters configuration mode on first run.

## Usage

`mdtool` exposes three operation flags; exactly one must be specified per invocation. An optional modifier flag is available for `--togit`.

| Flag              | Purpose                                                                         |
|-------------------|---------------------------------------------------------------------------------|
| `--config`        | Re-run interactive configuration. Overwrites `config.ini`.                      |
| `--fromgit`       | Generate `.docx` working copies in `WorkDir` from `.md` files in `GitDir`.     |
| `--togit`         | Convert edited `.docx` files in `WorkDir` back to `.md` files in `GitDir`.     |
| `-s` / `--scale`  | *(modifier for `--togit`)* Scale extracted images to Word's display size instead of using native source resolution. |

Examples:

```
python mdtool.py --fromgit
python mdtool.py --togit
python mdtool.py --togit -s
python mdtool.py --config
```

`--fromgit` and `--togit` are mutually exclusive. If neither is supplied (and config is present), the tool exits with an error.

### Image resolution

By default, `--togit` extracts images at their native source resolution — the actual pixel dimensions of the image embedded in the `.docx` file, after any cropping applied in Word. This preserves maximum image quality and is the recommended mode.

Passing `-s` / `--scale` instead scales each image to Word's configured display size at 150 DPI. Use this if you need all images to be a predictable, document-controlled width regardless of their source resolution.

Switching between the two modes will cause all images to be regenerated in the destination `assets/` directory on the next `--togit` run, since the output pixel dimensions differ.

## General workflow

1. **Pull or update** your Markdown repository in `GitDir` using your normal Git workflow. `mdtool` does not do this for you.
2. **Run `--fromgit`** to produce `.docx` working copies of any `.md` files that do not already have one in `WorkDir`. Files that already exist in `WorkDir` are skipped, so existing edits are never overwritten.
3. **Edit** the `.docx` files in Word as normal. Save when finished. Accept or reject all tracked changes before the next step.
4. **Run `--togit`** to convert the edited `.docx` files back into `.md` files in `GitDir`. Embedded images are extracted into an `assets/` subdirectory of `GitDir`. Files that have not changed since the last sync are skipped.
5. **Review the diff** in your Git client and commit. `mdtool` will not commit on your behalf.

Round-tripping the same file repeatedly is supported: the tool tracks file state in `.mdwordtool-state.json` (in `WorkDir`) and `.mdwordtool-manifest.json` (in `GitDir`) to detect unchanged files and to reuse stable asset filenames.

## Limitations

- **Windows-only by design.** Paths, the Word integration, and case-insensitive filename handling assume Windows. Other platforms are not supported.
- **Root level only.** Only `.md` files in the top level of `GitDir` are processed. Subdirectories are ignored. There is no recursive traversal.
- **No Git or GitHub integration.** The tool does not run `git pull`, `git add`, `git commit`, `git push`, or any other Git command. You must manage version control yourself.
- **No filesystem monitoring.** The tool is invoked manually. It does not watch for changes.
- **Tracked changes block conversion.** `.docx` files containing unresolved tracked changes will not be converted by `--togit`. Accept or reject them in Word first.
- **Floating images are not supported.** Anchored or floating pictures (images not placed inline with text) are rejected with an error. Only inline images are supported.
- **Annotation shapes are composited, not preserved as vector.** Floating geometric shapes (rectangles, ellipses, lines, arrows, text boxes) that sit wholly within the bounds of an inline image are composited onto that image at the configured DPI (150 by default). Shapes that extend outside an inline image, or that appear in a paragraph without an inline image, are rejected with an error. Shapes with gradient, pattern, or texture fills are rendered with a transparent fill.
- **Source image formats.** Only PNG and JPEG source images embedded in `.docx` files are extracted. Other raster or vector formats (EMF, WMF, SVG, etc.) are not handled.
- **Markdown dialect.** `--togit` recognises the styles and formatting that `--fromgit` produces: headings (H1–H6), bold, italic, bold-italic, strikethrough, inline code, fenced code blocks, blockquotes, ordered and unordered lists (nested), GFM pipe tables with column alignment, horizontal rules, hyperlinks, and inline images. Hand-authored Markdown using exotic constructs, raw HTML, or non-standard extensions may not round-trip cleanly.
- **Word-side formatting drift.** Custom styles, manual formatting, or layout changes applied in Word that do not correspond to a Markdown construct will be lost on `--togit`.
- **No conflict resolution.** If both the `.md` file in `GitDir` and its `.docx` working copy in `WorkDir` have been edited independently, the tool does not detect or merge the divergence. Whichever operation you run last wins.
- **No undo.** Conversions overwrite the destination file once written.

## Changelog

**1.4**
- Default image output changed to native source resolution. Images are now extracted at their actual embedded pixel dimensions rather than being scaled to Word's display size.
- Added `-s` / `--scale` flag to restore the previous display-size scaling behaviour.
- Sub-pixel left/right crop artefacts introduced by Word's crop tool are now snapped to zero, preventing inconsistent widths across images that were cropped only on the vertical axis.

**1.3**
- Floating annotation shapes (rectangles, ellipses, lines, arrows, text boxes) that sit wholly within an inline image are composited onto that image during `--togit`.

**1.2**
- Image asset manifest added for stable filename reuse across syncs.

**1.1**
- `--togit` image extraction and compositing with Pillow.

**1.0**
- Initial release: `--fromgit`, `--togit`, `--config`.

## Disclaimer

This tool is provided as-is, with no warranty of any kind, express or implied. It modifies files in your Git working tree and your working directory. It can overwrite, replace, or delete content if used incorrectly or if a bug is encountered.

**Use at your own risk.** Before running `--togit`, ensure your Markdown source is committed to Git so that any unintended changes can be reverted. Before running `--fromgit`, ensure you do not have unsaved Word edits you wish to preserve. The author accepts no responsibility for lost data, corrupted files, or any other adverse outcome resulting from use of this tool.
