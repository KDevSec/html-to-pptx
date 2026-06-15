# HTML → PPTX Replication Playbook

The full method behind `html-to-pptx`: how to turn a self-contained HTML "figure"
page into a faithful, editable PowerPoint deck. Read this before your first build.

## Contents
- [§0 Environment & install](#0-environment--install)
- [§1 Coordinate system](#1-coordinate-system)
- [§2 Geometry capture (CDP / getBoundingClientRect)](#2-geometry-capture)
- [§3 The native-vs-cutout decision](#3-the-native-vs-cutout-decision)
- [§4 Slicing PNG cutouts](#4-slicing-png-cutouts)
- [§5 python-pptx assembly recipes](#5-python-pptx-assembly-recipes)
- [§6 Cross-platform compatibility](#6-cross-platform-compatibility)
- [§7 Verification](#7-verification)
- [§8 Gotchas (hard-won)](#8-gotchas)

---

## §0 Environment & install

**First, every time:** run the dependency doctor — it checks everything below in one
shot and `--fix` installs the pip deps + Chromium (stdlib-only, runs even when
nothing is installed):
```bash
python scripts/check_deps.py            # report; exit 0 = all present
python scripts/check_deps.py --fix      # auto-install pip deps + chromium
```
Tiers: **CORE** (build) = python-pptx, Pillow, playwright, Chromium. **VERIFY**
(QA compare only) = LibreOffice; `pdftoppm` is **optional** (LibreOffice can export
PNG itself, so `verify.py` falls back to `soffice --convert-to png`).
What it verifies (equivalent manual checks):
```bash
python3 -c "import pptx,PIL,playwright; print('py libs ok')"   # NOT playwright.__version__ (no such attr)
soffice --version || /opt/libreoffice*/program/soffice --version   # LibreOffice
which pdftoppm   # optional
```
Install if missing:
```bash
pip install --user python-pptx Pillow playwright        # add --break-system-packages on PEP668 distros
python -m playwright install chromium                    # if no Chromium cached
# LibreOffice: apt-get install -y libreoffice-impress  (Linux) / brew install --cask libreoffice (mac) / winget install TheDocumentFoundation.LibreOffice (win)
# pdftoppm (OPTIONAL, higher-res compare): poppler-utils (Linux) / brew install poppler (mac) / skip on Windows
```
Any Chromium works (`google-chrome`, the playwright-bundled Chromium, or Edge).

## §1 Coordinate system

- Deck 16:9 = **12192000 × 6858000 EMU** = 13.333 × 7.5 in.
- Render/capture at **1280×720, deviceScaleFactor=1** → **1 css px = 9525 EMU**
  (12192000/1280, same vertically) and **pt = css px × 0.75** (72/96).
- Consequence: any `getBoundingClientRect()` px ×9525 = EMU lands on the deck.
  Capture **once**; reuse the same rects for native shapes, native text, and image
  slices → everything lines up. Never measure in two passes.
- Capture at dsf=1 for geometry (you want CSS px). Capture at dsf=3 for **pixels**
  (icon/decoration slices) so cutouts are crisp when scaled onto the deck.

## §2 Geometry capture

`scripts/capture.py geom` renders a **full-bleed** version of the page (the
`--region` element pinned `position:fixed; inset:0; 100vw×100vh`, so its nesting
depth is irrelevant and dev chrome behind it is covered) and dumps every
renderable element's `{path, rect, text, fontSize, color, bg, border, radius,
before_bg}`.

Why `before_bg`: accent bars, dividers, and glow strips are frequently CSS
`::before`/`::after` pseudo-elements — their color lives in `getComputedStyle(el,
'::before')`, not on the element. Capture it so you can recreate the accent.

Two-line labels & inline chips: an element like `<div class=ln><span class=ltag>L2
</span>基础设施<br>系统软件层</div>` needs special handling — split `innerHTML`
on `<br>`, pull the chip (`.ltag`) out as its own rect, and render line 1 *after*
the chip's right edge so they don't overlap (otherwise the L1/L2 chip sits on the
first word).

Inline aliases / parentheticals: a product name like `Token 服务平台（AIStation）`
where `（AIStation）` is a styled `.paren` span — capture the span's own rect and
render it separately so its position/size/color match.

For a complex figure, **write a tailored capture** that collects exactly the
logical elements you need (bands, cards, labels, icon rects, accent colors) into a
clean typed schema. The generic `geom` dump is great for understanding + simple
figures; a tailored collector is better when structure is rich.

## §3 The native-vs-cutout decision

Goal = 完全复刻 (look identical) **and** editable. Resolve the tension per element:

**Native (PowerPoint draws it 1:1, keep editable):**
- All text/labels/numbers — *always native, always on top.*
- Solid fills; 1–2 stop linear gradients; borders incl. dashed; thin lines.
- Circles/dots (LEDs), simple solid chips/badges, accent bars.

**PNG cutout (PowerPoint can't match — slice from the real render):**
- Backdrop-blur "glass" panels (the frame only; text goes native on top).
- Outer glow / drop-shadow halos.
- Radial / conic / 3+ stop gradients.
- Complex multi-path SVG icons (no vector import path) → icon **tiles** as images.
- Mesh / topology / particle / chart SVGs.
- Background textures (dot grids, noise) — or approximate with a flat near-color.

**Tie-breaker:** if unsure a native approximation looks identical, cut the PNG.
But never sink text into a cutout — that's the line that keeps the deck editable.

Why not fake glass with a translucent native rect? It reads flat and "off" against
the crisp source and usually gets rejected. A sliced PNG of the real glass is
pixel-identical and renders the same in every app (§6).

## §4 Slicing PNG cutouts

`scripts/capture.py slice --selectors "..."` element-screenshots each match with a
**transparent** background (`omit_background`) at 3×, and records each rect to
`slices.json`.

- `--isolate` hides everything except the selected elements before the shot — use
  for decorations (mesh, glow panels) so nothing bleeds in.
- Omit `--isolate` for icon tiles you want captured as-rendered (with their own
  tinted background).
- Place a cutout in the deck at its rect from `slices.json` (or the element's rect
  in `geo.json`) with `picture(slide, png, x, y, w, h)`.
- Transparent PNG composites correctly over the deck background in every app —
  this is why cutouts are the portable choice.

## §5 python-pptx assembly recipes

Import `scripts/pptx_helpers.py`. Place **back-to-front**: background → cutouts →
native shapes → native text (last = on top).

- **Translucent fill (glass approximation, only when staying native):** solidFill
  then append `<a:alpha val="‰"/>` — the helper does this via `fa`/`la` (0–100):
  `rrect(s, x,y,w,h, 12, fill='#ffffff', fa=8, line='#ffffff', la=20)`.
- **Dashed frame:** `rrect(..., line='#7fb4ff', la=50, dash='dash')`.
- **Accent bar:** thin solid `rrect` at `(band.x, band.y+band.h*0.13, 3.5,
  band.h*0.74)` using the captured `before_bg`.
- **Chip/badge:** `chip(s, x,y,w,h, 'L2', P(8), fill, fg)` (rounded, centered).
- **Two-line label with chip:** draw the chip at its rect; render line 1 in a
  textbox starting at `chip.x+chip.w+gap`; render remaining lines at the label's
  left x, one line height down.
- **Name + parenthetical:** strip the paren text off the name, render the name at
  its rect and the paren at the paren's own rect.
- **Title with an emphasis span:** one textbox, two runs — prefix in INK, the
  emphasis suffix in the accent color. Do **not** use a run-level gradient (§8).
- **Background:** `bg_image(slide, png)` for a captured/painted board, or
  `bg_fill(slide, '#05070c')` for a flat near-color.
- Always `prs.slide_layouts[6]` (blank). `save()` prints editable counts.

## §6 Cross-platform compatibility

Output is standard `.pptx` → opens in PowerPoint (Win/Mac), Keynote (imports
`.pptx`), LibreOffice (Linux), Google Slides. We never emit `.key`.

- **Most-portable feature set:** PNG images, native text, solid/linear-gradient
  shapes, alpha, dashed lines — render consistently across all four apps. The
  skill deliberately stays inside this set natively and pushes everything fancier
  to PNG, so Keynote/LibreOffice don't silently drop effects.
- **Fonts = the main risk.** A font absent on the *presenting* machine is
  substituted → reflow. Strategy:
  - Set the audience-OS font: `target_cjk_font('windows'|'mac'|'linux')`
    → 微软雅黑 / PingFang SC / Noto Sans CJK SC. Set all three (latin/ea/cs) slots
    (the helper does).
  - For cross-machine delivery, **embed fonts** (PowerPoint: Options ▸ Save ▸ Embed
    fonts; or LibreOffice save-with-fonts). Only embedding truly locks the look.
  - Last resort for a critical headline: cut it as PNG (loses editability).
- **Verification renderer:** LibreOffice headless on every OS (`verify.py` locates
  it cross-platform). Keynote/PowerPoint can't render headless well — use for the
  final human check on the target machine.

## §7 Verification

```bash
python scripts/verify.py OUT.pptx --source INPUT.html --region .slide
```
Prints `text_runs / shapes / pics` (prove it's not one flat image — text_runs
should be well above zero and shapes/pics reflect your build) and writes
`compare.png` (source left, built right). Read it, find drift, fix, rebuild.

Also sanity-unpack:
```python
import zipfile,re; x=zipfile.ZipFile('OUT.pptx').read('ppt/slides/slide1.xml').decode()
print('texts', len([t for t in re.findall(r'<a:t>([^<]*)</a:t>',x) if t.strip()]))
```
LibreOffice ≠ PowerPoint pixel-for-pixel (fonts/kerning). Use it for layout;
confirm final look in the user's actual app.

## §8 Gotchas

1. **Don't fake glass/glow with native shapes** — flat, gets rejected. Slice PNG;
   put text native on top. (The whole point of fidelity-first.)
2. **`MSO_SHAPE.ISOSCELES_TRIANGLE` etc. missing in python-pptx 1.0.x** — for tiny
   arrowheads use a text glyph (`▾ ↕ ↓`) in a textbox instead.
3. **Run-level text gradient** (`a:gradFill` on a run) is finicky in PowerPoint and
   usually invisible in LibreOffice → use a solid accent color for emphasis text.
4. **Two-line label / chip overlap** — render line 1 after the chip's right edge
   (see §2/§5), or the L1/L2 chip sits on top of the first word.
5. **`::before` accents** — their color isn't on the element; read it via
   `getComputedStyle(el,'::before')` (captured as `before_bg`).
6. **Full-bleed capture** — pin the region `position:fixed; inset:0` and force
   `aspect-ratio:auto`; otherwise the 16:9 `aspect-ratio` rule leaves letterbox
   gaps and your rects are offset.
7. **Spaced / CJK file paths** — load via `pathlib.Path(p).resolve().as_uri()`
   (handles encoding); don't hand-build `file://` strings.
8. **Capture at dsf=1 for geometry, dsf=3 for pixels** — mixing them scales your
   rects wrong.
9. **Near-black vs navy background** — match the *template/page* background, not an
   inner figure's gradient. A figure image may bake in its own gradient that isn't
   the slide background; check the actual page/template `background`.
10. **Source typos** — replicate text faithfully, but fix obvious typos and tell
    the user (e.g. a mislabeled unit).
11. **`playwright.__version__` raises AttributeError** — the module has no such
    attribute. Check presence with `import playwright` (or `importlib.util.find_spec`),
    never by reading `__version__`. `check_deps.py` does this right.
12. **Windows: `soffice` not on PATH** — the installer doesn't add it. `verify.py`
    and `check_deps.py` auto-probe `C:\Program Files\LibreOffice\program\soffice.exe`
    (and the macOS app bundle), so a default install is found even without PATH.
13. **Windows: pdftoppm/poppler is painful** — don't fight it. It's optional;
    `verify.py` falls back to `soffice --convert-to png` (single-slide export), so
    the QA compare still works with LibreOffice alone.
