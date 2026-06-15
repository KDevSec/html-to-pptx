---
name: html-to-pptx
description: >-
  Convert a single self-contained HTML page or "figure" (architecture diagrams,
  dark-glass dashboards, data 大屏, infographics, posters, any one-page .html
  design or mockup) into a high-fidelity, EDITABLE PowerPoint .pptx that looks
  identical to the source and opens cleanly across PowerPoint, Keynote,
  LibreOffice, and Google Slides. Strategy = fidelity-first hybrid: whatever
  PowerPoint can draw 1:1 (text, solid/linear-gradient shapes, lines, dashed
  borders, dots, chips) becomes native editable objects; whatever it cannot
  (backdrop-blur glass, outer glow, radial/conic gradients, complex SVG icons,
  mesh/topology graphics, textured backgrounds) is cut as a transparent PNG and
  overlaid, with all TEXT kept native and editable on top. STRONGLY prefer this
  skill whenever the user points at a .html file (or pastes HTML) and wants it as
  PPT / PowerPoint / slides / 幻灯片 / 可编辑 PPT, or says "把这个 HTML / 网页 /
  设计图 / 大屏 转成（可编辑）PPT", "完全复刻 HTML 到 pptx", "html 转 ppt",
  "把这张图做成可编辑的 PPT", "convert this html / webpage / mockup to an editable
  slide or deck", or just "make this html a slide" — even when they never say the
  word "editable". This is for HTML→PPTX; for Markdown→报告 PPTX use the
  md-to-pptx skill instead. Pipeline: headless Chromium geometry capture +
  python-pptx assembly + LibreOffice verification.
---

# HTML → 可编辑 PPTX（完全复刻）

Turn a self-contained HTML "figure" page into a PowerPoint slide that looks the
same **and** is editable. The hard truth this skill is built around: PowerPoint
cannot render some things browsers can (frosted glass, glow, fancy gradients,
arbitrary SVG). Trying to fake those with native shapes looks flat and gets
rejected. So we split every element two ways and get the best of both.

## Core principle: fidelity-first hybrid

For each visual element, ask **"can PowerPoint draw this 1:1?"**

- **Yes → native** (editable shape/text). Cheap, crisp, editable.
- **No → transparent PNG cutout**, sliced from the real browser render and laid on
  top at its exact rect. Pixel-perfect.

**Text is special: it is ALWAYS native and ALWAYS on top**, never baked into a
cutout — that is what keeps the deck editable even when the panel behind it is an
image. A "glass card with a title" becomes: PNG of the glass card (no text) +
native textbox of the title placed over it.

When you are unsure whether a native approximation will look identical, **cut the
PNG** — the user asked for 完全复刻 (exact replication), so bias toward fidelity.

### Decision table

| Element | PPT native? | Action |
|---|---|---|
| Any text / label / number | yes | **native textbox, on top, editable** |
| Solid fill rect / rounded rect | yes | native `rrect`/`rect` |
| 1–2 stop **linear** gradient fill | mostly | native (inject gradient) or cut if critical |
| Border, **dashed** border, thin line | yes | native line / `dash='dash'` |
| Solid dot / circle / oval (e.g. LED) | yes | native `oval` |
| Small solid chip / badge / tag | yes | native `chip` |
| Backdrop-blur "glass" panel | **no** | **PNG cutout** (frame only), text native over it |
| Outer glow / drop-shadow halo | **no** | bake into the PNG cutout |
| Radial / conic / 3+ stop gradient | **no** (unreliable) | **PNG cutout** |
| Complex multi-path SVG icon | **no** (no vector import) | **PNG cutout** (icon tile, replaceable) |
| Mesh / topology / particle / chart svg | **no** | **PNG cutout** |
| Background texture (dot grid, noise) | **no** | PNG cutout, or flat near-color if faint |

Full reasoning, edge cases, and gotchas: `references/playbook.md` (read it before
your first build).

## Coordinate system (memorize)

- Target deck is 16:9 = **12192000 × 6858000 EMU** = 13.333 × 7.5 in.
- Capture the page at **1280×720, deviceScaleFactor=1**. Then **1 css px = 9525
  EMU** and **font pt = css px × 0.75**.
- Therefore every `getBoundingClientRect()` value ×9525 lands exactly on the deck.
  Native text, native shapes, and PNG slices that come from the *same* capture all
  align automatically — never measure twice.

`scripts/pptx_helpers.py` encodes this: `E(px)` → EMU, `P(px)` → pt.

## Workflow

Work in a scratch dir (e.g. `<figure>-work/`). Keep the source HTML untouched.

### 1. Capture geometry + reference render
```bash
python <skill>/scripts/capture.py geom INPUT.html WORK/ --region .slide --bg '#05070c'
```
`--region` is the element that is the slide (default `.slide`; use `body` if the
whole page is the figure). `--bg` is the deck background color (match the page /
template; near-black for dark decks). Produces `WORK/geo.json` (every renderable
element: rect, text, font, color, bg, border, `before_bg` for `::before` accent
bars) + `WORK/full_3x.png` (hi-res reference) + `full_1x.png`.

Now **look**: Read `full_3x.png`, open the HTML source, and skim `geo.json`. Build
a mental map of the figure's structure.

### 2. Classify every element (native vs cutout)
Walk the figure and tag each element using the decision table. Produce two lists:
- **native**: text, simple boxes, lines, dots, chips, dashed frames, accent bars.
- **cutout**: glass panels, glow, fancy gradients, SVG icons, mesh, textures.

For icons, the cutout is the icon *tile* (so it's a replaceable image). For glass
panels, the cutout is the panel frame **without its text**.

### 3. Cut the non-native pixels — two patterns

**(a) Fidelity-first board (recommended default for glassy/effect-heavy figures).**
Render the whole figure with **text + icons hidden**, transparent bg → one
pixel-perfect visual board capturing every glass/glow/gradient/mesh/accent that
PPT can't draw. Then you only re-add native text + icon images on top.
```bash
python <skill>/scripts/capture.py board INPUT.html WORK/ \
  --hide ".title,.brand .bt,.brand .bs,.blabel,.card .rn,.card .dz,.appbox span,.ic,.ci" --region .slide
```
→ `WORK/board.png` (place full-region at 0,0,1280,720 px). This is the highest-
fidelity, most cross-app-portable result; the tradeoff is that the boxes live in
the image (text stays editable, box frames don't).

**(b) Per-element slices (when you want some boxes to stay editable shapes).**
Cut only the specific elements PPT can't do; draw the rest natively.
```bash
python <skill>/scripts/capture.py slice INPUT.html WORK/ --selectors ".netfab,.glasspanel" --isolate   # decorations
python <skill>/scripts/capture.py slice INPUT.html WORK/ --selectors ".card .ic"                        # icon tiles
```
Writes `WORK/slices/<slug>_<i>.png` (transparent, 3×) + `slices.json`. `--isolate`
hides all but the selected elements (clean decoration cutouts); omit it to capture
as-rendered (icon tiles with their tint).

Choose per the user's intent: **完全复刻 / 像素级 → board (a)**; **框也要可编辑 →
slices (b)** + native shapes. You can mix (board for the busy region, native boxes
elsewhere).

### 4. Build the deck
Write `WORK/build.py`. Import the helpers, then place **back-to-front**:
background → cutout PNGs (at their rects from `geo.json`/`slices.json`) → native
shapes → native text (last, so it's on top).
```python
import sys; sys.path.insert(0, '<skill>/scripts')
from pptx_helpers import *
prs, slide = new_deck()
bg_fill(slide, '#05070c')                                   # or bg_image(...)
picture(slide, 'WORK/slices/netfab_0.png', nf['x'], nf['y'], nf['w'], nf['h'])
rrect(slide, b['x'], b['y'], b['w'], b['h'], 14, fill='#ffffff', fa=5, line='#ffffff', la=14)
textbox(slide, t['x'], t['y'], t['w']+8, t['h']+4, [[(t['text'], P(t['fs']), rgb(t['color']), True)]])
save(prs, 'output/figure.pptx')
```
Pull x/y/w/h/text/fs/color straight from `geo.json`. The helper API and assembly
patterns (alpha fills, dashed frames, two-line labels with an inline chip,
name+paren split, gradient text caveat) are in `references/playbook.md`.

### 5. Verify (and iterate)
```bash
python <skill>/scripts/verify.py output/figure.pptx --source INPUT.html --region .slide
```
Prints the editable-text/shape/pic counts (proves it's not one flat image) and
writes `compare.png` (source HTML left, built PPTX right). **Read compare.png**,
find drift (overlaps, wrong sizes, missing cutouts), fix `build.py`, rebuild.
Repeat until faithful.

> LibreOffice rendering differs slightly from real PowerPoint (fonts, kerning).
> Use it for layout/structure comparison; confirm final look in PowerPoint on the
> target machine. Embed or substitute CJK fonts there if needed.

## 跨平台兼容（PowerPoint / Keynote / LibreOffice / Google Slides）

The output is always a standard **`.pptx` (OOXML)** — one file that opens in
PowerPoint (Windows/Mac), Keynote (Mac, imports `.pptx`), LibreOffice Impress
(Linux), and Google Slides. We do **not** emit Keynote `.key` (proprietary, no
reliable writer); Mac users open the `.pptx` in Keynote or PowerPoint.

**The fidelity-first design is itself a portability win.** PowerPoint-only effects
(true outer-glow, soft-edge, 3-D, text gradients, embedded video timing) are
exactly the things Keynote and LibreOffice drop or render differently — and we
already bake those into **PNG cutouts**, which every app renders identically. So
the only real cross-app risk left is **fonts**.

- **Fonts (the #1 risk).** A font named in the file but absent on the *presenting*
  machine gets substituted → text reflows / shifts. Mitigate:
  - Set the typeface to the **audience's** OS font, not the build machine's:
    `target_cjk_font('windows')` → 微软雅黑 · `'mac'` → PingFang SC · `'linux'`
    → Noto Sans CJK SC. (Helper sets all three latin/ea/cs slots.)
  - For decks shared across machines, **embed fonts** — in PowerPoint: File ▸
    Options ▸ Save ▸ *Embed fonts in the file*; or re-save via LibreOffice with
    font embedding. Embedding is the only way to truly lock the look everywhere.
  - If a specific headline MUST look identical and can be non-editable, cut it as
    a PNG too (last resort — loses editability).
- **Verification renderer is LibreOffice on every OS** (`scripts/verify.py` finds
  it on Linux `/opt`, macOS app bundle, Windows Program Files, or PATH). It's the
  one headless renderer available cross-platform. Keynote/PowerPoint have no good
  headless render path — use them for the **final human eyeball** on the target.
- **Shapes/alpha/dashed/linear-gradient/PNG** are standard OOXML and render
  consistently across all four apps; that's why the skill sticks to them natively
  and pushes everything fancier into PNG.

## Bundled resources

- `scripts/capture.py` — Chromium geometry dump (`geom`), fidelity board (`board`), transparent cutouts (`slice`).
- `scripts/pptx_helpers.py` — python-pptx helpers (`E`,`P`,`rgb`,`new_deck`,`rrect`,`rect`,`oval`,`chip`,`textbox`,`picture`,`bg_fill`,`bg_image`,`save`,`set_font_family`,`target_cjk_font`).
- `scripts/verify.py` — LibreOffice render + side-by-side compare + editability report.
- `references/playbook.md` — the full method: coordinate math, capture technique, the native-vs-cutout rubric in depth, python-pptx assembly recipes, and a hard-won gotchas list. **Read before first build.**

## Prerequisites

`python-pptx`, `Pillow`, `playwright` (Chromium), and LibreOffice (`soffice`) +
`pdftoppm`. If `python -c "import pptx,PIL,playwright"` or `soffice --version`
fails, install them first (see `references/playbook.md` §0).
