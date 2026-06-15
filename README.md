# html-to-pptx

A Claude Code / Agent **Skill** that converts a single self-contained **HTML page**
(architecture diagrams, dark-glass dashboards, data 大屏, infographics, posters)
into a **high-fidelity, editable PowerPoint `.pptx`** that looks identical to the
source and opens cleanly across **PowerPoint, Keynote, LibreOffice, and Google
Slides**.

## The idea: fidelity-first hybrid

PowerPoint can't render everything a browser can (frosted glass, outer glow, fancy
gradients, arbitrary SVG). Faking those with native shapes looks flat. So the skill
splits every element two ways:

- **Native** (editable) for whatever PowerPoint draws 1:1 — text, solid/linear
  shapes, lines, dashed borders, dots, chips.
- **Transparent PNG cutout** for what it can't — backdrop-blur glass, glow,
  radial/conic gradients, complex SVG icons, mesh/topology, textures.
- **Text is always native and on top**, so the deck stays editable even where the
  panel behind it is an image.

That split is also why the output is portable: the effects other apps drop are
exactly the ones baked into PNGs (which render identically everywhere).

## Install

Clone into your Claude skills directory:

```bash
git clone https://github.com/KDevSec/html-to-pptx.git ~/.claude/skills/html-to-pptx
```

It then triggers automatically when you ask to turn an HTML file into an editable
PPT. (Works the same in any agent harness that loads `SKILL.md` skills.)

## Prerequisites

`python-pptx`, `Pillow`, `playwright` (Chromium), and `LibreOffice` (`soffice`) +
`pdftoppm`. See `references/playbook.md` §0 for one-line installs.

## Usage

Point at a self-contained `.html` and ask:

> 把这个 html 转成可编辑 PPT（完全复刻）  ·  *convert this html into an editable, pixel-faithful deck*

The pipeline (also runnable by hand):

```bash
SK=~/.claude/skills/html-to-pptx
# 1. capture geometry + reference render
python $SK/scripts/capture.py geom INPUT.html WORK/ --region .slide
# 2a. fidelity-first: one visual board with text+icons hidden
python $SK/scripts/capture.py board INPUT.html WORK/ --hide "<all text + .ic,.ci selectors>"
# 2b. or per-element cutouts (keep some boxes editable)
python $SK/scripts/capture.py slice INPUT.html WORK/ --selectors ".mesh,.glasspanel" --isolate
# 3. write WORK/build.py using scripts/pptx_helpers.py  (bg -> cutouts -> native shapes -> native text)
# 4. verify
python $SK/scripts/verify.py OUT.pptx --source INPUT.html --region .slide
```

## What's inside

| Path | Purpose |
|---|---|
| `SKILL.md` | The skill: triggering, decision table (native vs PNG), coordinate system, workflow, cross-platform notes |
| `scripts/capture.py` | Headless-Chromium geometry dump (`geom`), fidelity board (`board`), transparent cutouts (`slice`) |
| `scripts/pptx_helpers.py` | python-pptx helper library (`E`/`P`/`rrect`/`textbox`/`chip`/`oval`/`picture`/`save`/`target_cjk_font`/…) |
| `scripts/verify.py` | LibreOffice render + side-by-side compare + editability report |
| `references/playbook.md` | Full method: coordinate math, capture technique, assembly recipes, gotchas |

## Fonts & cross-platform

Output is always standard `.pptx` (we don't emit Keynote `.key`). The default CJK
font is **Microsoft YaHei (微软雅黑)**; switch per audience with
`target_cjk_font('mac'|'linux')`, or embed fonts in PowerPoint for true
cross-machine fidelity. Verification uses LibreOffice on any OS; confirm the final
look in the app you'll actually present from.

## License

[Apache-2.0](LICENSE) © 2026 KDevSec
