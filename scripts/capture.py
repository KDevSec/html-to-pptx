#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Capture geometry + image slices from a self-contained HTML design page.

Why: PPTX replication needs every element's pixel rect (-> EMU) plus PNG cutouts of
the bits PowerPoint can't draw. Both come from ONE headless-Chromium pass at 1280x720
so they share a coordinate system (1 css px = 9525 EMU).

Subcommands
-----------
geom  : render the page full-bleed and dump a geometry/style inventory of every
        renderable element -> geo.json, plus full_1x.png and full_3x.png.
            python capture.py geom INPUT.html OUTDIR [--region .slide] [--bg '#05070c']
slice : produce transparent PNG cutouts (3x) for the elements PowerPoint can't
        reproduce natively (glass / glow / gradients / complex svg / mesh / icons).
            python capture.py slice INPUT.html OUTDIR --selectors ".netfab,.pcard .ic"

Notes
- --region is the element that becomes the full 16:9 frame (default '.slide'; use
  'body' if the whole page is the slide). It is pinned position:fixed inset:0 so its
  nesting depth doesn't matter and any dev chrome behind it is covered.
- Chinese / spaced file paths are fine (paths are passed via file:// as_uri()).
"""
import argparse, json, pathlib, re
from playwright.sync_api import sync_playwright
from PIL import Image

W, H = 1280, 720

def fullbleed_css(region, bg, transparent=False, only=None, hide=None):
    bgrule = 'transparent' if transparent else bg
    css = f"""<style id="cap_override">
html,body{{margin:0!important;padding:0!important;overflow:hidden!important;background:{bgrule}!important}}
{region}{{position:fixed!important;inset:0!important;left:0!important;top:0!important;
 width:100vw!important;height:100vh!important;max-width:none!important;max-height:none!important;
 margin:0!important;border-radius:0!important;border:0!important;box-shadow:none!important;transform:none!important;aspect-ratio:auto!important}}
"""
    if transparent:
        css += f"{region}::before{{display:none!important}}\n"
    if only:                                   # show ONLY these (clean decoration cutouts)
        vis = ','.join(only) + ',' + ','.join(s + ' *' for s in only)
        css += f"{region} *{{visibility:hidden!important}}\n{vis}{{visibility:visible!important}}\n"
    if hide:                                    # hide these, keep the rest (fidelity board)
        css += ','.join(hide) + f"{{visibility:hidden!important}}\n"
    return css + "</style></head>"

def write_capture(src_html, out_html, region, bg, transparent=False, only=None, hide=None):
    html = pathlib.Path(src_html).read_text(encoding='utf-8')
    html = html.replace('</head>', fullbleed_css(region, bg, transparent, only, hide))
    pathlib.Path(out_html).write_text(html, encoding='utf-8')
    return pathlib.Path(out_html).resolve()

GEOM_JS = r"""
(region) => {
  const root = document.querySelector(region) || document.body;
  const rrect = el => { const r = el.getBoundingClientRect();
    return {x:+r.x.toFixed(2), y:+r.y.toFixed(2), w:+r.width.toFixed(2), h:+r.height.toFixed(2)}; };
  const path = el => {
    const p=[]; while(el && el!==document.body && p.length<6){
      let s=el.tagName.toLowerCase();
      if(el.id) s+='#'+el.id;
      else if(el.className && typeof el.className==='string') s+='.'+el.className.trim().split(/\s+/).slice(0,3).join('.');
      p.unshift(s); el=el.parentElement;
    } return p.join('>'); };
  const directText = el => [...el.childNodes].filter(n=>n.nodeType===3).map(n=>n.textContent).join('').trim();
  const transparent = c => !c || c==='transparent' || /rgba\(0, 0, 0, 0\)/.test(c);
  const items = [];
  const all = [root, ...root.querySelectorAll('*')];
  for (const el of all) {
    const r = el.getBoundingClientRect();
    if (r.width < 2 || r.height < 2) continue;
    const cs = getComputedStyle(el);
    const tag = el.tagName.toLowerCase();
    const dt = directText(el);
    const isImg = tag==='img' || tag==='svg' || tag==='canvas';
    const hasBg = !transparent(cs.backgroundColor);
    const hasBgImg = cs.backgroundImage && cs.backgroundImage!=='none';
    const hasBorder = parseFloat(cs.borderTopWidth)>0 && !transparent(cs.borderTopColor);
    if (!(dt || isImg || hasBg || hasBgImg || hasBorder)) continue;
    items.push({
      path: path(el), tag,
      cls: (typeof el.className==='string'? el.className.trim(): ''),
      ...rrect(el),
      text: dt,
      fs: parseFloat(cs.fontSize), fw: cs.fontWeight,
      color: cs.color, bg: cs.backgroundColor,
      bgImage: hasBgImg ? cs.backgroundImage.slice(0,40) : null,
      border: hasBorder ? `${cs.borderTopWidth} ${cs.borderTopStyle} ${cs.borderTopColor}` : null,
      radius: cs.borderTopLeftRadius,
      before_bg: getComputedStyle(el,'::before').backgroundColor,   // accent bars often live here
      img: isImg,
    });
  }
  return { viewport:{w:innerWidth,h:innerHeight}, region: rrect(root), elements: items };
}
"""

def cmd_geom(a):
    out = pathlib.Path(a.outdir); out.mkdir(parents=True, exist_ok=True)
    cap = write_capture(a.input, out / '_capture.html', a.region, a.bg)
    with sync_playwright() as p:
        br = p.chromium.launch()
        pg = br.new_page(viewport={'width': W, 'height': H}, device_scale_factor=1)
        pg.goto(cap.as_uri()); pg.wait_for_timeout(a.wait)
        geo = pg.evaluate(GEOM_JS, a.region)
        (out / 'geo.json').write_text(json.dumps(geo, ensure_ascii=False, indent=1), encoding='utf-8')
        pg.screenshot(path=str(out / 'full_1x.png'))
        pg3 = br.new_page(viewport={'width': W, 'height': H}, device_scale_factor=3)
        pg3.goto(cap.as_uri()); pg3.wait_for_timeout(a.wait)
        pg3.screenshot(path=str(out / 'full_3x.png'))
        br.close()
    print(f'geom: {len(geo["elements"])} elements -> {out/"geo.json"}')
    print(f'  full_1x.png / full_3x.png saved; region={geo["region"]}')

def slug(sel, i):
    s = re.sub(r'[^a-zA-Z0-9]+', '_', sel).strip('_')
    return f'{s}_{i}'

def cmd_slice(a):
    out = pathlib.Path(a.outdir); sl = out / 'slices'; sl.mkdir(parents=True, exist_ok=True)
    sels = [s.strip() for s in a.selectors.split(',') if s.strip()]
    cap = write_capture(a.input, out / '_capture_slice.html', a.region, a.bg,
                        transparent=True, only=(sels if a.isolate else None))
    rec = {}
    with sync_playwright() as p:
        br = p.chromium.launch()
        pg = br.new_page(viewport={'width': W, 'height': H}, device_scale_factor=a.scale)
        pg.goto(cap.as_uri()); pg.wait_for_timeout(a.wait)
        for sel in sels:
            loc = pg.locator(sel)
            n = loc.count()
            for i in range(n):
                name = slug(sel, i) + '.png'
                loc.nth(i).screenshot(path=str(sl / name), omit_background=True)
                box = loc.nth(i).bounding_box()
                rec[name] = {k: round(box[k], 2) for k in ('x', 'y', 'width', 'height')} if box else None
        br.close()
    (sl / 'slices.json').write_text(json.dumps(rec, ensure_ascii=False, indent=1), encoding='utf-8')
    print(f'slice: {len(rec)} cutouts -> {sl}')
    for k, v in rec.items(): print(f'  {k}: {v}')

def cmd_board(a):
    """Fidelity-first board: render the whole region with text+icons (and any
    other --hide selectors) hidden, transparent bg -> board.png. Place it full-region
    in the deck, then overlay native text + icon slices. Captures everything PPT
    can't draw (glass/glow/gradients/mesh/accents) in one pixel-perfect image."""
    out = pathlib.Path(a.outdir); out.mkdir(parents=True, exist_ok=True)
    hide = [s.strip() for s in a.hide.split(',') if s.strip()]
    cap = write_capture(a.input, out / '_capture_board.html', a.region, a.bg,
                        transparent=True, hide=hide)
    with sync_playwright() as p:
        br = p.chromium.launch()
        pg = br.new_page(viewport={'width': W, 'height': H}, device_scale_factor=a.scale)
        pg.goto(cap.as_uri()); pg.wait_for_timeout(a.wait)
        pg.locator(a.region).screenshot(path=str(out / 'board.png'), omit_background=True)
        br.close()
    print(f'board: hid {len(hide)} selector groups -> {out/"board.png"} (place full-region, 0,0,1280,720 px)')

if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest='cmd', required=True)
    g = sub.add_parser('geom'); g.add_argument('input'); g.add_argument('outdir')
    g.add_argument('--region', default='.slide'); g.add_argument('--bg', default='#05070c')
    g.add_argument('--wait', type=int, default=600); g.set_defaults(fn=cmd_geom)
    b = sub.add_parser('board'); b.add_argument('input'); b.add_argument('outdir')
    b.add_argument('--hide', required=True, help='comma selectors to hide (text+icons): keep everything else as the visual board')
    b.add_argument('--region', default='.slide'); b.add_argument('--bg', default='#05070c')
    b.add_argument('--scale', type=int, default=3); b.add_argument('--wait', type=int, default=600); b.set_defaults(fn=cmd_board)
    s = sub.add_parser('slice'); s.add_argument('input'); s.add_argument('outdir')
    s.add_argument('--selectors', required=True); s.add_argument('--region', default='.slide')
    s.add_argument('--bg', default='#05070c'); s.add_argument('--scale', type=int, default=3)
    s.add_argument('--isolate', action='store_true', help='hide all but the selected elements (clean decoration cutouts)')
    s.add_argument('--wait', type=int, default=500); s.set_defaults(fn=cmd_slice)
    a = ap.parse_args(); a.fn(a)
