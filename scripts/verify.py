#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Render a built PPTX (and optionally the source HTML) to PNG and stitch a
side-by-side compare image, so you can eyeball fidelity and prove editability.

Why LibreOffice: it renders .pptx headless. Fonts/letter-spacing differ slightly
from real PowerPoint, so this is for LAYOUT comparison; confirm final look in
PowerPoint on the target machine.

    python verify.py OUT.pptx [--source INPUT.html] [--region .slide] [--out compare.png]

Finds soffice at /opt/libreoffice*/program/soffice or on PATH.
"""
import argparse, glob, os, re, shutil, subprocess, tempfile, zipfile, pathlib
from PIL import Image

def find_soffice():
    # cross-platform: Linux /opt + PATH, macOS app bundle, Windows Program Files
    cands = sorted(glob.glob('/opt/libreoffice*/program/soffice'), reverse=True)
    cands += ['/Applications/LibreOffice.app/Contents/MacOS/soffice',          # macOS
              r'C:\Program Files\LibreOffice\program\soffice.exe',             # Windows
              r'C:\Program Files (x86)\LibreOffice\program\soffice.exe']
    for c in cands:
        if os.path.exists(c): return c
    return shutil.which('soffice') or shutil.which('libreoffice')

def pptx_to_png(pptx, dpi=130):
    so = find_soffice()
    if not so: raise SystemExit('LibreOffice (soffice) not found — needed for the verify step')
    d = tempfile.mkdtemp(prefix='h2p_')
    if shutil.which('pdftoppm'):                       # preferred: pptx->pdf->png (DPI control)
        subprocess.run([so, '--headless', '--convert-to', 'pdf', '--outdir', d, pptx],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        pdf = glob.glob(os.path.join(d, '*.pdf'))[0]
        subprocess.run(['pdftoppm', '-png', '-r', str(dpi), pdf, os.path.join(d, 'p')],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        return sorted(glob.glob(os.path.join(d, 'p*.png')))[0]
    # fallback (no poppler, e.g. Windows): LibreOffice exports the first slide to PNG itself
    subprocess.run([so, '--headless', '--convert-to', 'png', '--outdir', d, pptx],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
    pngs = glob.glob(os.path.join(d, '*.png'))
    if not pngs: raise SystemExit('LibreOffice PNG export produced no file')
    return pngs[0]

def html_to_png(src, region, bg='#05070c'):
    from playwright.sync_api import sync_playwright
    html = pathlib.Path(src).read_text(encoding='utf-8')
    css = (f"<style>html,body{{margin:0!important;padding:0!important;overflow:hidden!important;background:{bg}!important}}"
           f"{region}{{position:fixed!important;inset:0!important;width:100vw!important;height:100vh!important;"
           f"max-width:none!important;margin:0!important;border-radius:0!important;border:0!important;box-shadow:none!important;aspect-ratio:auto!important}}</style></head>")
    tmp = tempfile.NamedTemporaryFile(suffix='.html', delete=False, mode='w', encoding='utf-8')
    tmp.write(html.replace('</head>', css)); tmp.close()
    png = tmp.name + '.png'
    with sync_playwright() as p:
        br = p.chromium.launch()
        pg = br.new_page(viewport={'width': 1280, 'height': 720}, device_scale_factor=2)
        pg.goto(pathlib.Path(tmp.name).resolve().as_uri()); pg.wait_for_timeout(600)
        pg.screenshot(path=png); br.close()
    return png

def report_editable(pptx):
    xml = zipfile.ZipFile(pptx).read('ppt/slides/slide1.xml').decode()
    n = len([t for t in re.findall(r'<a:t>([^<]*)</a:t>', xml) if t.strip()])
    print(f'editable: text_runs={n} shapes={xml.count("<p:sp>")} pics={xml.count("<p:pic>")}')

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx'); ap.add_argument('--source'); ap.add_argument('--region', default='.slide')
    ap.add_argument('--out', default='compare.png'); ap.add_argument('--bg', default='#05070c')
    a = ap.parse_args()
    report_editable(a.pptx)
    built = Image.open(pptx_to_png(a.pptx)).convert('RGB')
    if not a.source:
        built.save(a.out); print('rendered ->', a.out); return
    src = Image.open(html_to_png(a.source, a.region, a.bg)).convert('RGB')
    h = 720;
    def fit(im): return im.resize((int(im.width * h / im.height), h))
    a_im, b_im = fit(src), fit(built)
    gap = 24
    canvas = Image.new('RGB', (a_im.width + b_im.width + gap, h), (20, 20, 24))
    canvas.paste(a_im, (0, 0)); canvas.paste(b_im, (a_im.width + gap, 0))
    canvas.save(a.out)
    print(f'compare (left=source HTML, right=built PPTX) -> {a.out}')

if __name__ == '__main__':
    main()
