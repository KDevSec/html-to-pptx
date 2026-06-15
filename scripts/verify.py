#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Render a built PPTX (and optionally the source HTML) to PNG, stitch a side-by-side
compare image, and prove the deck is editable.

The renderer is the platform's OWN office app where possible — that's the most
faithful check, since it's the software you'll actually present in — with
LibreOffice as the cross-platform fallback:

  Windows  PowerPoint via COM automation   (needs `pywin32` + PowerPoint installed)
  macOS    Keynote via AppleScript          (or PowerPoint)
  Linux    LibreOffice (soffice)            (+ pdftoppm for higher-res; optional)

  python verify.py OUT.pptx [--source INPUT.html] [--region .slide] [--out compare.png]
  python verify.py OUT.pptx --open                         # just open it in your office app
  python verify.py OUT.pptx --renderer powerpoint|keynote|libreoffice|auto

Note: the Windows (COM) and macOS (AppleScript) paths follow the standard automation
APIs but should be confirmed on those OSes; only the LibreOffice path is exercised on
Linux. Any renderer failure falls through to the next, then to a "open it yourself" hint.
"""
import argparse, glob, os, platform, re, shutil, subprocess, tempfile, zipfile, pathlib
from PIL import Image

def find_soffice():
    cands = sorted(glob.glob('/opt/libreoffice*/program/soffice'), reverse=True)
    cands += ['/Applications/LibreOffice.app/Contents/MacOS/soffice',
              r'C:\Program Files\LibreOffice\program\soffice.exe',
              r'C:\Program Files (x86)\LibreOffice\program\soffice.exe']
    for c in cands:
        if os.path.exists(c):
            return c
    return shutil.which('soffice') or shutil.which('libreoffice')

# ---------- renderers (each returns a PNG path or None) ----------
def render_powerpoint(pptx, _dpi=None):
    """Windows: drive PowerPoint via COM to export slide 1 as PNG (most faithful)."""
    if platform.system() != 'Windows':
        return None
    try:
        import win32com.client  # pywin32
    except Exception:
        print('  (PowerPoint render needs pywin32: pip install pywin32)')
        return None
    app = pres = None
    try:
        app = win32com.client.Dispatch('PowerPoint.Application')
        pres = app.Presentations.Open(os.path.abspath(pptx), True, False, False)  # ReadOnly, !Untitled, !WithWindow
        out = os.path.join(tempfile.mkdtemp(prefix='h2p_pp_'), 'slide.png')
        pres.Slides(1).Export(out, 'PNG', 1920, 1080)
        return out if os.path.exists(out) else None
    except Exception as e:
        print('  PowerPoint COM render failed:', e)
        return None
    finally:
        try:
            if pres: pres.Close()
            if app: app.Quit()
        except Exception:
            pass

def render_keynote(pptx, _dpi=None):
    """macOS: import into Keynote and export slide images via AppleScript."""
    if platform.system() != 'Darwin' or not os.path.exists('/Applications/Keynote.app'):
        return None
    outdir = tempfile.mkdtemp(prefix='h2p_kn_')
    script = ('tell application "Keynote"\n'
              f'  set theDoc to open POSIX file "{os.path.abspath(pptx)}"\n'
              f'  export theDoc to POSIX file "{outdir}" as slide images with properties {{image format:PNG}}\n'
              '  close theDoc saving no\n'
              'end tell')
    try:
        subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
        pngs = sorted(glob.glob(os.path.join(outdir, '**', '*.png'), recursive=True))
        return pngs[0] if pngs else None
    except Exception as e:
        print('  Keynote render failed:', e)
        return None

def render_soffice(pptx, dpi=130):
    so = find_soffice()
    if not so:
        return None
    d = tempfile.mkdtemp(prefix='h2p_lo_')
    try:
        if shutil.which('pdftoppm'):                 # pptx->pdf->png (DPI control)
            subprocess.run([so, '--headless', '--convert-to', 'pdf', '--outdir', d, pptx],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
            pdf = glob.glob(os.path.join(d, '*.pdf'))[0]
            subprocess.run(['pdftoppm', '-png', '-r', str(dpi), pdf, os.path.join(d, 'p')],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
            return sorted(glob.glob(os.path.join(d, 'p*.png')))[0]
        subprocess.run([so, '--headless', '--convert-to', 'png', '--outdir', d, pptx],  # poppler-free fallback
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        pngs = glob.glob(os.path.join(d, '*.png'))
        return pngs[0] if pngs else None
    except Exception as e:
        print('  LibreOffice render failed:', e)
        return None

def render_pptx(pptx, dpi=130, renderer='auto'):
    named = {'powerpoint': [render_powerpoint], 'keynote': [render_keynote],
             'libreoffice': [render_soffice]}
    if renderer in named:
        chain = named[renderer]
    else:                                            # auto: native first, LibreOffice fallback
        s = platform.system()
        chain = ([render_powerpoint] if s == 'Windows' else
                 [render_keynote] if s == 'Darwin' else []) + [render_soffice]
    for fn in chain:
        png = fn(pptx, dpi)
        if png:
            return png
    return None

def open_in_office(pptx):
    """Open the deck in the OS default office app — the real, human verify."""
    p = os.path.abspath(pptx); s = platform.system()
    try:
        if s == 'Windows':
            os.startfile(p)                          # default handler = PowerPoint
        elif s == 'Darwin':
            subprocess.run(['open', p])
        elif shutil.which('xdg-open'):
            subprocess.run(['xdg-open', p])
        elif find_soffice():
            subprocess.Popen([find_soffice(), p])
        else:
            print('open it manually:', p); return False
        print('opened in your office app — eyeball it:', p)
        return True
    except Exception as e:
        print('could not open:', e); return False

# ---------- source HTML render (for the compare image) ----------
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
    ap.add_argument('--open', action='store_true', help='just open the deck in your office app')
    ap.add_argument('--renderer', default='auto', choices=['auto', 'powerpoint', 'keynote', 'libreoffice'])
    a = ap.parse_args()

    report_editable(a.pptx)
    if a.open:
        open_in_office(a.pptx); return

    png = render_pptx(a.pptx, renderer=a.renderer)
    if not png:
        print('no headless renderer available.\n'
              '  → On Windows/macOS just open it in your office app:  python verify.py "%s" --open\n'
              '  → or install LibreOffice for an automated cross-platform compare.' % a.pptx)
        return
    built = Image.open(png).convert('RGB')
    if not a.source:
        built.save(a.out); print('rendered ->', a.out); return
    src = Image.open(html_to_png(a.source, a.region, a.bg)).convert('RGB')
    h = 720
    def fit(im): return im.resize((int(im.width * h / im.height), h))
    a_im, b_im = fit(src), fit(built)
    gap = 24
    canvas = Image.new('RGB', (a_im.width + b_im.width + gap, h), (20, 20, 24))
    canvas.paste(a_im, (0, 0)); canvas.paste(b_im, (a_im.width + gap, 0))
    canvas.save(a.out)
    print(f'compare (left=source HTML, right=built PPTX) -> {a.out}')

if __name__ == '__main__':
    main()
