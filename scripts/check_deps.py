#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dependency doctor for html-to-pptx — run this FIRST on every use.

Two tiers:
  CORE (needed to build the .pptx):  python-pptx · Pillow · playwright · Chromium
  VERIFY (needed only for the render-compare QA step):  LibreOffice (soffice)
         · pdftoppm is OPTIONAL — soffice can export PNG on its own.

Stdlib-only, so it runs even when nothing is installed.

  python check_deps.py          # report (exit 0 = CORE present)
  python check_deps.py --fix    # auto-install CORE pip deps + Chromium; print the
                                # OS command for the system packages (LibreOffice…)

Exit 0 = all CORE deps present (build can run). Exit 1 = a CORE dep is missing.
Missing VERIFY deps only warn (you can still build; the QA compare is skipped).
"""
import argparse, glob, importlib.util, os, platform, shutil, subprocess, sys

OK, BAD, WARN = '✓', '✗', '!'   # ✓ ✗ !

def has_module(name):
    # NOTE: use find_spec, never `playwright.__version__` (that attribute doesn't exist).
    return importlib.util.find_spec(name) is not None

def find_soffice():
    cands = sorted(glob.glob('/opt/libreoffice*/program/soffice'), reverse=True)
    cands += ['/Applications/LibreOffice.app/Contents/MacOS/soffice',
              r'C:\Program Files\LibreOffice\program\soffice.exe',
              r'C:\Program Files (x86)\LibreOffice\program\soffice.exe']
    for c in cands:
        if os.path.exists(c):
            return c
    return shutil.which('soffice') or shutil.which('libreoffice')

def playwright_cache_dirs():
    home = os.path.expanduser('~')
    return [os.environ.get('PLAYWRIGHT_BROWSERS_PATH', ''),
            os.path.join(home, '.cache', 'ms-playwright'),              # Linux
            os.path.join(home, 'Library', 'Caches', 'ms-playwright'),   # macOS
            os.path.join(home, 'AppData', 'Local', 'ms-playwright')]    # Windows

def has_chromium():
    if not has_module('playwright'):
        return None  # can't tell until playwright itself is present
    for d in playwright_cache_dirs():
        if d and (glob.glob(os.path.join(d, 'chromium-*')) or
                  glob.glob(os.path.join(d, 'chromium_headless_shell-*'))):
            return True
    if shutil.which('google-chrome') or shutil.which('chromium') or shutil.which('chromium-browser') or \
       shutil.which('msedge') or os.path.exists(r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'):
        return True   # a system Chromium/Edge/Chrome also works for capture
    return False

def soffice_hint():
    s = platform.system()
    if s == 'Linux':   return 'sudo apt-get install -y libreoffice-impress   # (or your distro pkg manager)'
    if s == 'Darwin':  return 'brew install --cask libreoffice'
    if s == 'Windows': return 'winget install TheDocumentFoundation.LibreOffice   # or installer from libreoffice.org'
    return 'install LibreOffice via your OS package manager'

def poppler_hint():
    s = platform.system()
    if s == 'Linux':   return 'sudo apt-get install -y poppler-utils      # optional (higher-res compare)'
    if s == 'Darwin':  return 'brew install poppler                        # optional (higher-res compare)'
    if s == 'Windows': return 'choco install -y poppler                    # optional; skip it — LibreOffice exports PNG itself'
    return 'install poppler (pdftoppm) — optional'

def pip_install(pkgs):
    base = [sys.executable, '-m', 'pip', 'install', '--user']
    if subprocess.run(base + pkgs).returncode != 0:     # PEP 668 -> retry
        print('  pip retry with --break-system-packages')
        return subprocess.run(base + ['--break-system-packages'] + pkgs).returncode == 0
    return True

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--fix', action='store_true', help='auto-install CORE pip deps + Chromium')
    a = ap.parse_args()

    core_py = {'pptx': 'python-pptx', 'PIL': 'Pillow', 'playwright': 'playwright'}

    if a.fix:
        miss = [pkg for mod, pkg in core_py.items() if not has_module(mod)]
        if miss:
            print(f'fixing: pip install {" ".join(miss)}'); pip_install(miss)
        if has_module('playwright') and has_chromium() is not True:
            print('fixing: playwright install chromium')
            subprocess.run([sys.executable, '-m', 'playwright', 'install', 'chromium'])

    print('\nhtml-to-pptx dependency check')
    print('=' * 38)
    core_missing, verify_missing = [], []

    print('CORE (build the .pptx):')
    for mod, pkg in core_py.items():
        ok = has_module(mod)
        print(f'  [{OK if ok else BAD}] python: {mod:11s} ({pkg})')
        if not ok: core_missing.append(pkg)
    chrome = has_chromium()
    cmark = OK if chrome is True else (WARN if chrome is None else BAD)
    print(f'  [{cmark}] playwright Chromium' + ('' if chrome is True else ' → python -m playwright install chromium'))
    if chrome is False: core_missing.append('chromium')

    print('VERIFY (render-compare QA step; skip if you only need the file):')
    so = find_soffice()
    print(f'  [{OK if so else BAD}] LibreOffice (soffice)' + (f' @ {so}' if so else ''))
    if not so: verify_missing.append('soffice')
    pt = shutil.which('pdftoppm')
    print(f'  [{OK if pt else WARN}] pdftoppm (poppler) — OPTIONAL' +
          (f' @ {pt}' if pt else '; LibreOffice exports PNG on its own'))

    print('=' * 38)
    if core_missing:
        print(f'{BAD} CORE incomplete — cannot build yet. Fix:')
        pips = [p for p in core_missing if p != 'chromium']
        if pips: print(f'  python {os.path.basename(__file__)} --fix')
        if 'chromium' in core_missing: print('  python -m playwright install chromium')
        return 1

    if verify_missing:
        print(f'{OK} CORE ready — you can build.  {WARN} VERIFY step needs LibreOffice:')
        print(f'  {soffice_hint()}')
        print(f'  (pdftoppm optional: {poppler_hint()})')
        return 0

    if not shutil.which('pdftoppm'):
        print(f'{OK} all set — build + verify ready (pdftoppm absent is fine; soffice exports PNG).')
    else:
        print(f'{OK} all dependencies present.')
    return 0

if __name__ == '__main__':
    sys.exit(main())
