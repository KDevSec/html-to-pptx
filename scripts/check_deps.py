#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dependency doctor + auto-installer for html-to-pptx — run this FIRST, every time.

Default behaviour = AUTO-INSTALL whatever is missing (don't just nag the user).

Tiers:
  CORE (build the .pptx):  python-pptx · Pillow · playwright + its Chromium
  VERIFY (render-check):   the platform's OWN office app is the real verifier —
      Windows  PowerPoint (COM, via pip `pywin32`)
      macOS    Keynote / PowerPoint (AppleScript, built-in)
      Linux    LibreOffice (soffice)
    LibreOffice is only auto-installed when NO native office app exists; it's the
    cross-platform fallback renderer. pdftoppm/poppler is optional (soffice exports
    PNG itself; installed on Linux/macOS only).

System packages install via the detected manager (apt/dnf/pacman/zypper +sudo · brew
· winget); anything that can't auto-install prints the manual command.

  python check_deps.py            # auto-install missing, then report
  python check_deps.py --check    # report only, install nothing

Exit 0 once CORE is present; exit 1 if CORE is still missing.  Stdlib-only.
"""
import argparse, glob, importlib, importlib.util, os, platform, shutil, subprocess, sys

OK, BAD, WARN = '✓', '✗', '!'
WIN, MAC, LINUX = platform.system() == 'Windows', platform.system() == 'Darwin', platform.system() == 'Linux'

# ---------- detection ----------
def has_module(name):  # NEVER read playwright.__version__ (no such attribute)
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

def find_powerpoint():
    if not WIN:
        return None
    for ver in ('Office16', 'Office15', 'root\\Office16'):
        for base in (r'C:\Program Files\Microsoft Office', r'C:\Program Files (x86)\Microsoft Office'):
            p = os.path.join(base, ver, 'POWERPNT.EXE')
            if os.path.exists(p):
                return p
    return shutil.which('powerpnt')

def find_keynote():
    return '/Applications/Keynote.app' if (MAC and os.path.exists('/Applications/Keynote.app')) else None

def find_pp_mac():
    return '/Applications/Microsoft PowerPoint.app' if (MAC and os.path.exists('/Applications/Microsoft PowerPoint.app')) else None

def _pw_cache_dirs():
    h = os.path.expanduser('~')
    return [os.environ.get('PLAYWRIGHT_BROWSERS_PATH', ''),
            os.path.join(h, '.cache', 'ms-playwright'),
            os.path.join(h, 'Library', 'Caches', 'ms-playwright'),
            os.path.join(h, 'AppData', 'Local', 'ms-playwright')]

def has_chromium():
    if not has_module('playwright'):
        return None
    for d in _pw_cache_dirs():
        if d and (glob.glob(os.path.join(d, 'chromium-*')) or glob.glob(os.path.join(d, 'chromium_headless_shell-*'))):
            return True
    if (shutil.which('google-chrome') or shutil.which('chromium') or shutil.which('chromium-browser')
            or shutil.which('msedge') or os.path.exists(r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe')):
        return True
    return False

def verify_renderers():
    """List (name, ready) renderers usable for the automated render-check."""
    r = []
    if WIN and find_powerpoint():
        r.append(('PowerPoint (COM)', has_module('win32com')))   # ready needs pywin32
    if MAC and find_keynote():
        r.append(('Keynote', True))
    if MAC and find_pp_mac():
        r.append(('PowerPoint', True))
    if find_soffice():
        r.append(('LibreOffice', True))
    return r

# ---------- install helpers ----------
def run(cmd, why=''):
    print(f'  $ {" ".join(cmd)}' + (f'   # {why}' if why else ''))
    try:
        return subprocess.run(cmd).returncode == 0
    except Exception as e:
        print(f'    (could not run: {e})'); return False

def _sudo():
    if hasattr(os, 'geteuid') and os.geteuid() == 0:
        return []
    return ['sudo'] if shutil.which('sudo') else []

def pip_install(pkgs):
    base = [sys.executable, '-m', 'pip', 'install', '--user']
    return run(base + pkgs, 'pip') or run(base + ['--break-system-packages'] + pkgs, 'pip (PEP668 retry)')

def install_libreoffice():
    if LINUX:
        if shutil.which('apt-get'): return run(_sudo() + ['apt-get', 'install', '-y', 'libreoffice-impress'], 'LibreOffice')
        if shutil.which('dnf'):     return run(_sudo() + ['dnf', 'install', '-y', 'libreoffice-impress'], 'LibreOffice')
        if shutil.which('pacman'):  return run(_sudo() + ['pacman', '-S', '--noconfirm', 'libreoffice-still'], 'LibreOffice')
        if shutil.which('zypper'):  return run(_sudo() + ['zypper', '--non-interactive', 'install', 'libreoffice-impress'], 'LibreOffice')
    elif MAC and shutil.which('brew'):
        return run(['brew', 'install', '--cask', 'libreoffice'], 'LibreOffice')
    elif WIN:
        if shutil.which('winget'):  return run(['winget', 'install', '-e', '--id', 'TheDocumentFoundation.LibreOffice',
                                                '--accept-package-agreements', '--accept-source-agreements'], 'LibreOffice')
        if shutil.which('choco'):   return run(['choco', 'install', '-y', 'libreoffice-fresh'], 'LibreOffice')
    return False

def install_poppler():
    if LINUX:
        if shutil.which('apt-get'): return run(_sudo() + ['apt-get', 'install', '-y', 'poppler-utils'], 'poppler')
        if shutil.which('dnf'):     return run(_sudo() + ['dnf', 'install', '-y', 'poppler-utils'], 'poppler')
        if shutil.which('pacman'):  return run(_sudo() + ['pacman', '-S', '--noconfirm', 'poppler'], 'poppler')
    elif MAC and shutil.which('brew'):
        return run(['brew', 'install', 'poppler'], 'poppler')
    return False  # Windows: skip (optional; soffice/PowerPoint render without it)

def manual_hint():
    if LINUX:  return 'sudo apt-get install -y libreoffice-impress poppler-utils'
    if MAC:    return 'brew install --cask libreoffice   (or just use Keynote)'
    if WIN:    return 'use PowerPoint (pip install pywin32), or winget install -e --id TheDocumentFoundation.LibreOffice'
    return 'install LibreOffice, or use your platform office app'

# ---------- report ----------
def report():
    core = {'pptx': 'python-pptx', 'PIL': 'Pillow', 'playwright': 'playwright'}
    core_missing = []
    print('\nhtml-to-pptx dependency check'); print('=' * 42)
    print('CORE (build the .pptx):')
    for mod, pkg in core.items():
        ok = has_module(mod); print(f'  [{OK if ok else BAD}] python: {mod:11s} ({pkg})')
        if not ok: core_missing.append(pkg)
    ch = has_chromium(); m = OK if ch is True else (WARN if ch is None else BAD)
    print(f'  [{m}] playwright Chromium' + ('' if ch is True else ' (python -m playwright install chromium)'))
    if ch is False: core_missing.append('chromium')

    print('VERIFY (render-check; the office app you present in):')
    ren = verify_renderers()
    ready = [n for n, ok in ren if ok]
    for n, ok in ren:
        if ok: print(f'  [{OK}] {n}')
        else:  print(f'  [{WARN}] {n} present — needs pywin32 for auto-render (pip install pywin32); or open the file manually')
    if not ren:
        print(f'  [{WARN}] no office app detected — install LibreOffice, or just open the .pptx where you present')
    pt = shutil.which('pdftoppm')
    if find_soffice():
        print(f'  [{OK if pt else WARN}] pdftoppm (poppler) — OPTIONAL' + (f' @ {pt}' if pt else '; soffice exports PNG itself'))
    print('=' * 42)
    return core_missing, ready

# ---------- main ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true', help='report only, install nothing')
    ap.add_argument('--fix', action='store_true', help='(default) install missing deps')
    a = ap.parse_args()

    if not a.check:                                   # AUTO-INSTALL (default)
        core = {'pptx': 'python-pptx', 'PIL': 'Pillow', 'playwright': 'playwright'}
        miss = [pkg for mod, pkg in core.items() if not has_module(mod)]
        if miss or (has_module('playwright') and has_chromium() is not True) or not verify_renderers():
            print('installing missing dependencies (use --check to only report)...')
        if miss:
            pip_install(miss); importlib.invalidate_caches()
        if has_module('playwright') and has_chromium() is not True:
            run([sys.executable, '-m', 'playwright', 'install', 'chromium'], 'Chromium')
        # VERIFY: prefer the native office app; LibreOffice only if no native one exists
        if WIN:
            if find_powerpoint():
                if not has_module('win32com'):
                    pip_install(['pywin32'])          # enable PowerPoint COM render (no LibreOffice needed)
            elif not find_soffice():
                if not install_libreoffice():
                    print(f'  install a renderer manually: {manual_hint()}')
        elif MAC:
            if not (find_keynote() or find_pp_mac()) and not find_soffice():
                if not install_libreoffice():
                    print(f'  install a renderer manually: {manual_hint()}')
        else:  # Linux: LibreOffice is the native office app
            if not find_soffice() and not install_libreoffice():
                print(f'  install a renderer manually: {manual_hint()}')
            if not shutil.which('pdftoppm'):
                install_poppler()                     # optional, best-effort

    core_missing, ready = report()
    if core_missing:
        print(f'{BAD} CORE still missing: {", ".join(core_missing)}')
        print('  retry: python -m pip install --user --break-system-packages python-pptx Pillow playwright && python -m playwright install chromium')
        return 1
    if not ready:
        print(f'{OK} CORE ready — you can build. {WARN} no auto-render office app; verify by opening the .pptx where you present.')
        return 0
    print(f'{OK} all set — build ready; verify renderer: {ready[0]}.')
    return 0

if __name__ == '__main__':
    sys.exit(main())
