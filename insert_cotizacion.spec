# -*- mode: python ; coding: utf-8 -*-

import os
import platform
import sys


def _target_triple() -> str:
    machine = platform.machine().lower()
    arch = "x86_64" if machine in ("x86_64", "amd64") else machine
    if sys.platform.startswith("linux"):
        return f"{arch}-unknown-linux-gnu"
    if sys.platform == "darwin":
        return f"{arch}-apple-darwin"
    if sys.platform == "win32":
        return f"{arch}-pc-windows-msvc"
    raise RuntimeError(f"Plataforma no soportada: {sys.platform}")


a = Analysis(
    ['backend/insert_wrapper.py'],
    pathex=[os.path.abspath('backend')],
    binaries=[],
    datas=[],
    hiddenimports=['insert_cotizacion', 'extract_cotizacion', 'models', 'xlsx_manager'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=f'insert_cotizacion-{_target_triple()}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
