# -*- mode: python ; coding: utf-8 -*-
import sys, shutil

a = Analysis(
    ['C:\\Users\\Toan\\WordToExcel\\gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    a.binaries + [('msvcp100.dll', 'C:\\Windows\\System32\\msvcp100.dll', 'BINARY'),
            ('msvcr100.dll', 'C:\\Windows\\System32\\msvcr100.dll', 'BINARY')],
    a.zipfiles, a.datas,
    exclude_binaries=True,
    name=os.path.join('dist', 'WordToExcel'  + ('.exe' if sys.platform == 'win32' else '')),
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['C:\\Users\\Toan\\WordToExcel\\Images\\logo.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='WordToExcel',
)
shutil.copyfile('Images/logo.png', '{0}/WordToExcel/logo.png'.format(DISTPATH))
