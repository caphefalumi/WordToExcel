# -*- mode: python ; coding: utf-8 -*-
import shutil

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
            ('msvcr100.dll', 'C:\\Windows\\System32\\msvcr100.dll', 'BINARY'), 
            ('vcruntime140_1.dll', 'C:\Windows\System32\vcruntime140_1.dll', 'BINARY'),
            ('vcruntime140.dll', 'C:\Windows\System32\vcruntime140.dll', 'BINARY'),
            ('pywintypes311.dll', 'C:\Windows\System32\pywintypes311.dll', 'BINARY'),
            ('pythoncom311.dll', 'C:\Windows\System32\pythoncom311.dll', 'BINARY')],
    a.zipfiles, a.datas,
    exclude_binaries=True,
    name='WordToExcel',
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
shutil.rmtree('build')
#pyinstaller -noconfirm  --onedir --windowed --icon="C:\Users\Toan\WordToExcel\Images\logo.ico"  --upx-dir "C:\Users\Toan\upx-4.1.0-win64\" "C:\Users\Toan\WordToExcel\gui.py" -n "WordToExcel" --noconfirm
#pyinstaller  --upx-dir "C:\Users\Toan\upx-4.1.0-win64\" "C:\Users\Toan\WordToExcel\WordToExcel.spec" --noconfirm   