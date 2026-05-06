# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_dynamic_libs

datas = [('icon.ico', '.'), ('icon_status_on.ico', '.'), ('icon_status_off.ico', '.')]
binaries = []
datas += collect_data_files('pycaw')
binaries += collect_dynamic_libs('comtypes')


block_cipher = None


a = Analysis(['VCVM.py'],
             pathex=[],
             binaries=binaries,
             datas=datas,
             hiddenimports=['pystray', 'pystray._win32', 'PIL', 'PIL.Image', 'PIL.ImageDraw', 'pycaw', 'pycaw.pycaw', 'comtypes', 'comtypes.client'],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=['numpy', 'pandas', 'scipy', 'matplotlib', 'tkinter'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,  
          [],
          name='VCVM',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None , icon='icon.ico')
