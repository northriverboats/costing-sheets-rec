# -*- mode: python -*-

block_cipher = None


a = Analysis(['main.py'],
             pathex=['C:\\Development\\costing-sheets-rec'],
             binaries=[],
             datas=[('sheets.ico','.'),('.env','.'),('CostingSheetTemplate.xlsx','.')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='Costing Sheets for Rec',
          debug=False,
          strip=False,
          upx=True,
          console=True , icon='sheets.ico')