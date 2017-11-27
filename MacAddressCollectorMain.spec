# -*- mode: python -*-

block_cipher = None


a = Analysis(['MacAddressCollectorMain.py'],
             pathex=['C:\\Python\\Python35-32\\Lib\\site-packages\\pyttsx3\\drivers', 'F:\\Workspace\\GitHub\\MacAdressCollector'],
             binaries=[],
             datas=[],
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
          name='MacAddressCollectorMain',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True , icon='110.ico')
