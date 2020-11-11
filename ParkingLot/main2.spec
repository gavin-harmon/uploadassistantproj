import sys
sys.setrecursionlimit(500000)


# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['C:\\Users\\gharmo\\PycharmProjects\\DataCollectionAssistant\\DataCollectionAssistant\\main.py'],
             pathex=['C:\\Users\\gharmo\\PycharmProjects\\DataCollectionAssistant\\DataCollectionAssistant\\'],
             binaries =[],
             datas=[('c:\\Users\\gharmo\\PycharmProjects\\DataCollectionAssistant\\UAclipboard.ico','.'), ],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],		 			 
			 excludes=[],
             win_no_prefer_redirects=True,
             win_private_assemblies=True,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
		  a.binaries,
          a.zipfiles,
          a.datas,
          name='DataCollectionAssistant',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=False,
		  clean=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
		  icon='c:\\Users\\gharmo\\PycharmProjects\\DataCollectionAssistant\\UAclipboard.ico')