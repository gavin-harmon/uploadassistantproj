# -*- mode: python ; coding: utf-8 -*-

import sys
sys.setrecursionlimit(500000)


block_cipher = None


a = Analysis(['C:\\Users\\gharmo\\PycharmProjects\\uploadassistantproj\\uploadassistant\\main.py'],
             pathex=['C:\\Users\\gharmo\\PycharmProjects\\uploadassistantproj\\uploadassistant'],
             binaries=[],
             datas=[('c:\\Users\\gharmo\\PycharmProjects\\uploadassistantproj\\UAclipboard.ico','.'), ],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries - TOC([ ('arrow.dll', None, None), ('mkl_avx512.dll', None, None), ('libopenblas.NOIJJG62EMASZI6NYURL6JBKM4EVBGM7.gfortran-win_amd64.dll', None, None), ('icudt58.dll', None, None), ('libmmd.dll', None, None), ('mfc140u.dll', None, None), ('mkl_avx.dll', None, None), ('mkl_avx2.dll', None, None), ('mkl_core.dll', None, None), ('mkl_def.dll', None, None), ('mkl_mc.dll', None, None), ('mkl_mc3.dll', None, None), ('mkl_pgi_thread.dll', None, None), ('mkl_rt.dll', None, None), ('mkl_scalapack_ilp64.dll', None, None), ('mkl_scalapack_lp64.dll', None, None), ('mkl_sequential.dll', None, None), ('mkl_tbb_thread.dll', None, None), ('mkl_vml_avx.dll', None, None), ('mkl_vml_avx2.dll', None, None), ('mkl_vml_avx512.dll', None, None), ('mkl_vml_cmpt.dll', None, None), ('mkl_vml_def.dll', None, None), ('', None, None), ('mkl_vml_mc.dll', None, None), ('mkl_vml_mc2.dll', None, None), ('mkl_vml_mc3.dll', None, None), ('opengl32sw.dll', None, None), ('Qt5Core.dll', None, None), ('Qt5Gui.dll', None, None), ('Qt5Widgets.dll', None, None) ,('arrow_flight.dll', None, None), ('arrow_python.dll', None, None), ('hdf5.dll', None, None), ('iconv.dll', None, None), ('icuin58.dll', None, None), ('icuuc58.dll', None, None), ('libcrypto-1_1-x64.dll', None, None), ('libifcoremd.dll', None, None), ('libiomp5md.dll', None, None), ('libprotobuf.dll', None, None), ('libxml2.dll', None, None), ('parquet.dll', None, None), ('Qt5Network.dll', None, None), ('Qt5Qml.dll', None, None), ('Qt5Quick.dll', None, None), ('sqlite3.dll', None, None), ('tcl86t.dll', None, None), ('tk86t.dll', None, None), ('ucrtbase.dll', None, None), ('winpty.dll', None, None)     ]),
          a.zipfiles,
          a.datas,
          [],
          name='Upload Assistant',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
		  icon='c:\\Users\\gharmo\\PycharmProjects\\DataCollectionAssistant\\UAclipboard.ico')
