# -*- mode: python -*-
import sys
from os import path
site_packages = next(p for p in sys.path if 'site-packages' in p)
block_cipher = None


a = Analysis(['SportsGrabber.py'],
             pathex=['/Users/scottdligon/Desktop/2018-2019/2014-Jan2018/Menlo Sports Statistics Updater/Menlo-College-Sports-Statistics-Updater'],
             binaries=[(path.join('/Library/Frameworks/Python.framework/Versions/2.7/lib/python2.7/site-packages',"docx","templates"), 
                "docx/templates")],
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
          exclude_binaries=True,
          name='SportsGrabber',
          debug=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='SportsGrabber')
