# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['CreditSafe.py'],
             pathex=[],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=['astor', 'bcrypt', 'boto3', 'botocore', 'cryptography', 'dask', 'docutils', 'google', 'h5py', 'IPython', 'jinja2', 'markdown', 'nacl', 'nbconvert', 'nbformat', 'notebook', 'pycparser', 'pywintypes', 'sklearn', 'tensorflow', 'h5py', 'pyinstaller-4.7.dist-info', 'IPython', 'pytz', 'altgraph-0.17.2.dist-info', 'bcrypt', 'scipy', 'boto3', 'setuptools-41.0.1.dist-info', 'botocore', 'Markdown-3.1.dist-info', 'share', 'markupsafe', 'sklearn', 'cryptography', 'matplotlib', 'tcl', 'cryptography-2.6.1.dist-info', 'msgpack', 'tcl8', 'dask', 'nacl', 'tensorflow', 'docutils', 'nbconvert', 'tk', 'etc', 'nbconvert-5.6.1.dist-info', 'tornado', 'google', 'nbformat', 'win32com', 'google_api_core-1.14.2.dist-info', 'notebook', 'winpty', 'google_cloud_core-0.29.1.dist-info', 'numpy', 'zmq', 'google_cloud_storage-1.14.0.dist-info', 'pandas', 'grpc', 'psutil'],
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
          name='CreditSafe',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None )
