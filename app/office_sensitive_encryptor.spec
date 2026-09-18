# -*- mode: python ; coding: utf-8 -*-
# 文档敏感词加解密工具 - PyInstaller 打包配置
# 用法: 由 build.bat 调用（一条龙），或手动:
#   pyinstaller office_sensitive_encryptor.spec --noconfirm --distpath msix_packaging\app
# 产物: msix_packaging/app/OfficeSensitiveEncryptor.exe （单文件）
#
# 可选: 把 icon.ico 放在本 spec 同目录并去掉下面 icon 行的注释

a = Analysis(
    ['office_sensitive_encryptor.py'],
    pathex=[],
    binaries=[],
    datas=[('icon.ico', '.'),],   # icon.ico 同时作为数据文件，供运行时设置窗口图标
    hiddenimports=[
        # cryptography 的 C 扩展后端，显式声明防止个别环境漏收集
        'cryptography.hazmat.backends.openssl.backend',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 剔除用不到的重量级模块，缩小体积
        'unittest', 'pydoc', 'doctest', 'tkinter.test',
    ],
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
    name='OfficeSensitiveEncryptor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # GUI 程序：不弹出黑色控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',      # exe 文件图标（icon.ico 已内置项目根目录）
)
