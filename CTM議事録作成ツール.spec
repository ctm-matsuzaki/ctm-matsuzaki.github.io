# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


ROOT = Path.cwd()
APP_NAME = "議事録作成ツール"
EXECUTABLE_NAME = "minutes_tool_app"
ICON_PATH = ROOT / "minutes_tool" / "resources" / "icon.icns"

a = Analysis(
    ["run_minutes_tool.py"],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[
        (
            str(ROOT / "minutes_tool" / "resources" / "会議議事録テンプレ.xlsx"),
            "minutes_tool/resources",
        ),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=EXECUTABLE_NAME,
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
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=EXECUTABLE_NAME,
)

app = BUNDLE(
    coll,
    name=f"{APP_NAME}.app",
    icon=str(ICON_PATH),
    bundle_identifier="jp.ctm.minutes-tool",
    info_plist={
        "CFBundleName": APP_NAME,
        "CFBundleDisplayName": APP_NAME,
        "CFBundleShortVersionString": "1.0.0",
        "CFBundleVersion": "1.0.0",
        "NSHighResolutionCapable": True,
    },
)
