import sys
from pathlib import Path


project = Path(SPECPATH)
analysis = Analysis(
    [str(project / "main.py")],
    pathex=[str(project)],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pandas", "numpy"],
    noarchive=False,
)
archive = PYZ(analysis.pure)
single_file = sys.platform == "win32"
executable = EXE(
    archive,
    analysis.scripts,
    analysis.binaries if single_file else [],
    analysis.datas if single_file else [],
    exclude_binaries=not single_file,
    name="ExcelConversion",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
if not single_file:
    distribution = COLLECT(
        executable,
        analysis.binaries,
        analysis.datas,
        strip=False,
        upx=False,
        name="ExcelConversion",
    )
if sys.platform == "darwin":
    application = BUNDLE(
        distribution,
        name="ExcelConversion.app",
        icon=None,
        bundle_identifier="com.excelconversion.desktop",
        info_plist={"NSHighResolutionCapable": True},
    )