"""cx_Freeze build script — produces a machine-wide .msi installer."""

from cx_Freeze import setup, Executable

build_options = {
    "packages": ["openpyxl", "docx", "lxml", "PIL"],
    "includes": [
        "openpyxl.cell._writer",
        "lxml.etree",
        "lxml._elementpath",
        "PIL.Image",
        "PIL.ImageTk",
    ],
    "include_files": [
        ("brand/NMPRC_logo-1.png", "brand/NMPRC_logo-1.png"),
        ("GridTemplate.xlsx", "GridTemplate.xlsx"),
    ],
}

# Never change this GUID — Windows uses it to recognise upgrades as the
# same product rather than installing a second copy.
UPGRADE_CODE = "{4E7C1A0F-32B5-4D8E-9A6C-B3F27E84D510}"

SHORTCUT_COMP_GUID = "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}"

bdist_msi_options = {
    "all_users": True,          # machine-wide (required for SYSTEM/Automox deployment)
    "initial_target_dir": r"[ProgramFilesFolder]PRCe360ServiceListConverter",
    "upgrade_code": UPGRADE_CODE,
    "add_to_path": False,
    "data": {
        "Component": [
            # Shortcut-only component; KeyPath=None means the component directory is the key.
            ("ShortcutComp", SHORTCUT_COMP_GUID, "TARGETDIR", 0, None, None),
        ],
        "FeatureComponents": [
            ("default", "ShortcutComp"),
        ],
        "Shortcut": [
            # (Shortcut, Directory_, Name, Component_, Target, Arguments,
            #  Description, Hotkey, Icon_, IconIndex, ShowCmd, WkDir)
            ("DesktopShortcut", "DesktopFolder",
             "PRCe360 Service List Converter", "ShortcutComp",
             "[TARGETDIR]PRCe360ServiceListConverter.exe",
             None, None, None, None, None, None, "TARGETDIR"),
            ("StartMenuShortcut", "ProgramMenuFolder",
             "PRCe360 Service List Converter", "ShortcutComp",
             "[TARGETDIR]PRCe360ServiceListConverter.exe",
             None, None, None, None, None, None, "TARGETDIR"),
        ],
        "RemoveFile": [
            # Clean up shortcuts on uninstall.
            ("RemoveShortcuts", "ShortcutComp", None, "TARGETDIR", 2),
        ],
    },
}

setup(
    name="PRCe360 Service List Converter",
    version="1.0.0",
    description="Converts telecom service lists to GridTemplate format",
    options={
        "build_exe": build_options,
        "bdist_msi": bdist_msi_options,
    },
    executables=[
        Executable(
            "gui.py",
            base="gui",
            target_name="PRCe360ServiceListConverter.exe",
            icon="brand/NMPRC_logo.ico",
        ),
    ],
)
