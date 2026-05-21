"""cx_Freeze build script — produces a per-user .msi installer."""

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
    ],
}

# Never change these GUIDs — Windows uses them to recognise upgrades as the
# same product rather than installing a second copy.
UPGRADE_CODE        = "{4E7C1A0F-32B5-4D8E-9A6C-B3F27E84D510}"
DESKTOP_COMP_GUID   = "{B3C2D1E0-F4A5-6789-BCDE-F01234567890}"
STARTMENU_COMP_GUID = "{C4D3E2F1-05B6-789A-CDEF-012345678901}"

bdist_msi_options = {
    "all_users": False,                                     # per-user, no admin needed
    "initial_target_dir": r"[LocalAppDataFolder]\PRCe360ServiceListConverter",
    "upgrade_code": UPGRADE_CODE,
    "add_to_path": False,
    # Shortcuts need their own components so they're cleanly removed on uninstall.
    "shortcuts": [
        {
            "shortcut_id": "DesktopShortcut",
            "directory":   "DesktopFolder",
            "name":        "PRCe360 Service List Converter",
            "component":   "DesktopShortcutComp",
            "target":      "[TARGETDIR]PRCe360ServiceListConverter.exe",
            "wk_dir":      "TARGETDIR",
        },
        {
            "shortcut_id": "StartMenuShortcut",
            "directory":   "ProgramMenuFolder",
            "name":        "PRCe360 Service List Converter",
            "component":   "StartMenuShortcutComp",
            "target":      "[TARGETDIR]PRCe360ServiceListConverter.exe",
            "wk_dir":      "TARGETDIR",
        },
    ],
    "data": {
        "Component": [
            # (Component, ComponentId, Directory_, Attributes, Condition, KeyPath)
            ("DesktopShortcutComp",  DESKTOP_COMP_GUID,   "TARGETDIR", 0, None, None),
            ("StartMenuShortcutComp", STARTMENU_COMP_GUID, "TARGETDIR", 0, None, None),
        ],
        "FeatureComponents": [
            ("MainFeature", "DesktopShortcutComp"),
            ("MainFeature", "StartMenuShortcutComp"),
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
            base="Win32GUI",
            target_name="PRCe360ServiceListConverter.exe",
        ),
    ],
)
