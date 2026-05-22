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

bdist_msi_options = {
    "all_users": True,          # machine-wide (required for SYSTEM/Automox deployment)
    "initial_target_dir": r"[ProgramFilesFolder]PRCe360ServiceListConverter",
    "upgrade_code": UPGRADE_CODE,
    "add_to_path": False,
    # Shortcuts are created by the Automox remediation script instead.
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
        ),
    ],
)
