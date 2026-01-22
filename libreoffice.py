import subprocess
import os
import shutil
from pathlib import Path
from typing import Union

def resolve_libreoffice_path() -> Path:
    """
    Resolve the LibreOffice executable path across Linux/macOS installs.

    Precedence:
    1) LIBREOFFICE_BIN / LIBREOFFICE_PATH env var
    2) PATH lookup for 'libreoffice' or 'soffice'
    3) macOS default app bundle path
    """
    env_override = os.getenv("LIBREOFFICE_BIN") or os.getenv("LIBREOFFICE_PATH")
    if env_override:
        candidate = Path(env_override)
        if candidate.exists():
            return candidate

    for name in ("libreoffice", "soffice"):
        found = shutil.which(name)
        if found:
            return Path(found)

    mac_default = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if mac_default.exists():
        return mac_default

    mac_alt = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice-bin")
    if mac_alt.exists():
        return mac_alt

    raise FileNotFoundError(
        "LibreOffice executable not found. "
        "Install LibreOffice or set LIBREOFFICE_BIN to the full path."
    )


def convert_pptx_to_pdf(input_path: Union[str, os.PathLike], output_dir: Union[str, os.PathLike]) -> Path:
    """
    Convert a PPT/PPTX file to PDF using LibreOffice in headless mode.

    Args:
        input_path: Path to the input PPT/PPTX file.
        output_dir: Directory where the resulting PDF will be written.

    Returns:
        Path to the generated PDF file.
    """
    input_path = Path(input_path)
    output_dir = Path(output_dir)

    if not output_dir.exists():
        output_dir.mkdir(parents=True, exist_ok=True)

    libreoffice_path = resolve_libreoffice_path()
    command = [
        str(libreoffice_path),
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--nolockcheck",
        "--norestore",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(input_path),
    ]

    # Run conversion; raise CalledProcessError if it fails
    subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # LibreOffice writes <stem>.pdf into output_dir
    pdf_path = output_dir / (input_path.stem + '.pdf')
    if not pdf_path.exists():
        raise FileNotFoundError(f"Expected PDF not found: {pdf_path}")

    return pdf_path


if __name__ == '__main__':
    # Example usage guarded for direct execution only. Update paths as needed.
    example_input = '/path/to/input.pptx'
    example_outdir = './outputs'
    try:
        result = convert_pptx_to_pdf(example_input, example_outdir)
        print(f'Converted {example_input} to PDF at {result}')
    except Exception as exc:
        print(f'Conversion failed: {exc}')
