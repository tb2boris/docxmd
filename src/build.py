"""Сборка docxmd.exe. Ядро DP152 и данные natasha кладутся внутрь exe."""

from __future__ import annotations

import importlib
import subprocess
import sys
from pathlib import Path

SRC = Path(__file__).resolve().parent
DP152_SRC = SRC.parents[1] / "dp152" / "src"


def _pkg_data(package: str, subpath: str) -> str:
    pkg = importlib.import_module(package)
    src = Path(pkg.__file__).parent / subpath
    return f"{src};{package}/{subpath}"


def main() -> None:
    if not DP152_SRC.is_dir():
        raise SystemExit(f"Не найден каталог DP152: {DP152_SRC}")
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--windowed",
        "--name",
        "docxmd",
        "--paths",
        str(SRC),
        "--paths",
        str(DP152_SRC),
        "--add-data",
        f"{DP152_SRC / 'config'};config",
        "--add-data",
        _pkg_data("natasha", "data"),
        "--hidden-import",
        "core.processor",
        "--hidden-import",
        "core.detector",
        "--hidden-import",
        "core.generator",
        "--hidden-import",
        "core.mapper",
        "--hidden-import",
        "core.md_shield",
        "--hidden-import",
        "openpyxl",
        "--collect-submodules",
        "openpyxl",
        "--hidden-import",
        "fitz",
        "--collect-all",
        "pymupdf",
        "--exclude-module",
        "torch",
        "--exclude-module",
        "torchvision",
        "--exclude-module",
        "tensorflow",
        "--exclude-module",
        "pandas",
        "--exclude-module",
        "scipy",
        "--exclude-module",
        "sklearn",
        "--exclude-module",
        "sympy",
        "--exclude-module",
        "pytest",
        "--exclude-module",
        "IPython",
        "--exclude-module",
        "matplotlib",
        "--exclude-module",
        "notebook",
        "--noconfirm",
        "--clean",
        "gui.py",
    ]
    print(" ".join(cmd))
    subprocess.run(cmd, check=True, cwd=SRC)
    print(f"Сборка завершена: {SRC / 'dist' / 'docxmd.exe'}")


if __name__ == "__main__":
    main()
