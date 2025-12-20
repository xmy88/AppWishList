"""Convenience builder for the bid document manager.

This script wraps PyInstaller to produce a GUI binary for macOS or Windows.
It picks sensible defaults and normalizes ``--add-data`` separators so a single
command works cross‑platform.
"""
from __future__ import annotations

import argparse
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path

DEFAULT_APP_NAME = "BidManager"
DEFAULT_ENTRYPOINT = Path(__file__).with_name("bid_manager.py")
DEFAULT_DATA = Path(__file__).parent.parent / "data"


class BuildError(RuntimeError):
    pass


def detect_pyinstaller() -> str:
    exe = shutil.which("pyinstaller")
    if not exe:
        raise BuildError("PyInstaller is not installed. Run `pip install pyinstaller`. ")
    return exe


def normalize_add_data(src: Path, dest: str | None = None) -> str:
    """Format ``--add-data`` for the current OS.

    PyInstaller expects ``src:dest`` on POSIX and ``src;dest`` on Windows.
    """
    if dest is None:
        dest = src.name
    sep = ";" if os.name == "nt" else ":"
    return f"{src}{sep}{dest}"


def build(args: argparse.Namespace) -> None:
    pyinstaller = detect_pyinstaller()

    if not DEFAULT_ENTRYPOINT.exists():
        raise BuildError(f"Entrypoint not found: {DEFAULT_ENTRYPOINT}")

    cmd = [
        pyinstaller,
        "--noconfirm",
        "--windowed",
        f"--name={args.name}",
    ]

    if args.clean:
        cmd.append("--clean")

    if args.onefile:
        cmd.append("--onefile")
    else:
        cmd.append("--onedir")

    if args.icon:
        cmd.append(f"--icon={args.icon}")

    add_data = list(args.add_data)
    if args.include_default_data and DEFAULT_DATA.exists():
        add_data.append(normalize_add_data(DEFAULT_DATA.resolve()))

    for item in add_data:
        cmd.extend(["--add-data", item])

    cmd.append(str(DEFAULT_ENTRYPOINT))

    print("Running:", " ".join(cmd))
    try:
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as exc:
        raise BuildError(f"PyInstaller exited with status {exc.returncode}") from exc

    dist = Path("dist") / args.name
    bundle_hint = f"{dist}\n"
    if os.name == "nt":
        bundle_hint += f"Executable: {dist / f'{args.name}.exe'}"
    elif platform.system() == "Darwin":
        bundle_hint += f"App bundle: {dist}.app"
    else:
        bundle_hint += "Check the dist folder for outputs."

    print("\nBuild complete! Output located at:\n" + bundle_hint)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build Bid Manager with PyInstaller")
    parser.add_argument("--name", default=DEFAULT_APP_NAME, help="Application name")
    parser.add_argument("--icon", help="Path to .ico/.icns icon")
    parser.add_argument(
        "--onefile",
        action="store_true",
        help="Create a single-file executable (slower startup)",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Remove PyInstaller cache before building",
    )
    parser.add_argument(
        "--no-default-data",
        dest="include_default_data",
        action="store_false",
        help="Skip bundling the repository data directory",
    )
    parser.add_argument(
        "--add-data",
        action="append",
        default=[],
        metavar="SRC{sep}DEST".format(sep=";" if os.name == "nt" else ":"),
        help=(
            "Extra --add-data entries. Use ':' on macOS/Linux or ';' on Windows. "
            "This flag can be repeated."
        ),
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        build(args)
    except BuildError as exc:
        print(f"Build failed: {exc}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
