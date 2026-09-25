"""Portable PyInstaller builds; no global pip upgrades or user-cache deletion."""
import argparse
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--kind", choices=["gui", "cli", "all"], default="gui")
    parser.add_argument("--output", default=str(ROOT / "dist"))
    args = parser.parse_args()
    output = Path(args.output).resolve()
    work = ROOT / ".artifacts" / "pyinstaller"
    specs = ROOT / ".artifacts" / "specs"
    specs.mkdir(parents=True, exist_ok=True)
    common = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
              "--icon", str(ROOT / "assets" / "app.ico"),
              "--additional-hooks-dir", str(ROOT / "build"),
              "--version-file", str(ROOT / "build" / "version_info.txt"),
              "--distpath", str(output), "--workpath", str(work), "--specpath", str(specs)]
    if args.kind in ("gui", "all"):
        for name, onefile in [("GalBGMController", False), ("GalBGMController-single", True)]:
            command = common + ["--name", name, "--noconsole", "--add-data", f"{ROOT / 'assets'};assets"]
            if onefile:
                command.append("--onefile")
            subprocess.run(command + [str(ROOT / "app_gui.py")], check=True, cwd=ROOT)
    if args.kind in ("cli", "all"):
        subprocess.run(common + ["--name", "GalBGMController-CLI", "--onefile", "--console",
                                 str(ROOT / "app_cli.py")], check=True, cwd=ROOT)
    print(f"Build succeeded: {output}")


if __name__ == "__main__":
    main()
