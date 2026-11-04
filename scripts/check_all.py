import fnmatch
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ASSETS = HERE.parent / "assets"
EXTENSIONS = {"deck": "pptx", "sop": "docx"}
PATTERNS = {"html": "*report*.html", "sop": "*sop-sample*.docx"}
SKIP_NAMES = {"docx": "*sop*.docx"}
BOARD = HERE.parents[2]


def data_assets():
    if not (BOARD / "board_check.py").is_file():
        return None
    sys.path.insert(0, str(BOARD))
    import board_check
    return board_check.data_path("") / HERE.parent.relative_to(BOARD) / "v1" / "assets"


def main():
    assets = Path(sys.argv[1]) if len(sys.argv) > 1 else ASSETS
    if len(sys.argv) > 1 and not assets.is_dir():
        print(f"SKIP: data not on this device: {assets.as_posix()}")
        return 0
    checkers = sorted(p for p in HERE.glob("check_*.py") if p.name != "check_all.py")
    if not checkers:
        print("SKIP: no scripts/check_*.py yet")
        return 0
    failed = 0
    for checker in checkers:
        kind = checker.stem[len("check_"):]
        ext = EXTENSIONS.get(kind, kind)
        pattern = PATTERNS.get(kind, f"*sample*.{ext}")
        skip = SKIP_NAMES.get(kind)

        def find(folder):
            return sorted(p for p in folder.glob(pattern) if not (skip and fnmatch.fnmatch(p.name, skip)))

        samples = find(assets) if assets.is_dir() else []
        where = assets
        if not samples and len(sys.argv) == 1:
            other = data_assets()
            if other is not None and other.is_dir():
                samples = find(other)
                where = other
        if not samples:
            print(f"SKIP {checker.name}: no {pattern} in {assets.name}/")
            continue
        for sample in samples:
            run = subprocess.run([sys.executable, "-B", str(checker), str(sample)], capture_output=True, text=True)
            status = "PASS" if run.returncode == 0 else "FAIL"
            print(f"{status} {checker.name} on {where.name}/{sample.name}" + (" (data root)" if where != assets else ""))
            for line in run.stdout.splitlines():
                if line.startswith("FAIL") or line.startswith("SKIP"):
                    print("  " + line)
            if run.returncode != 0:
                failed += 1
                if run.stderr.strip():
                    print("  " + run.stderr.strip().splitlines()[-1])
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
