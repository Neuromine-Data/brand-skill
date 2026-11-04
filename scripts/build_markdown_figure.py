import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.dont_write_bytecode = True
sys.path.insert(0, str(HERE))
import build_word

FIGURE_NAME = "neuromine-sample-figure-1.png"


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parent / "assets"
    out.mkdir(parents=True, exist_ok=True)
    path = out / FIGURE_NAME
    path.write_bytes(build_word.figure_png())
    print(f"wrote {path.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
