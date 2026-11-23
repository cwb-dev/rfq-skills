import sys
from pathlib import Path

import pymupdf


def main():
    if len(sys.argv) < 4:
        raise SystemExit("usage: merge_drawings.py <contract.pdf> <output.pdf> <drawing.pdf> [...]")
    contract, output, *drawings = map(Path, sys.argv[1:])
    merged = pymupdf.open(contract)
    for drawing in drawings:
        source = pymupdf.open(drawing)
        merged.insert_pdf(source)
        source.close()
    merged.save(output)
    merged.close()
    print(f"merged {len(drawings)} drawing file(s) -> {output}")


if __name__ == "__main__":
    main()
