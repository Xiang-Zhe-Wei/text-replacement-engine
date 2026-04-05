import re
import sys
from pathlib import Path
from replacer import replace_in_docx, replace_multiple_in_docx

REPLACEMENTS_MD = Path(__file__).parent / "replacements.md"


def load_replacements(md_path: Path) -> list[tuple[str, str]]:
    """Parse replacements.md and return a list of (target, replacement) tuples.
    Rows marked with '需依文意選擇' are skipped.
    """
    replacements = []
    for line in md_path.read_text(encoding="utf-8").splitlines():
        if not line.startswith("|"):
            continue
        parts = [p.strip() for p in line.strip("|").split("|")]
        if len(parts) < 3:
            continue
        no, target, replacement = parts[0], parts[1], parts[2]
        if not no.isdigit():
            continue
        if "需依文意選擇" in replacement:
            continue
        replacements.append((target, replacement))
    return replacements


def main():
    if len(sys.argv) == 2:
        input_file = sys.argv[1]
        replacements = load_replacements(REPLACEMENTS_MD)
        if not replacements:
            print("Error: No replacements found in replacements.md")
            sys.exit(1)
        try:
            output_path, counts = replace_multiple_in_docx(input_file, replacements)
            total = sum(counts.values())
            for target, count in counts.items():
                if count > 0:
                    replacement = next(r for t, r in replacements if t == target)
                    print(f"  '{target}' → '{replacement}': {count} 處")
            print(f"\n共替換 {total} 處，輸出至：{output_path}")
        except (FileNotFoundError, ValueError) as e:
            print(f"Error: {e}")
            sys.exit(1)

    elif len(sys.argv) == 4:
        input_file = sys.argv[1]
        target = sys.argv[2]
        replacement = sys.argv[3]
        try:
            output_path, count = replace_in_docx(input_file, target, replacement)
            print(f"Replaced {count} occurrence(s) of '{target}' with '{replacement}'")
            print(f"Output saved to: {output_path}")
        except (FileNotFoundError, ValueError) as e:
            print(f"Error: {e}")
            sys.exit(1)

    else:
        print("Usage:")
        print("  python main.py <input_file>                        # 套用 replacements.md 所有規則")
        print("  python main.py <input_file> <target> <replacement> # 單一替換")
        sys.exit(1)


if __name__ == "__main__":
    main()
