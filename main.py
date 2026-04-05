import sys
from replacer import replace_in_docx


def main():
    if len(sys.argv) != 4:
        print("Usage: python main.py <input_file> <target> <replacement>")
        print("Example: python main.py test.docx 的 之")
        sys.exit(1)

    input_file = sys.argv[1]
    target = sys.argv[2]
    replacement = sys.argv[3]

    try:
        output_path, count = replace_in_docx(input_file, target, replacement)
        print(f"Replaced {count} occurrence(s) of '{target}' with '{replacement}'")
        print(f"Output saved to: {output_path}")
    except FileNotFoundError as e:
        print(f"Error: {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
