from pathlib import Path
from docx import Document


def replace_in_docx(input_path: str, target: str, replacement: str) -> str:
    """
    Replace all occurrences of target with replacement in a .docx file.
    Returns the path of the newly created output file.
    """
    path = Path(input_path)

    if not path.exists():
        raise FileNotFoundError(f"File not found: {input_path}")
    if path.suffix.lower() != ".docx":
        raise ValueError(f"Not a .docx file: {input_path}")

    doc = Document(input_path)
    count = 0

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if target in run.text:
                count += run.text.count(target)
                run.text = run.text.replace(target, replacement)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if target in run.text:
                            count += run.text.count(target)
                            run.text = run.text.replace(target, replacement)

    output_path = path.parent / f"{path.stem}_modified{path.suffix}"
    doc.save(output_path)

    return str(output_path), count
