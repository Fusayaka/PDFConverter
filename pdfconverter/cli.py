import os
import argparse
import win32com.client


SUPPORTED_EXTENSIONS = (".docx", ".pptx")
WORD_PDF_FORMAT = 17
POWERPOINT_PDF_FORMAT = 32


def collect_file_paths(input_paths):
    """Recursively collect .docx and .pptx file paths from files or directories."""
    file_paths = []

    for path in input_paths:
        if os.path.isfile(path):
            if path.lower().endswith(SUPPORTED_EXTENSIONS):
                file_paths.append(path)
            else:
                print(f"Skipped (unsupported format): {path}")

        elif os.path.isdir(path):
            for root, _, files in os.walk(path):
                for file in files:
                    if file.lower().endswith(SUPPORTED_EXTENSIONS):
                        file_paths.append(os.path.join(root, file))

        else:
            print(f"Skipped (not found): {path}")

    return file_paths


def init_com_applications(file_paths):
    """Initialize and return Word and PowerPoint COM application instances as needed."""
    word = None
    powerpoint = None

    if any(p.lower().endswith(".docx") for p in file_paths):
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

    if any(p.lower().endswith(".pptx") for p in file_paths):
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = False

    return word, powerpoint


def resolve_output_path(input_path, output_dir=None):
    """Return the output PDF path, using output_dir if provided, else same dir as input."""
    base_name = os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        return os.path.join(output_dir, base_name)
    return os.path.splitext(os.path.abspath(input_path))[0] + ".pdf"


def convert_to_pdf(input_path, word=None, powerpoint=None, output_dir=None):
    """Convert a single .docx or .pptx file to PDF. Returns True on success."""
    abs_input = os.path.abspath(input_path)
    abs_output = resolve_output_path(input_path, output_dir)
    ext = os.path.splitext(input_path)[1].lower()

    if ext == ".docx" and word:
        doc = word.Documents.Open(abs_input)
        doc.SaveAs(abs_output, FileFormat=WORD_PDF_FORMAT)
        doc.Close()

    elif ext == ".pptx" and powerpoint:
        presentation = powerpoint.Presentations.Open(abs_input)
        presentation.SaveAs(abs_output, POWERPOINT_PDF_FORMAT)
        presentation.Close()

    else:
        return False

    print(f"Converted: {abs_output}")
    return True


def convert_files(file_paths, output_dir=None):
    """Convert a batch of .docx and .pptx files to PDF using COM automation."""
    word, powerpoint = init_com_applications(file_paths)

    try:
        for file_path in file_paths:
            try:
                convert_to_pdf(file_path, word, powerpoint, output_dir)
            except Exception as e:
                print(f"Error converting {file_path}: {e}")
    finally:
        if word:
            word.Quit()
        if powerpoint:
            powerpoint.Quit()


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Batch convert Word (.docx) and PowerPoint (.pptx) files to PDF "
            "using Microsoft Office COM automation. Requires Word and/or "
            "PowerPoint to be installed on Windows."
        )
    )
    parser.add_argument(
        "paths",
        nargs="+",
        metavar="PATH",
        help=(
            "One or more file or folder paths to process. "
            "Files must be .docx or .pptx; other formats are skipped. "
            "Folders are scanned recursively for supported files. "
        ),
    )
    args = parser.parse_args()

    file_paths = collect_file_paths(args.paths)

    if not file_paths:
        print("No supported .docx or .pptx files found in the provided paths.")
        return

    print(f"Found {len(file_paths)} file(s) to convert.")
    convert_files(file_paths, args.output)


if __name__ == "__main__":
    main()