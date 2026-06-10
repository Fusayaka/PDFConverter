import os
import argparse
import win32com.client


def collect_filepaths(path):
    filepaths = []

    if os.path.isfile(path):
        filepaths.append(path)

    elif os.path.isdir(path):
        for root, _, fs in os.walk(path):
            for f in fs:
                if f.lower().endswith((".docx", ".pptx")):
                    filepaths.append(os.path.join(root, f))

    return filepaths


def convert_batch(paths):
    has_docx = any(p.lower().endswith(".docx") for p in paths)
    has_pptx = any(p.lower().endswith(".pptx") for p in paths)

    word = None
    powerpoint = None

    if has_docx:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

    if has_pptx:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1

    for input_path in paths:
        ext = input_path.split('.')[-1].lower()
        output_path = input_path.rsplit('.', 1)[0] + ".pdf"

        try:
            if ext == "docx":
                doc = word.Documents.Open(os.path.abspath(input_path))
                doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
                doc.Close()

            elif ext == "pptx":
                presentation = powerpoint.Presentations.Open(os.path.abspath(input_path))
                presentation.SaveAs(os.path.abspath(output_path), 32)
                presentation.Close()

            else:
                print(f"Skip: {input_path}")
                continue

            print(f"Converted: {output_path}")

        except Exception as e:
            print(f"Error with {input_path}: {e}")

    if word:
        word.Quit()
    if powerpoint:
        powerpoint.Quit()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("paths", nargs="+", help="File(s) or folder(s)")
    args = parser.parse_args()

    all_files = []
    for p in args.paths:
        all_files.extend(collect_filepaths(p))

    if not all_files:
        print("No valid files found")
        return

    convert_batch(all_files)


if __name__ == "__main__":
    main()