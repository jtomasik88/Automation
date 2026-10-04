#!/usr/bin/env python3
"""SuperSearch: search every document in the SuperSearch Sources folder for a word or term.

Double-click this file (or run it with no arguments) and it will ask what to search for.
By default it searches the folder in DEFAULT_FOLDER below, including subfolders,
ignoring upper/lower case.

Supported formats:
  - Plain text: .txt .md .csv .tsv .log .json .xml .html .htm .yaml .yml .ini .cfg
                .py .js .ts .java .c .cpp .h .cs .go .rb .sh .sql .rtf (and more)
  - Word:        .docx      (built in, no extra packages)
  - PowerPoint:  .pptx      (built in)
  - Excel:       .xlsx      (built in)
  - OpenDocument: .odt .ods .odp (built in)
  - PDF:         .pdf       (requires `pip install pypdf`)

Examples:
  python SuperSearch.py                      (asks for the search term)
  python SuperSearch.py "invoice"
  python SuperSearch.py "net profit" --whole-word
  python SuperSearch.py "Invoice" --case-sensitive --ext .docx .pdf
  python SuperSearch.py "colou?r" --regex
  python SuperSearch.py "invoice" --folder "C:\\Some\\Other Folder"
"""

import argparse
import html
import re
import sys
import zipfile
from pathlib import Path

# The folder SuperSearch looks through unless --folder is given.
DEFAULT_FOLDER = Path(r"C:\Users\jtomasik\OneDrive - BlueScope\Desktop\THE MACHINE\SuperSearch Sources")

TEXT_EXTENSIONS = {
    ".txt", ".md", ".markdown", ".csv", ".tsv", ".log", ".json", ".xml",
    ".html", ".htm", ".yaml", ".yml", ".ini", ".cfg", ".conf", ".rtf",
    ".py", ".js", ".ts", ".java", ".c", ".cpp", ".h", ".hpp", ".cs", ".go",
    ".rb", ".php", ".sh", ".bat", ".ps1", ".sql", ".tex", ".rst",
}
OFFICE_EXTENSIONS = {".docx", ".pptx", ".xlsx", ".odt", ".ods", ".odp"}
PDF_EXTENSIONS = {".pdf"}
SUPPORTED_EXTENSIONS = TEXT_EXTENSIONS | OFFICE_EXTENSIONS | PDF_EXTENSIONS

# XML elements that mark a paragraph/row/slide break inside Office documents.
_BREAK_TAGS = re.compile(
    r"</(?:w:p|a:p|text:p|text:h|si|c|table:table-cell)>|<w:br/>|<w:tab/>|<text:line-break/>"
)
_XML_TAG = re.compile(r"<[^>]+>")


class ExtractionError(Exception):
    pass


def _xml_to_text(xml: str) -> str:
    xml = _BREAK_TAGS.sub("\n", xml)
    return html.unescape(_XML_TAG.sub("", xml))


def read_text_file(path: Path) -> str:
    for encoding in ("utf-8", "utf-16", "cp1252", "latin-1"):
        try:
            return path.read_text(encoding=encoding)
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ExtractionError("could not decode text")


def read_office_file(path: Path) -> str:
    ext = path.suffix.lower()
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            if ext == ".docx":
                parts = [n for n in names if re.fullmatch(
                    r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml", n)]
            elif ext == ".pptx":
                parts = sorted(
                    (n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                    key=lambda n: int(re.search(r"\d+", n.rsplit("/", 1)[1]).group()),
                )
            elif ext == ".xlsx":
                parts = [n for n in names if n == "xl/sharedStrings.xml"
                         or re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)]
            else:  # OpenDocument
                parts = ["content.xml"] if "content.xml" in names else []
            return "\n".join(
                _xml_to_text(zf.read(p).decode("utf-8", errors="replace")) for p in parts
            )
    except zipfile.BadZipFile as exc:
        raise ExtractionError("not a valid Office file") from exc


def read_pdf_file(path: Path) -> str:
    try:
        from pypdf import PdfReader
    except ImportError:
        raise ExtractionError("PDF support needs `pip install pypdf`")
    try:
        reader = PdfReader(str(path))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as exc:  # pypdf raises many different error types
        raise ExtractionError(f"could not read PDF ({exc})") from exc


def extract_text(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in OFFICE_EXTENSIONS:
        return read_office_file(path)
    if ext in PDF_EXTENSIONS:
        return read_pdf_file(path)
    return read_text_file(path)


def build_pattern(term: str, use_regex: bool, whole_word: bool, ignore_case: bool) -> re.Pattern:
    expr = term if use_regex else re.escape(term)
    if whole_word:
        expr = rf"\b(?:{expr})\b"
    return re.compile(expr, re.IGNORECASE if ignore_case else 0)


def iter_files(folder: Path, recursive: bool, extensions: set):
    candidates = folder.rglob("*") if recursive else folder.glob("*")
    for path in sorted(candidates):
        if path.is_file() and path.suffix.lower() in extensions and not path.name.startswith("~$"):
            yield path


def run_search(folder: Path, term: str, args) -> int:
    try:
        pattern = build_pattern(term, args.regex, args.whole_word, not args.case_sensitive)
    except re.error as exc:
        print(f"Error: invalid regular expression: {exc}", file=sys.stderr)
        return 2

    extensions = ({e.lower() if e.startswith(".") else f".{e.lower()}" for e in args.ext}
                  if args.ext else SUPPORTED_EXTENSIONS)

    files_searched = files_matched = total_matches = 0
    for path in iter_files(folder, not args.no_subfolders, extensions):
        name = path.relative_to(folder)
        try:
            text = extract_text(path)
        except (ExtractionError, OSError) as exc:
            print(f"[skipped] {name}: {exc}", file=sys.stderr)
            continue
        files_searched += 1

        hits = [(n, line.strip()) for n, line in enumerate(text.splitlines(), 1)
                if pattern.search(line)]
        if not hits:
            continue

        count = sum(len(pattern.findall(line)) for _, line in hits)
        files_matched += 1
        total_matches += count
        if args.files_only:
            print(name)
            continue
        print(f"\n{name}  ({count} match{'es' if count != 1 else ''})")
        for n, line in hits:
            snippet = line if len(line) <= 200 else line[:197] + "..."
            print(f"  line {n}: {snippet}")

    print(f"\nSearched {files_searched} file(s): '{term}' found {total_matches} time(s) "
          f"in {files_matched} file(s).")
    return 0 if files_matched else 1


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Search every document in the SuperSearch Sources folder for a word or term.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Examples:" + __doc__.split("Examples:", 1)[1],
    )
    parser.add_argument("term", nargs="?",
                        help="word or phrase to look for (leave out to be asked)")
    parser.add_argument("-f", "--folder", type=Path, default=DEFAULT_FOLDER,
                        help="folder to search (default: SuperSearch Sources)")
    parser.add_argument("-c", "--case-sensitive", action="store_true",
                        help="match upper/lower case exactly")
    parser.add_argument("-w", "--whole-word", action="store_true", help="match whole words only")
    parser.add_argument("-n", "--no-subfolders", action="store_true",
                        help="only search the top folder, not its subfolders")
    parser.add_argument("-e", "--regex", action="store_true", help="treat TERM as a regular expression")
    parser.add_argument("-l", "--files-only", action="store_true", help="only list matching file names")
    parser.add_argument("--ext", nargs="+", metavar="EXT",
                        help="only search these extensions, e.g. --ext .docx .pdf")
    args = parser.parse_args()

    folder = args.folder.expanduser()
    interactive = args.term is None
    if not folder.is_dir():
        print(f"Error: '{folder}' is not a folder.", file=sys.stderr)
        if interactive:
            input("\nPress Enter to close...")
        return 2

    if not interactive:
        return run_search(folder, args.term, args)

    # No term given (e.g. the file was double-clicked): keep asking until a blank entry.
    print(f"SuperSearch - searching: {folder}")
    while True:
        try:
            term = input("\nSearch for (press Enter on a blank line to quit): ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            return 0
        if not term:
            return 0
        run_search(folder, term, args)


if __name__ == "__main__":
    sys.exit(main())
