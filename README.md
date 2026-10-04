# Automation

## search_documents.py

Search every document in a folder for a word or term.

```bash
python search_documents.py <folder> "<term>" [options]
```

| Option | Meaning |
| --- | --- |
| `-i`, `--ignore-case` | Case-insensitive search |
| `-w`, `--whole-word` | Match whole words only |
| `-r`, `--recursive` | Include subfolders |
| `-e`, `--regex` | Treat the term as a regular expression |
| `-l`, `--files-only` | Only list the names of matching files |
| `--ext .docx .pdf` | Only search these file types |

Supported formats: plain text and code files (`.txt`, `.md`, `.csv`, `.json`, `.html`, ...),
Word (`.docx`), Excel (`.xlsx`), PowerPoint (`.pptx`), OpenDocument (`.odt`, `.ods`, `.odp`)
and PDF. Everything works with the standard library except PDF, which needs `pip install pypdf`.

Example:

```bash
python search_documents.py ~/Documents "invoice" -i -r
```

The script exits with code 0 if it found a match, 1 if nothing matched, and 2 on an error.
