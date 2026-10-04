# Automation

## SuperSearch.py

Searches every document in

```
C:\Users\jtomasik\OneDrive - BlueScope\Desktop\THE MACHINE\SuperSearch Sources
```

(including subfolders) for a word or term. Upper/lower case is ignored by default.

**Easiest way:** double-click `SuperSearch.py`, type what you're looking for, and press Enter.
Keep searching as many times as you like; press Enter on a blank line to close.

**From Command Prompt:**

```
py SuperSearch.py "invoice"
```

| Option | Meaning |
| --- | --- |
| `-c`, `--case-sensitive` | Match upper/lower case exactly |
| `-w`, `--whole-word` | Match whole words only |
| `-n`, `--no-subfolders` | Only search the top folder |
| `-e`, `--regex` | Treat the term as a regular expression |
| `-l`, `--files-only` | Only list the names of matching files |
| `--ext .docx .pdf` | Only search these file types |
| `-f "C:\Other Folder"` | Search a different folder this time |

To change the folder permanently, edit the `DEFAULT_FOLDER` line near the top of the script.

Supported formats: plain text and code files (`.txt`, `.md`, `.csv`, `.json`, `.html`, ...),
Word (`.docx`), Excel (`.xlsx`), PowerPoint (`.pptx`), OpenDocument (`.odt`, `.ods`, `.odp`)
and PDF. Everything works with the standard library except PDF, which needs `py -m pip install pypdf`.
