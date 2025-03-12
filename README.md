# Create hyperlinks from citations to bibliographies for Zotero

This Python script can add hyperlinks from Zotero's citations to bibliographies, so you can jump from citation to its bibliography like normal cross-reference.

## How does it work?

Thanks to [gwyn-hopkins](https://forums.zotero.org/discussion/comment/418013/#Comment_418013)'s VBA script, I implement the same function with Python.

- This script scans all bibliographies and adds unique bookmarks for each one.
- This script scans all citations and sets corresponding hyperlink to the bookmark.
- This script can also set color of the citation for you :).

## How to use?

**This script can only work in Windows.**

**See [issue#1](https://github.com/Syize/link-zotero-citation-bibliography/issues/1).**

### Download the script

Only [link-zotero-citation-bibliography.py](link-zotero-citation-bibliography.py) is needed. Files end with `.pyi` is the file I use to trick PyCharm (because I update this code under Linux).

You can clone this repo or just download [link-zotero-citation-bibliography.py](link-zotero-citation-bibliography.py).

### Install dependencies

Use your favorite Python package manager to install:

- pywin32
- rich

### Modify and run

1. Change the `word_file_path` and `new_file_path` in script, an absolute path is recommended.
2. Run the script.
