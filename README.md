# Create hyperlinks from citations to bibliographies for Zotero

<p align="center">English | <a href="README_CN.md">中文文档</a></p>

This Python script can add hyperlinks from Zotero's citations to bibliographies, so you can jump from citation to its bibliography like normal cross-reference.

## How does it work?

- This script scans all bibliographies and adds unique bookmarks for each one.
- This script scans all citations and sets corresponding hyperlink to the bookmark.
- This script can also set font color and underline style of the citation for you :).

## Important Note

- **This script can only work in Windows.**
- **Numbered citation format hasn't been tested.**

## How to use?

1. Clone this repo.
2. Install following dependencies:
   - pywin32
   - pyzotero
   - rich
3. Modify the `main.py`:
   - `zotero_id` and `zotero_api_key` will be used to connect to Zotero. Check the documentation of [pyzotero](https://pyzotero.readthedocs.io/en/latest/index.html) to know how to get your ID and API key.
   - `word_file_path` is the absolute path of your Word document.
   - `new_file_path` is the absolute save path of the new Word document.
4. Run `main.py`.
